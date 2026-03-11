' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.MarkupReview.vb
' Purpose: Implements Markup Review — a DocCheck mode that compares a tracked-
'          changes document against acceptability constraints (provided as Word
'          comments in a separate "playbook" document) and evaluates whether
'          each revision is within acceptable bounds, including cross-clause
'          undermining detection.
'
' Architecture:
'  - Parameter Collection: Gathers cutoff date, constraint document path,
'    and output options through the standard ShowCustomVariableInputForm dialog.
'  - Constraint Loading: Opens the constraint .docx read-only/invisible,
'    extracts comments grouped by anchor text, and closes it immediately.
'  - Constraint-Driven Extraction: For each constraint, finds the anchor text
'    in the live document using FindLongTextAnchoredFast (searchOriginal mode),
'    then reads the original and marked-up text from the same Range object
'    in Original and Final views respectively.
'  - Pass 1 — Compliance: Per-constraint LLM calls comparing original vs
'    marked-up clause text against the constraint boundaries.
'  - Pass 2 — Cross-Clause (optional): Per-constraint LLM calls with full
'    marked-up contract text to detect undermining clauses elsewhere.
'  - Context Management: Uses truncation or chunking when the marked-up
'    contract text exceeds the configured context window.
'  - Output: Inserts Word comment bubbles on the marked-up document using
'    the pre-resolved Range objects from the extraction phase or produces a
'    report when bubble output is disabled or unavailable.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides Markup Review orchestration as part of the DocCheck feature set.
''' </summary>
Partial Public Class ThisAddIn

    ' ========================= Data Types =========================

    ''' <summary>
    ''' Represents one acceptability constraint extracted from the playbook document.
    ''' Multiple comments on the same anchor range are merged into a single entry.
    ''' </summary>
    Private Class MarkupConstraint
        Public Property Id As System.Int32
        Public Property AnchorText As System.String
        Public Property Instructions As System.Collections.Generic.List(Of System.String)
        Public Property AnchorStart As System.Int32
        Public Property MatchedClauseIndex As System.Int32 = -1
    End Class

    ''' <summary>
    ''' Represents one aligned clause pair (original text vs marked-up text).
    ''' </summary>
    Private Class ClausePair
        Public Property Index As System.Int32
        Public Property Heading As System.String
        Public Property OriginalText As System.String
        Public Property MarkedUpText As System.String
        Public Property DocStartPos As System.Int32
        Public Property DocEndPos As System.Int32
    End Class

    ''' <summary>
    ''' Holds the result of a single Pass 1 compliance check.
    ''' </summary>
    Private Class ComplianceResult
        Public Property ConstraintId As System.Int32
        Public Property ClauseIndex As System.Int32
        Public Property IsAcceptable As System.Boolean
        Public Property Reasoning As System.String
        Public Property CompromiseRedraft As System.String
        Public Property RawLLMResponse As System.String
    End Class

    ''' <summary>
    ''' Holds the result of a single Pass 2 cross-clause check.
    ''' </summary>
    Private Class CrossClauseResult
        Public Property ConstraintId As System.Int32
        Public Property UnderminingFound As System.Boolean
        Public Property Details As System.String
        Public Property RawLLMResponse As System.String
    End Class

    ' ========================= Main Entry Point =========================

    ''' <summary>
    ''' Runs the Markup Review workflow for the active document.
    ''' </summary>
    Private Async Function RunMarkupReview() As System.Threading.Tasks.Task
        Dim do2ndModel As System.Boolean = False
        Dim sw As Stopwatch = Stopwatch.StartNew()

        Try
            ' ── 0) Validate environment ──
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No open document.")
                Return
            End If
            Dim activeDoc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If activeDoc Is Nothing Then
                ShowCustomMessageBox("Active document was not found.")
                Return
            End If

            If activeDoc.Revisions Is Nothing OrElse activeDoc.Revisions.Count = 0 Then
                ShowCustomMessageBox("The active document has no tracked changes (revisions). Markup Review requires a document with tracked changes.")
                Return
            End If

            ' ── 1) Collect revision dates for the cutoff picker ──
            Dim revDates As New System.Collections.Generic.SortedSet(Of System.DateTime)()
            For Each rev As Microsoft.Office.Interop.Word.Revision In activeDoc.Revisions
                Try
                    revDates.Add(rev.Date.Date)
                Catch
                End Try
            Next

            Dim dateOptions As New System.Collections.Generic.List(Of System.String)()
            dateOptions.Add("(All revisions)")
            Dim now As System.DateTime = System.DateTime.Now
            For Each d As System.DateTime In revDates.Reverse()
                Dim daysAgo As System.Int32 = System.Math.Max(0, CInt((now.Date - d.Date).TotalDays))
                dateOptions.Add($"{d:yyyy-MM-dd} ({daysAgo} days ago)")
            Next

            ' ── 2) Build playbook file selection options ──
            Dim DocCheckPath As System.String = ExpandEnvironmentVariables(INI_DocCheckPath)
            If Not System.String.IsNullOrEmpty(DocCheckPath) AndAlso Not DocCheckPath.EndsWith("\", System.StringComparison.Ordinal) Then
                DocCheckPath &= "\"
            End If
            Dim DocCheckPathLocal As System.String = ExpandEnvironmentVariables(INI_DocCheckPathLocal)
            If Not System.String.IsNullOrEmpty(DocCheckPathLocal) AndAlso Not DocCheckPathLocal.EndsWith("\", System.StringComparison.Ordinal) Then
                DocCheckPathLocal &= "\"
            End If

            Dim constraintDisplayToPath As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
            Dim constraintOptions As New System.Collections.Generic.List(Of System.String)()

            ' First option: browse/drag-drop
            Dim browseOption As System.String = "(Browse or drag & drop a file...)"
            constraintOptions.Add(browseOption)
            constraintDisplayToPath.Add(browseOption, "")

            ' Enumerate *.docx playbook files from configured paths
            Dim dcPaths As New System.Collections.Generic.List(Of (p As System.String, isLocal As System.Boolean))()
            If Not System.String.IsNullOrWhiteSpace(DocCheckPath) AndAlso System.IO.Directory.Exists(DocCheckPath) Then
                dcPaths.Add((DocCheckPath, False))
            End If
            If Not System.String.IsNullOrWhiteSpace(DocCheckPathLocal) AndAlso System.IO.Directory.Exists(DocCheckPathLocal) Then
                dcPaths.Add((DocCheckPathLocal, True))
            End If

            For Each tuple In dcPaths
                Dim basePath As System.String = tuple.p
                Dim isLocal As System.Boolean = tuple.isLocal
                Try
                    If System.IO.Directory.Exists(basePath) Then
                        Dim files As System.String() = System.IO.Directory.GetFiles(basePath, $"{AN2}-dc-*.docx", System.IO.SearchOption.TopDirectoryOnly)
                        For Each f As System.String In files
                            Dim disp As System.String = System.IO.Path.GetFileName(f)
                            If isLocal Then disp &= " (local)"
                            If Not constraintDisplayToPath.ContainsKey(disp) Then
                                constraintDisplayToPath.Add(disp, f)
                                constraintOptions.Add(disp)
                            End If
                        Next
                    End If
                Catch
                End Try
            Next

            ' ── 3) Parameter form ──

            ' Load persisted Markup Review settings.
            ' MR_HasSaved distinguishes "user never ran Markup Review" (use code defaults
            ' that depend on INI values) from "user saved at least once" (trust stored values).
            Dim savedCrossClause As System.Boolean = True
            Dim savedBubbles As System.Boolean = True
            Dim savedMarkdownBubbles As System.Boolean = CBool(INI_MarkdownBubbles)
            Dim savedTruncation As System.Boolean = True
            Dim savedMaxChars As System.Int32 = 80000
            Dim savedUseSecond As System.Boolean = False
            Dim savedLanguage As System.String = INI_Language1
            Dim savedOther As System.String = ""

            If My.Settings.MR_HasSaved Then
                savedCrossClause = My.Settings.MR_CrossClause
                savedBubbles = My.Settings.MR_Bubbles
                savedMarkdownBubbles = My.Settings.MR_MarkdownBubbles
                savedTruncation = My.Settings.MR_Truncation
                savedMaxChars = My.Settings.MR_MaxChars
                If savedMaxChars < 5000 Then savedMaxChars = 80000
                savedUseSecond = My.Settings.MR_UseSecondModel
                Dim lang As System.String = My.Settings.MR_Language
                If Not System.String.IsNullOrWhiteSpace(lang) Then savedLanguage = lang
                savedOther = If(My.Settings.MR_OtherInstructions, "")
            End If

            Dim pCutoff As New SLib.InputParameter("Revision cutoff date", "(All revisions)")
            pCutoff.Options = New System.Collections.Generic.List(Of System.String)(dateOptions)

            Dim pConstraintDoc As New SLib.InputParameter("Constraint playbook", If(constraintOptions.Count > 1, constraintOptions(1), browseOption))
            pConstraintDoc.Options = New System.Collections.Generic.List(Of System.String)(constraintOptions)

            Dim pCrossClause As New SLib.InputParameter("Cross-clause undermining check", savedCrossClause)
            Dim pBubbles As New SLib.InputParameter("Output as Word bubbles", savedBubbles)
            Dim pMarkdownBubbles As New SLib.InputParameter("Markdown formatting in bubbles", savedMarkdownBubbles)
            Dim pTruncation As New SLib.InputParameter("Truncate (fast) for cross-check (vs. check all)", savedTruncation)
            Dim pMaxChars As New SLib.InputParameter("Cross-check context window (chars)", savedMaxChars)

            do2ndModel = savedUseSecond
            Dim p2nd As SLib.InputParameter
            If Not System.String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                p2nd = New SLib.InputParameter("Use a secondary model", do2ndModel)
            ElseIf INI_SecondAPI Then
                p2nd = New SLib.InputParameter("Use the secondary model", do2ndModel)
            Else
                p2nd = New SLib.InputParameter("Use the secondary model", do2ndModel)
            End If

            Dim pLanguage As New SLib.InputParameter("Language of output", savedLanguage)
            Dim pOther As New SLib.InputParameter("Other instructions", savedOther)

            Dim params() As SLib.InputParameter = {pCutoff, pConstraintDoc, pCrossClause, pBubbles, pMarkdownBubbles, pTruncation, pMaxChars, p2nd, pLanguage, pOther}

            If ShowCustomVariableInputForm("Please set the Markup Review parameters:", AN & " Markup Review", params) = False Then
                Return
            End If

            ' ── 4) Read back values ──
            Dim cutoffChoice As System.String = System.Convert.ToString(params(0).Value)
            Dim constraintChoice As System.String = System.Convert.ToString(params(1).Value)

            Dim doCrossClause As System.Boolean = False
            If TypeOf params(2).Value Is System.Boolean Then doCrossClause = CBool(params(2).Value)

            Dim doBubbles As System.Boolean = False
            If TypeOf params(3).Value Is System.Boolean Then doBubbles = CBool(params(3).Value)

            Dim doMarkdownBubbles As System.Boolean = False
            If TypeOf params(4).Value Is System.Boolean Then doMarkdownBubbles = CBool(params(4).Value)

            Dim doAutoTruncate As System.Boolean = False
            If TypeOf params(5).Value Is System.Boolean Then doAutoTruncate = CBool(params(5).Value)

            Dim maxChars As System.Int32 = 80000
            If TypeOf params(6).Value Is System.Int32 Then
                maxChars = CInt(params(6).Value)
                If maxChars < 5000 Then maxChars = 5000
            End If

            Dim useSecond = params(7).Value
            If TypeOf useSecond Is System.Boolean Then do2ndModel = CBool(useSecond) Else do2ndModel = False

            OutputLanguage = System.Convert.ToString(params(8).Value)
            OtherPrompt = System.Convert.ToString(params(9).Value)

            ' Persist Markup Review settings
            My.Settings.MR_HasSaved = True
            My.Settings.MR_CrossClause = doCrossClause
            My.Settings.MR_Bubbles = doBubbles
            My.Settings.MR_MarkdownBubbles = doMarkdownBubbles
            My.Settings.MR_Truncation = doAutoTruncate
            My.Settings.MR_MaxChars = maxChars
            My.Settings.MR_UseSecondModel = do2ndModel
            My.Settings.MR_Language = If(OutputLanguage, "")
            My.Settings.MR_OtherInstructions = If(OtherPrompt, "")
            My.Settings.Save()

            ' ── 4a) Resolve constraint document path ──
            Dim constraintDocPath As System.String = ""

            ' Check if user picked a known file from the dropdown
            If Not System.String.IsNullOrWhiteSpace(constraintChoice) Then
                constraintDisplayToPath.TryGetValue(constraintChoice, constraintDocPath)
            End If

            ' If browse option or empty → show drag-drop dialog
            If System.String.IsNullOrWhiteSpace(constraintDocPath) Then
                DragDropFormLabel = "Drag and drop the constraint (playbook) document (.docx) with acceptability comments, or click Browse."
                DragDropFormFilter = "Word Documents (*.doc;*.docx)|*.doc;*.docx|All Files (*.*)|*.*"
                constraintDocPath = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""
            End If

            If System.String.IsNullOrWhiteSpace(constraintDocPath) Then
                ShowCustomMessageBox("No constraint document was provided — will abort.")
                Return
            End If
            If Not System.IO.File.Exists(constraintDocPath) Then
                ShowCustomMessageBox("The constraint document was not found: " & constraintDocPath)
                Return
            End If

            ' ── 4b) Parse cutoff date ──
            Dim cutoffDate As System.DateTime? = Nothing
            If cutoffChoice IsNot Nothing AndAlso Not cutoffChoice.StartsWith("(", System.StringComparison.Ordinal) Then
                Dim datePart As System.String = cutoffChoice.Split(" "c)(0)
                Dim parsed As System.DateTime
                If System.DateTime.TryParse(datePart, parsed) Then
                    cutoffDate = parsed.Date
                End If
            End If

            ' ── 5) Secondary model ──
            If do2ndModel Then
                If Not System.String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded — aborting.")
                        Return
                    End If
                End If
            End If

            ' ── 6) Load constraints from playbook ──
            ShowProgressBarInSeparateThread(AN & " Markup Review", "Preparing...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressLabel = "Loading constraints from playbook..."

            Dim constraints As System.Collections.Generic.List(Of MarkupConstraint) = LoadConstraintsFromDocx(app, constraintDocPath)

            If constraints Is Nothing OrElse constraints.Count = 0 Then
                ProgressBarModule.CancelOperation = True
                ShowCustomMessageBox("No comments (constraints) were found in the constraint document.")
                Return
            End If

            If ProgressBarModule.CancelOperation Then Return

            ' ── 7) Resolve constraint regions (constraint-driven extraction) ──
            GlobalProgressLabel = "Locating constraint targets in document..."
            Dim clausePairs As New System.Collections.Generic.List(Of ClausePair)()
            Dim clauseRanges As New System.Collections.Generic.Dictionary(Of System.Int32, Microsoft.Office.Interop.Word.Range)()
            ResolveConstraintRegions(app, activeDoc, constraints, clausePairs, clauseRanges)

            Dim unmatchedCount As System.Int32 = 0
            For Each c As MarkupConstraint In constraints
                If c.MatchedClauseIndex < 0 Then unmatchedCount += 1
            Next
            If unmatchedCount > 0 Then
                Dim proceed As System.Int32 = ShowCustomYesNoBox(
                    $"{unmatchedCount} constraint(s) could not be located in the document. " &
                    "Unmatched constraints will be skipped. Do you want to continue?",
                    "Yes, continue", "No, abort")
                If proceed <> 1 Then
                    ProgressBarModule.CancelOperation = True
                    Return
                End If
            End If

            If ProgressBarModule.CancelOperation Then Return

            ' ── 8) PASS 1 — Compliance check ──
            Dim matchedConstraints As New System.Collections.Generic.List(Of MarkupConstraint)()
            For Each c As MarkupConstraint In constraints
                If c.MatchedClauseIndex >= 0 Then matchedConstraints.Add(c)
            Next

            Dim totalSteps As System.Int32 = matchedConstraints.Count + If(doCrossClause, matchedConstraints.Count, 0) + 1
            GlobalProgressMax = totalSteps
            GlobalProgressValue = 0

            Dim complianceResults As New System.Collections.Generic.List(Of ComplianceResult)()
            Dim idx As System.Int32 = 0

            For Each constraint As MarkupConstraint In matchedConstraints
                If ProgressBarModule.CancelOperation Then Exit For

                idx += 1
                GlobalProgressValue = idx
                GlobalProgressLabel = $"Pass 1: Checking constraint {idx} of {matchedConstraints.Count}..."

                Dim cp As ClausePair = clausePairs(constraint.MatchedClauseIndex)
                Dim result As ComplianceResult = Await RunComplianceCheck(constraint, cp, do2ndModel)
                If result IsNot Nothing Then complianceResults.Add(result)
            Next

            If ProgressBarModule.CancelOperation Then Return

            ' ── 9) PASS 2 — Cross-clause undermining (optional) ──
            Dim crossResults As New System.Collections.Generic.List(Of CrossClauseResult)()
            Dim truncatedConstraintIds As New System.Collections.Generic.HashSet(Of System.Int32)()
            Dim protectedZoneExceededIds As New System.Collections.Generic.HashSet(Of System.Int32)()
            Dim chunkedConstraintIds As New System.Collections.Generic.HashSet(Of System.Int32)()
            Dim chunkedChunkCount As System.Int32 = 0

            If doCrossClause Then
                ' Extract full marked-up text from Final view for cross-clause analysis
                Dim fullMarkedUp As System.String = ""
                Try
                    Dim win As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
                    Dim prevRV As Microsoft.Office.Interop.Word.WdRevisionsView = win.View.RevisionsView
                    Dim prevSR As System.Boolean = win.View.ShowRevisionsAndComments
                    Try
                        win.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                        win.View.ShowRevisionsAndComments = False
                        fullMarkedUp = If(activeDoc.Content.Text, "")
                    Finally
                        Try : win.View.RevisionsView = prevRV : Catch : End Try
                        Try : win.View.ShowRevisionsAndComments = prevSR : Catch : End Try
                    End Try
                Catch
                End Try

                If Not System.String.IsNullOrWhiteSpace(fullMarkedUp) Then
                    ' Pre-compute chunks if the document is large and truncation is off
                    Dim contractChunks As System.Collections.Generic.List(Of System.String) = Nothing
                    Dim useChunkedMode As System.Boolean = Not doAutoTruncate AndAlso fullMarkedUp.Length > maxChars
                    If useChunkedMode Then
                        contractChunks = ChunkContract(fullMarkedUp, maxChars)
                        chunkedChunkCount = If(contractChunks IsNot Nothing, contractChunks.Count, 0)
                    End If

                    ' Recalculate total steps now that we know the chunk count
                    If useChunkedMode AndAlso contractChunks IsNot Nothing Then
                        totalSteps = matchedConstraints.Count + (matchedConstraints.Count * contractChunks.Count) + 1
                        GlobalProgressMax = totalSteps
                    End If

                    Dim crossIdx As System.Int32 = 0
                    For Each constraint As MarkupConstraint In matchedConstraints
                        If ProgressBarModule.CancelOperation Then Exit For

                        crossIdx += 1

                        Dim p1Result As ComplianceResult = Nothing
                        For Each cr As ComplianceResult In complianceResults
                            If cr.ConstraintId = constraint.Id Then
                                p1Result = cr
                                Exit For
                            End If
                        Next

                        Dim crossCp As ClausePair = Nothing
                        If constraint.MatchedClauseIndex >= 0 AndAlso constraint.MatchedClauseIndex < clausePairs.Count Then
                            crossCp = clausePairs(constraint.MatchedClauseIndex)
                        End If

                        If useChunkedMode AndAlso contractChunks IsNot Nothing Then
                            ' ── Chunked mode: check every chunk, merge results ──
                            chunkedConstraintIds.Add(constraint.Id)
                            Dim mergedResult As New CrossClauseResult()
                            mergedResult.ConstraintId = constraint.Id
                            mergedResult.UnderminingFound = False
                            Dim detailParts As New System.Collections.Generic.List(Of System.String)()
                            Dim rawParts As New System.Collections.Generic.List(Of System.String)()

                            For chunkIdx As System.Int32 = 0 To contractChunks.Count - 1
                                If ProgressBarModule.CancelOperation Then Exit For

                                Dim progressOffset As System.Int32 = matchedConstraints.Count + ((crossIdx - 1) * contractChunks.Count) + chunkIdx + 1
                                GlobalProgressValue = progressOffset
                                GlobalProgressLabel = $"Pass 2: Constraint {crossIdx}/{matchedConstraints.Count}, chunk {chunkIdx + 1}/{contractChunks.Count}..."

                                Dim chunkResult As CrossClauseResult = Await RunCrossClauseCheck(constraint, contractChunks(chunkIdx), p1Result, crossCp, do2ndModel)
                                If chunkResult IsNot Nothing Then
                                    rawParts.Add(If(chunkResult.RawLLMResponse, ""))
                                    If chunkResult.UnderminingFound Then
                                        mergedResult.UnderminingFound = True
                                        If Not System.String.IsNullOrWhiteSpace(chunkResult.Details) Then
                                            detailParts.Add($"[Chunk {chunkIdx + 1}] {chunkResult.Details}")
                                        End If
                                    End If
                                End If
                            Next

                            mergedResult.Details = System.String.Join(System.Environment.NewLine, detailParts)
                            mergedResult.RawLLMResponse = System.String.Join(System.Environment.NewLine & "---" & System.Environment.NewLine, rawParts)
                            crossResults.Add(mergedResult)
                        Else
                            ' ── Single-call mode: truncate if needed, one call ──
                            GlobalProgressValue = matchedConstraints.Count + crossIdx
                            GlobalProgressLabel = $"Pass 2: Cross-clause check {crossIdx} of {matchedConstraints.Count}..."

                            Dim contractForThisCheck As System.String = fullMarkedUp
                            Dim protectedExceeded As System.Boolean = False
                            If doAutoTruncate AndAlso fullMarkedUp.Length > maxChars Then
                                contractForThisCheck = SmartTruncate(fullMarkedUp, crossCp, maxChars, protectedExceeded)
                                truncatedConstraintIds.Add(constraint.Id)
                                If protectedExceeded Then protectedZoneExceededIds.Add(constraint.Id)
                            End If

                            Dim crossResult As CrossClauseResult = Await RunCrossClauseCheck(constraint, contractForThisCheck, p1Result, crossCp, do2ndModel)
                            If crossResult IsNot Nothing Then crossResults.Add(crossResult)
                        End If
                    Next
                End If
            End If

            If ProgressBarModule.CancelOperation Then Return

            ' ── 10) Output results ──
            GlobalProgressLabel = "Generating output..."
            GlobalProgressValue = totalSteps

            Dim outputDoc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim outputSelection As Microsoft.Office.Interop.Word.Selection = Nothing
            Try
                outputDoc = app.ActiveDocument
                outputSelection = app.Selection
            Catch
            End Try

            If outputDoc Is Nothing Then
                ShowCustomMessageBox("The active document is no longer available. Results will be shown as a report.")
                Dim report As System.String = BuildMarkupReviewReport(constraints, complianceResults, crossResults, clausePairs)
                ShowDocCheckResult(report)
            Else
                InsertMarkupReviewResults(outputDoc, outputSelection, clausePairs, constraints,
                                          complianceResults, crossResults, clauseRanges, doBubbles, doMarkdownBubbles)
            End If

            ProgressBarModule.CancelOperation = True

            ' ── Build completion summary ──
            Dim acceptableCount As System.Int32 = 0
            Dim notAcceptableCount As System.Int32 = 0
            For Each cr As ComplianceResult In complianceResults
                If cr.IsAcceptable Then acceptableCount += 1 Else notAcceptableCount += 1
            Next
            Dim crossWarningCount As System.Int32 = 0
            For Each xr As CrossClauseResult In crossResults
                If xr.UnderminingFound Then crossWarningCount += 1
            Next

            Dim elapsedFmt As System.String = sw.Elapsed.ToString("hh\:mm\:ss\.fff")

            Dim summary As New System.Text.StringBuilder()
            summary.AppendLine("Markup Review completed.")
            summary.AppendLine()
            summary.AppendLine($"Constraints loaded:    {constraints.Count}")
            summary.AppendLine($"  Matched in document: {matchedConstraints.Count}")
            If unmatchedCount > 0 Then
                summary.AppendLine($"  Unmatched (skipped): {unmatchedCount}")
            End If
            summary.AppendLine()
            summary.AppendLine($"Pass 1 — Compliance:")
            summary.AppendLine($"  ✅ Acceptable:       {acceptableCount}")
            summary.AppendLine($"  ❌ Not acceptable:   {notAcceptableCount}")
            If doCrossClause Then
                summary.AppendLine()
                summary.AppendLine($"Pass 2 — Cross-clause:")
                summary.AppendLine($"  ⚠ Undermining found: {crossWarningCount}")
                If chunkedConstraintIds.Count > 0 Then
                    summary.AppendLine($"  📄 Chunked analysis:  {chunkedConstraintIds.Count} constraint(s) × {chunkedChunkCount} chunks " &
                                       $"= {chunkedConstraintIds.Count * chunkedChunkCount} LLM calls")
                End If
                If truncatedConstraintIds.Count > 0 Then
                    summary.AppendLine($"  ⚡ Truncated checks:  {truncatedConstraintIds.Count} (document exceeded {maxChars \ 1000}k chars)")
                End If
                If protectedZoneExceededIds.Count > 0 Then
                    summary.AppendLine()
                    summary.AppendLine($"  ⚠ Note: {protectedZoneExceededIds.Count} constraint(s) have very long clause text that " &
                                       "consumed most of the context window, leaving less room for the rest of the " &
                                       "contract. Cross-clause results for these may be incomplete. Consider increasing " &
                                       "the context window size or using a model with a larger context.")
                End If
            End If
            summary.AppendLine()
            summary.AppendLine($"Analysis time: {elapsedFmt}")

            ShowCustomMessageBox(summary.ToString())

        Catch ex As System.Exception
            ShowCustomMessageBox("Markup Review failed: " & ex.Message)
        Finally
            sw.Stop()
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Function

    ' ========================= Truncation Helper =========================

    ''' <summary>
    ''' Truncates a long contract text for cross-clause analysis while preserving
    ''' the neighbourhood of the checked clause and sampling the rest evenly.
    ''' </summary>
    Private Function SmartTruncate(ByVal fullText As System.String,
                                    ByVal checkedClause As ClausePair,
                                    ByVal maxChars As System.Int32,
                                    ByRef protectedZoneExceeded As System.Boolean) As System.String

        protectedZoneExceeded = False
        If fullText Is Nothing OrElse fullText.Length <= maxChars Then Return fullText

        Dim marker As System.String = vbCrLf & "[... content omitted for length ...]" & vbCrLf

        ' If no clause context, fall back to keeping head and tail equally
        If checkedClause Is Nothing OrElse System.String.IsNullOrWhiteSpace(checkedClause.MarkedUpText) Then
            Dim halfBudget As System.Int32 = (maxChars - marker.Length) \ 2
            Return fullText.Substring(0, halfBudget) & marker & fullText.Substring(fullText.Length - halfBudget)
        End If

        ' Find the checked clause text in the full document
        Dim clauseStart As System.Int32 = fullText.IndexOf(checkedClause.MarkedUpText, System.StringComparison.OrdinalIgnoreCase)
        Dim clauseEnd As System.Int32

        If clauseStart < 0 Then
            ' Clause not found verbatim — try first 80 chars as anchor
            Dim snippet As System.String = checkedClause.MarkedUpText
            If snippet.Length > 80 Then snippet = snippet.Substring(0, 80)
            clauseStart = fullText.IndexOf(snippet, System.StringComparison.OrdinalIgnoreCase)
            If clauseStart < 0 Then
                ' Give up on smart truncation — use head+tail
                Dim halfBudget As System.Int32 = (maxChars - marker.Length) \ 2
                Return fullText.Substring(0, halfBudget) & marker & fullText.Substring(fullText.Length - halfBudget)
            End If
            clauseEnd = System.Math.Min(fullText.Length, clauseStart + checkedClause.MarkedUpText.Length)
        Else
            clauseEnd = clauseStart + checkedClause.MarkedUpText.Length
        End If

        ' Budget allocation
        Dim protectedBudget As System.Int32 = CInt(maxChars * 0.4)
        Dim clauseLen As System.Int32 = clauseEnd - clauseStart

        ' Check if the clause itself exceeds the protected budget
        If clauseLen > protectedBudget Then
            protectedZoneExceeded = True
            ' Cap protected zone to 60% to leave at least 40% for the rest
            protectedBudget = CInt(maxChars * 0.6)
        End If

        Dim remainingBudget As System.Int32 = maxChars - System.Math.Min(protectedBudget, clauseLen + 2000) - (marker.Length * 2)
        If remainingBudget < 0 Then remainingBudget = 0

        ' Protected zone: centre on the clause, expand equally in both directions
        Dim protectedPad As System.Int32 = System.Math.Max(0, (protectedBudget - clauseLen) \ 2)
        Dim protectedStart As System.Int32 = System.Math.Max(0, clauseStart - protectedPad)
        Dim protectedEnd As System.Int32 = System.Math.Min(fullText.Length, clauseEnd + protectedPad)

        ' Text before and after the protected zone
        Dim headAvailable As System.Int32 = protectedStart
        Dim tailAvailable As System.Int32 = fullText.Length - protectedEnd

        Dim headBudget As System.Int32 = remainingBudget \ 2
        Dim tailBudget As System.Int32 = remainingBudget - headBudget

        ' Redistribute surplus if one side is shorter than its budget
        If headAvailable < headBudget Then
            tailBudget += (headBudget - headAvailable)
            headBudget = headAvailable
        ElseIf tailAvailable < tailBudget Then
            headBudget += (tailBudget - tailAvailable)
            tailBudget = tailAvailable
        End If

        ' Clamp to available
        headBudget = System.Math.Min(headBudget, headAvailable)
        tailBudget = System.Math.Min(tailBudget, tailAvailable)

        Dim sb As New System.Text.StringBuilder()

        ' Head section
        If headBudget > 0 AndAlso headBudget < headAvailable Then
            sb.Append(fullText.Substring(0, headBudget))
            sb.Append(marker)
        ElseIf headAvailable > 0 Then
            sb.Append(fullText.Substring(0, headAvailable))
        End If

        ' Protected zone (always included in full)
        sb.Append(fullText.Substring(protectedStart, protectedEnd - protectedStart))

        ' Tail section
        If tailBudget > 0 AndAlso tailBudget < tailAvailable Then
            sb.Append(marker)
            sb.Append(fullText.Substring(fullText.Length - tailBudget))
        ElseIf tailAvailable > 0 Then
            sb.Append(fullText.Substring(protectedEnd))
        End If

        Return sb.ToString()
    End Function


    ' ========================= Constraint Loading =========================

    ''' <summary>
    ''' Loads comment-based constraints from a playbook document.
    ''' </summary>
    Private Function LoadConstraintsFromDocx(ByVal app As Microsoft.Office.Interop.Word.Application,
                                             ByVal docxPath As System.String) As System.Collections.Generic.List(Of MarkupConstraint)
        Dim result As New System.Collections.Generic.List(Of MarkupConstraint)()

        Dim constraintDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = app.DisplayAlerts
        Dim prevScreenUpdating As System.Boolean = app.ScreenUpdating

        Try
            app.ScreenUpdating = False
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

            constraintDoc = app.Documents.Open(FileName:=docxPath, ReadOnly:=True, Visible:=False)

            If constraintDoc.Comments Is Nothing OrElse constraintDoc.Comments.Count = 0 Then
                Return result
            End If

            Dim anchorMap As New System.Collections.Generic.Dictionary(Of System.String, MarkupConstraint)(System.StringComparer.OrdinalIgnoreCase)
            Dim idCounter As System.Int32 = 0

            For Each cmt As Microsoft.Office.Interop.Word.Comment In constraintDoc.Comments
                Try
                    Dim anchorText As System.String = ""
                    Dim anchorStart As System.Int32 = 0
                    Try
                        anchorText = If(cmt.Scope.Text, "").Trim()
                        anchorStart = cmt.Scope.Start
                    Catch
                        Try
                            anchorText = If(cmt.Reference.Text, "").Trim()
                            anchorStart = cmt.Reference.Start
                        Catch
                            Continue For
                        End Try
                    End Try

                    If System.String.IsNullOrWhiteSpace(anchorText) Then Continue For

                    Dim instruction As System.String = If(cmt.Range.Text, "").Trim()
                    If System.String.IsNullOrWhiteSpace(instruction) Then Continue For

                    Dim key As System.String = anchorText
                    If key.Length > 200 Then key = key.Substring(0, 200)

                    Dim mc As MarkupConstraint = Nothing
                    If anchorMap.TryGetValue(key, mc) Then
                        mc.Instructions.Add(instruction)
                    Else
                        idCounter += 1
                        mc = New MarkupConstraint()
                        mc.Id = idCounter
                        mc.AnchorText = anchorText
                        mc.AnchorStart = anchorStart
                        mc.Instructions = New System.Collections.Generic.List(Of System.String)()
                        mc.Instructions.Add(instruction)
                        anchorMap.Add(key, mc)
                    End If
                Catch
                End Try
            Next

            result.AddRange(anchorMap.Values)
            result.Sort(Function(a, b) a.AnchorStart.CompareTo(b.AnchorStart))

        Catch ex As System.Exception
            ShowCustomMessageBox("Error loading constraints: " & ex.Message)
        Finally
            If constraintDoc IsNot Nothing Then
                Try : constraintDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            app.DisplayAlerts = prevAlerts
            app.ScreenUpdating = prevScreenUpdating
        End Try

        Return result
    End Function


    ' ========================= Constraint-Driven Extraction =========================

    ''' <summary>
    ''' For each constraint, finds the target region in the live document using
    ''' FindLongTextAnchoredFast(searchOriginal:=True) and extracts both original
    ''' and marked-up text from the same Range object in the two view modes.
    '''
    ''' Duplicate anchor handling: constraints are sorted by playbook position.
    ''' We track lastFoundEnd to search forward from the previous match. If a
    ''' forward search fails, we retry from the document start — but skip any
    ''' position already claimed by another constraint.
    '''
    ''' Deleted clause handling: when the anchor is found in Original view but
    ''' the Range yields empty text in Final view (clause entirely deleted),
    ''' we expand outward to find the nearest non-empty paragraph in Final view
    ''' so the comment anchors to visible content near the deletion site.
    ''' </summary>
    Private Sub ResolveConstraintRegions(ByVal app As Microsoft.Office.Interop.Word.Application,
                                          ByVal activeDoc As Microsoft.Office.Interop.Word.Document,
                                          ByVal constraints As System.Collections.Generic.List(Of MarkupConstraint),
                                          ByVal clausePairs As System.Collections.Generic.List(Of ClausePair),
                                          ByVal clauseRanges As System.Collections.Generic.Dictionary(Of System.Int32, Microsoft.Office.Interop.Word.Range))

        Dim win As Microsoft.Office.Interop.Word.Window = app.ActiveWindow
        Dim origRevView As Microsoft.Office.Interop.Word.WdRevisionsView = win.View.RevisionsView
        Dim origShowRevs As System.Boolean = win.View.ShowRevisionsAndComments
        Dim prevScreenUpdating As System.Boolean = True

        Try
            Try : prevScreenUpdating = app.ScreenUpdating : Catch : End Try
            Try : app.ScreenUpdating = False : Catch : End Try

            Dim lastFoundEnd As System.Int32 = activeDoc.Content.Start

            ' Track positions already claimed to avoid duplicate-anchor collisions
            Dim claimedPositions As New System.Collections.Generic.List(Of (startPos As System.Int32, endPos As System.Int32))()

            For Each constraint As MarkupConstraint In constraints
                If System.String.IsNullOrWhiteSpace(constraint.AnchorText) Then Continue For

                ' ── Step 1: Find anchor text in Original view ──
                Dim anchorRange As Microsoft.Office.Interop.Word.Range = Nothing
                Try
                    Dim sel As Microsoft.Office.Interop.Word.Selection = app.Selection

                    ' Forward search from last found position
                    sel.SetRange(lastFoundEnd, activeDoc.Content.End)
                    If FindLongTextAnchoredFast(sel, constraint.AnchorText, skipDeleted:=False,
                                                nWords:=4, cancel:=System.Threading.CancellationToken.None,
                                                timeoutSeconds:=15, searchOriginal:=True) Then
                        ' Check this position isn't already claimed
                        Dim candidate As Microsoft.Office.Interop.Word.Range = activeDoc.Range(sel.Range.Start, sel.Range.End)
                        If Not IsPositionClaimed(candidate.Start, candidate.End, claimedPositions) Then
                            anchorRange = candidate
                            lastFoundEnd = sel.Range.End
                        End If
                    End If

                    ' If forward search failed or hit a claimed position, retry from start
                    If anchorRange Is Nothing Then
                        sel.SetRange(activeDoc.Content.Start, activeDoc.Content.End)
                        Dim searchStart As System.Int32 = activeDoc.Content.Start
                        ' Keep searching until we find an unclaimed occurrence
                        While anchorRange Is Nothing
                            sel.SetRange(searchStart, activeDoc.Content.End)
                            If Not FindLongTextAnchoredFast(sel, constraint.AnchorText, skipDeleted:=False,
                                                            nWords:=4, cancel:=System.Threading.CancellationToken.None,
                                                            timeoutSeconds:=15, searchOriginal:=True) Then
                                Exit While ' No more occurrences
                            End If
                            Dim candidateRetry As Microsoft.Office.Interop.Word.Range = activeDoc.Range(sel.Range.Start, sel.Range.End)
                            If Not IsPositionClaimed(candidateRetry.Start, candidateRetry.End, claimedPositions) Then
                                anchorRange = candidateRetry
                            Else
                                ' Skip past this occurrence and try again
                                searchStart = sel.Range.End
                                If searchStart >= activeDoc.Content.End Then Exit While
                            End If
                        End While
                    End If
                Catch
                End Try

                If anchorRange Is Nothing Then Continue For

                ' Record this position as claimed
                claimedPositions.Add((anchorRange.Start, anchorRange.End))

                ' ── Step 2: Expand to paragraph boundaries in Original view ──
                win.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewOriginal
                win.View.ShowRevisionsAndComments = False
                Try
                    Dim startRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(anchorRange.Start, anchorRange.Start)
                    startRange.StartOf(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph,
                                       Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                    anchorRange.Start = startRange.Start

                    Dim endRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(anchorRange.End, anchorRange.End)
                    endRange.EndOf(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph,
                                   Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdMove)
                    anchorRange.End = endRange.End
                Catch
                End Try

                ' ── Step 3: Read original text ──
                Dim originalClause As System.String = ""
                Try
                    originalClause = If(anchorRange.Text, "").Trim()
                Catch
                End Try

                ' ── Step 4: Determine heading label ──
                Dim heading As System.String = ""
                Try
                    For pIdx As System.Int32 = activeDoc.Paragraphs.Count To 1 Step -1
                        Dim para As Microsoft.Office.Interop.Word.Paragraph = activeDoc.Paragraphs(pIdx)
                        If para.Range.Start > anchorRange.Start Then Continue For
                        Try
                            Dim styleName As System.String = CStr(para.Style.NameLocal)
                            If styleName IsNot Nothing AndAlso styleName.StartsWith("Heading", System.StringComparison.OrdinalIgnoreCase) Then
                                heading = If(para.Range.Text, "").Trim().TrimEnd(ChrW(13), ChrW(7), ChrW(11), ChrW(12)).Trim()
                                Exit For
                            End If
                        Catch
                        End Try
                        If para.Range.Start < anchorRange.Start - 5000 Then Exit For
                    Next
                Catch
                End Try
                If System.String.IsNullOrWhiteSpace(heading) Then
                    heading = constraint.AnchorText
                    If heading.Length > 80 Then heading = heading.Substring(0, 80) & "…"
                End If

                ' ── Step 5: Switch to Final view and read marked-up text ──
                win.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                win.View.ShowRevisionsAndComments = False

                Dim markedUpClause As System.String = ""
                Try
                    markedUpClause = If(anchorRange.Text, "").Trim()
                Catch
                End Try

                ' Handle entirely deleted clauses: the Range collapses to zero-width
                ' or yields empty text in Final view. Find the nearest visible paragraph
                ' so we can anchor the comment to something the user can see.
                Dim commentAnchorRange As Microsoft.Office.Interop.Word.Range = anchorRange
                If System.String.IsNullOrWhiteSpace(markedUpClause) Then
                    markedUpClause = "[Clause deleted by counterparty] " & originalClause
                    ' Try the paragraph immediately after the collapsed range
                    Try
                        Dim probeRange As Microsoft.Office.Interop.Word.Range = activeDoc.Range(anchorRange.End, anchorRange.End)
                        probeRange.EndOf(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph,
                                         Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        Dim probeText As System.String = If(probeRange.Text, "").Trim()
                        If probeText.Length > 0 Then
                            commentAnchorRange = probeRange
                        Else
                            ' Try the paragraph before
                            probeRange = activeDoc.Range(anchorRange.Start, anchorRange.Start)
                            probeRange.StartOf(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph,
                                               Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                            probeText = If(probeRange.Text, "").Trim()
                            If probeText.Length > 0 Then
                                commentAnchorRange = probeRange
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' ── Step 6: Create ClausePair and store Range ──
                Dim cp As New ClausePair()
                cp.Index = clausePairs.Count
                cp.Heading = heading
                cp.OriginalText = originalClause
                cp.MarkedUpText = markedUpClause
                cp.DocStartPos = commentAnchorRange.Start
                cp.DocEndPos = commentAnchorRange.End
                clausePairs.Add(cp)

                constraint.MatchedClauseIndex = cp.Index
                clauseRanges(cp.Index) = commentAnchorRange
            Next

        Catch ex As System.Exception
            ShowCustomMessageBox("Error resolving constraint regions: " & ex.Message)
        Finally
            Try : win.View.RevisionsView = origRevView : Catch : End Try
            Try : win.View.ShowRevisionsAndComments = origShowRevs : Catch : End Try
            Try : app.ScreenUpdating = prevScreenUpdating : Catch : End Try
        End Try
    End Sub

    ''' <summary>
    ''' Checks whether a candidate range overlaps with any already-claimed position.
    ''' Two ranges overlap if one starts before the other ends and vice versa.
    ''' </summary>
    Private Function IsPositionClaimed(ByVal startPos As System.Int32,
                                        ByVal endPos As System.Int32,
                                        ByVal claimed As System.Collections.Generic.List(Of (startPos As System.Int32, endPos As System.Int32))) As System.Boolean
        For Each c In claimed
            ' Overlap test: ranges overlap iff start < otherEnd AND end > otherStart
            If startPos < c.endPos AndAlso endPos > c.startPos Then
                Return True
            End If
        Next
        Return False
    End Function


    ' ========================= Pass 1 — Compliance =========================

    ''' <summary>
    ''' Executes a compliance check for a single constraint and clause pair.
    ''' </summary>
    Private Async Function RunComplianceCheck(ByVal constraint As MarkupConstraint,
                                               ByVal clausePair As ClausePair,
                                               ByVal useSecondAPI As System.Boolean) As System.Threading.Tasks.Task(Of ComplianceResult)
        Try
            Dim constraintsJson As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(constraint.Instructions)

            Dim sysPrompt As System.String = SP_MarkupReview_Compliance

            Dim userPrompt As System.String =
                "<ORIGINAL_CLAUSE>" & clausePair.OriginalText & "</ORIGINAL_CLAUSE>" & System.Environment.NewLine &
                "<MARKED_UP_CLAUSE>" & clausePair.MarkedUpText & "</MARKED_UP_CLAUSE>" & System.Environment.NewLine &
                "<CONSTRAINTS>" & constraintsJson & "</CONSTRAINTS>"

            Dim response As System.String = Await LLM(InterpolateAtRuntime(sysPrompt), userPrompt, "", "", 0, useSecondAPI)
            response = response.Trim()

            Dim result As New ComplianceResult()
            result.ConstraintId = constraint.Id
            result.ClauseIndex = clausePair.Index
            result.RawLLMResponse = response

            Dim token As Newtonsoft.Json.Linq.JToken = Nothing
            Dim cleanedResponse As System.String = StripMarkdownCodeFences(response)
            If TryParseJsonToken(cleanedResponse, token) Then
                If TypeOf token Is Newtonsoft.Json.Linq.JObject Then
                    Dim jo As Newtonsoft.Json.Linq.JObject = CType(token, Newtonsoft.Json.Linq.JObject)
                    Dim accToken As Newtonsoft.Json.Linq.JToken = jo("Acceptable")
                    Dim resToken As Newtonsoft.Json.Linq.JToken = jo("Reasoning")
                    Dim cmpToken As Newtonsoft.Json.Linq.JToken = jo("Compromise")
                    result.IsAcceptable = If(accToken IsNot Nothing, CBool(accToken.ToObject(Of System.Boolean)()), False)
                    result.Reasoning = If(resToken IsNot Nothing, resToken.ToString(), response)
                    result.CompromiseRedraft = If(cmpToken IsNot Nothing, cmpToken.ToString(), "")
                End If
            Else
                result.IsAcceptable = response.IndexOf("ACCEPTABLE", System.StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
                                      response.IndexOf("NOT_ACCEPTABLE", System.StringComparison.OrdinalIgnoreCase) < 0
                result.Reasoning = response
                result.CompromiseRedraft = ""
            End If

            Return result

        Catch ex As System.Exception
            Dim errResult As New ComplianceResult()
            errResult.ConstraintId = constraint.Id
            errResult.ClauseIndex = clausePair.Index
            errResult.Reasoning = "Error: " & ex.Message
            Return errResult
        End Try
    End Function

    ''' <summary>
    ''' Removes outer Markdown code fences from a response string.
    ''' </summary>
    Private Function StripMarkdownCodeFences(ByVal response As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(response) Then Return response

        Dim trimmed As System.String = response.Trim()

        If trimmed.StartsWith("```", System.StringComparison.Ordinal) Then
            Dim firstNewline As System.Int32 = trimmed.IndexOfAny(New System.Char() {CChar(vbLf), CChar(vbCr)})
            If firstNewline > 0 Then
                trimmed = trimmed.Substring(firstNewline).Trim()
            Else
                trimmed = trimmed.Substring(3).Trim()
            End If
        End If

        If trimmed.EndsWith("```", System.StringComparison.Ordinal) Then
            trimmed = trimmed.Substring(0, trimmed.Length - 3).Trim()
        End If

        Return trimmed
    End Function



    ' ========================= Pass 2 — Cross-Clause =========================

    ''' <summary>
    ''' Executes a cross-clause check for a single constraint against a contract text snapshot.
    ''' </summary>
    Private Async Function RunCrossClauseCheck(ByVal constraint As MarkupConstraint,
                                                ByVal fullMarkedUpText As System.String,
                                                ByVal pass1Result As ComplianceResult,
                                                ByVal clausePair As ClausePair,
                                                ByVal useSecondAPI As System.Boolean) As System.Threading.Tasks.Task(Of CrossClauseResult)
        Try
            Dim constraintsJson As System.String = Newtonsoft.Json.JsonConvert.SerializeObject(constraint.Instructions)
            Dim pass1Summary As System.String = ""
            If pass1Result IsNot Nothing Then
                pass1Summary = $"Pass 1 found: {If(pass1Result.IsAcceptable, "ACCEPTABLE", "NOT ACCEPTABLE")}. Reasoning: {pass1Result.Reasoning}"
            End If

            Dim clauseContext As System.String = ""
            If clausePair IsNot Nothing Then
                clauseContext =
                    "The constraint was checked against the following clause:" & System.Environment.NewLine &
                    "Clause heading: " & If(clausePair.Heading, "(unknown)") & System.Environment.NewLine &
                    "Clause text (marked-up version): " & clausePair.MarkedUpText & System.Environment.NewLine &
                    System.Environment.NewLine &
                    "You must EXCLUDE this clause from your cross-clause analysis — it has already been " &
                    "evaluated in Pass 1. Focus only on OTHER clauses in the contract."
            End If

            Dim sysPrompt As System.String = SP_MarkupReview_CrossClause

            Dim userPrompt As System.String =
                "<CONSTRAINT>" & constraintsJson & "</CONSTRAINT>" & System.Environment.NewLine &
                "<CONSTRAINT_ANCHOR>" & If(constraint.AnchorText, "") & "</CONSTRAINT_ANCHOR>" & System.Environment.NewLine &
                "<CHECKED_CLAUSE>" & clauseContext & "</CHECKED_CLAUSE>" & System.Environment.NewLine &
                "<PASS1_RESULT>" & pass1Summary & "</PASS1_RESULT>" & System.Environment.NewLine &
                "<FULL_CONTRACT>" & fullMarkedUpText & "</FULL_CONTRACT>"

            Dim response As System.String = Await LLM(InterpolateAtRuntime(sysPrompt), userPrompt, "", "", 0, useSecondAPI)
            response = response.Trim()

            Dim result As New CrossClauseResult()
            result.ConstraintId = constraint.Id
            result.RawLLMResponse = response

            Dim token As Newtonsoft.Json.Linq.JToken = Nothing
            Dim cleanedResponse As System.String = StripMarkdownCodeFences(response)
            If TryParseJsonToken(cleanedResponse, token) Then
                If TypeOf token Is Newtonsoft.Json.Linq.JObject Then
                    Dim jo As Newtonsoft.Json.Linq.JObject = CType(token, Newtonsoft.Json.Linq.JObject)
                    Dim undToken As Newtonsoft.Json.Linq.JToken = jo("UnderminingFound")
                    Dim detToken As Newtonsoft.Json.Linq.JToken = jo("Details")
                    result.UnderminingFound = If(undToken IsNot Nothing, undToken.ToObject(Of System.Boolean)(), False)
                    result.Details = If(detToken IsNot Nothing, detToken.ToString(), "")
                End If
            Else
                result.UnderminingFound = response.IndexOf("true", System.StringComparison.OrdinalIgnoreCase) >= 0
                result.Details = response
            End If

            Return result

        Catch ex As System.Exception
            Dim errResult As New CrossClauseResult()
            errResult.ConstraintId = constraint.Id
            errResult.Details = "Error: " & ex.Message
            Return errResult
        End Try
    End Function

    ' ========================= Output =========================

    ''' <summary>
    ''' Inserts Markup Review results into Word comments or produces a report.
    ''' </summary>
    Private Sub InsertMarkupReviewResults(ByVal doc As Microsoft.Office.Interop.Word.Document,
                                           ByVal currentSelection As Microsoft.Office.Interop.Word.Selection,
                                           ByVal clausePairs As System.Collections.Generic.List(Of ClausePair),
                                           ByVal constraints As System.Collections.Generic.List(Of MarkupConstraint),
                                           ByVal complianceResults As System.Collections.Generic.List(Of ComplianceResult),
                                           ByVal crossResults As System.Collections.Generic.List(Of CrossClauseResult),
                                           ByVal clauseRanges As System.Collections.Generic.Dictionary(Of System.Int32, Microsoft.Office.Interop.Word.Range),
                                           ByVal doBubbles As System.Boolean,
                                           ByVal doMarkdownBubbles As System.Boolean)

        If Not doBubbles Then
            Dim report As System.String = BuildMarkupReviewReport(constraints, complianceResults, crossResults, clausePairs)
            ShowDocCheckResult(report)
            Return
        End If

        Dim oldMarkdownBubbles As System.Boolean = INI_MarkdownBubbles
        Dim oldFormatInstruction As System.String = FormatInstruction

        Dim app As Microsoft.Office.Interop.Word.Application = doc.Application
        Dim win As Microsoft.Office.Interop.Word.Window = app.ActiveWindow

        Dim prevScreenUpdating As System.Boolean = True
        Dim prevShowRevs As System.Boolean = False
        Dim prevRevView As Microsoft.Office.Interop.Word.WdRevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal

        Try
            INI_MarkdownBubbles = doMarkdownBubbles
            FormatInstruction = If(doMarkdownBubbles, SP_Add_Bubbles_Format, "")

            Try : prevScreenUpdating = app.ScreenUpdating : Catch : End Try
            Try : app.ScreenUpdating = False : Catch : End Try
            Try
                prevShowRevs = win.View.ShowRevisionsAndComments
                prevRevView = win.View.RevisionsView
            Catch
            End Try

            Try
                win.View.RevisionsView = Microsoft.Office.Interop.Word.WdRevisionsView.wdRevisionsViewFinal
                win.View.ShowRevisionsAndComments = False
            Catch
            End Try

            For Each cr As ComplianceResult In complianceResults
                Dim constraint As MarkupConstraint = Nothing
                For Each c As MarkupConstraint In constraints
                    If c.Id = cr.ConstraintId Then
                        constraint = c
                        Exit For
                    End If
                Next
                If constraint Is Nothing Then Continue For

                Dim cp As ClausePair = Nothing
                If cr.ClauseIndex >= 0 AndAlso cr.ClauseIndex < clausePairs.Count Then
                    cp = clausePairs(cr.ClauseIndex)
                End If

                ' Build comment text
                Dim status As System.String = If(cr.IsAcceptable, "✅ ACCEPTABLE", "❌ NOT ACCEPTABLE")
                Dim cpHeading As System.String = If(cp IsNot Nothing, cp.Heading, "Unknown")
                Dim commentBody As System.String = $"**Markup Review — {cpHeading}**" & vbCrLf &
                    $"**Status:** {status}" & vbCrLf &
                    $"**Reasoning:** {cr.Reasoning}"

                Dim crossResult As CrossClauseResult = Nothing
                For Each xr As CrossClauseResult In crossResults
                    If xr.ConstraintId = cr.ConstraintId Then
                        crossResult = xr
                        Exit For
                    End If
                Next

                If crossResult IsNot Nothing AndAlso crossResult.UnderminingFound Then
                    commentBody &= vbCrLf & vbCrLf & "**⚠ Cross-Clause Warning:** " & crossResult.Details
                End If

                If Not cr.IsAcceptable AndAlso Not System.String.IsNullOrWhiteSpace(cr.CompromiseRedraft) Then
                    commentBody &= vbCrLf & vbCrLf & "**Suggested Compromise:**" & vbCrLf & cr.CompromiseRedraft
                End If

                If Not doMarkdownBubbles Then
                    commentBody = commentBody.Replace("**", "").Replace("*", "")
                End If

                ' Get the pre-resolved Range
                Dim anchorRange As Microsoft.Office.Interop.Word.Range = Nothing
                If cp IsNot Nothing Then clauseRanges.TryGetValue(cp.Index, anchorRange)

                ' Fallback: use constraint anchor text via SetBubbles
                If anchorRange Is Nothing Then
                    Dim fallbackAnchor As System.String = If(constraint.AnchorText, "")
                    If fallbackAnchor.Length > 200 Then fallbackAnchor = fallbackAnchor.Substring(0, 200)

                    If fallbackAnchor.Length > 0 Then
                        Dim bubblePayload As System.String = fallbackAnchor & "@@" & commentBody
                        SetBubbles(bubblePayload, currentSelection, True)
                        Continue For
                    Else
                        AddNoticeBubbleAt(doc, 0, commentBody, " Markup Review")
                        Continue For
                    End If
                End If

                ' Insert comment on pre-resolved range
                Try
                    If doMarkdownBubbles Then
                        Dim cmt As Microsoft.Office.Interop.Word.Comment = doc.Comments.Add(Range:=anchorRange, Text:="")
                        Dim cRng As Microsoft.Office.Interop.Word.Range = cmt.Range
                        cRng.Text = ""
                        Dim prevShow As System.Boolean = win.View.ShowRevisionsAndComments
                        Try
                            win.View.ShowRevisionsAndComments = True
                            InsertMarkdownToComment(cRng, $"{AN5}: " & commentBody)
                        Finally
                            win.View.ShowRevisionsAndComments = prevShow
                        End Try
                    Else
                        doc.Comments.Add(Range:=anchorRange, Text:=$"{AN5}: " & commentBody)
                    End If
                Catch
                    AddNoticeBubbleAt(doc, 0, commentBody, " Markup Review")
                End Try
            Next

            ' Orphaned cross-clause warnings
            For Each xr As CrossClauseResult In crossResults
                If Not xr.UnderminingFound Then Continue For
                Dim alreadyReported As System.Boolean = False
                For Each cr As ComplianceResult In complianceResults
                    If cr.ConstraintId = xr.ConstraintId Then
                        alreadyReported = True
                        Exit For
                    End If
                Next
                If alreadyReported Then Continue For
                Dim warningText As System.String = "⚠ Cross-Clause Warning: " & xr.Details
                If Not doMarkdownBubbles Then warningText = warningText.Replace("**", "").Replace("*", "")
                AddNoticeBubbleAt(doc, 0, warningText, " Markup Review")
            Next

        Catch ex As System.Exception
            ShowCustomMessageBox("Error inserting Markup Review results: " & ex.Message)
        Finally
            INI_MarkdownBubbles = oldMarkdownBubbles
            FormatInstruction = oldFormatInstruction
            Try : win.View.RevisionsView = prevRevView : Catch : End Try
            Try : win.View.ShowRevisionsAndComments = prevShowRevs : Catch : End Try
            Try : app.ScreenUpdating = prevScreenUpdating : Catch : End Try
        End Try
    End Sub


    ''' <summary>
    ''' Builds a Markdown report from compliance and cross-clause results.
    ''' </summary>
    Private Function BuildMarkupReviewReport(ByVal constraints As System.Collections.Generic.List(Of MarkupConstraint),
                                              ByVal complianceResults As System.Collections.Generic.List(Of ComplianceResult),
                                              ByVal crossResults As System.Collections.Generic.List(Of CrossClauseResult),
                                              ByVal clausePairs As System.Collections.Generic.List(Of ClausePair)) As System.String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("# Markup Review Report")
        sb.AppendLine()

        For Each cr As ComplianceResult In complianceResults
            Dim constraint As MarkupConstraint = Nothing
            For Each c As MarkupConstraint In constraints
                If c.Id = cr.ConstraintId Then constraint = c : Exit For
            Next

            Dim cp As ClausePair = Nothing
            If cr.ClauseIndex >= 0 AndAlso cr.ClauseIndex < clausePairs.Count Then cp = clausePairs(cr.ClauseIndex)

            Dim cpHeading As System.String = If(cp IsNot Nothing, cp.Heading, "Constraint " & cr.ConstraintId)
            sb.AppendLine($"## {cpHeading}")
            sb.AppendLine()
            sb.AppendLine($"**Status:** {If(cr.IsAcceptable, "✅ ACCEPTABLE", "❌ NOT ACCEPTABLE")}")
            sb.AppendLine()
            sb.AppendLine($"**Reasoning:** {cr.Reasoning}")
            sb.AppendLine()

            For Each xr As CrossClauseResult In crossResults
                If xr.ConstraintId = cr.ConstraintId AndAlso xr.UnderminingFound Then
                    sb.AppendLine($"**⚠ Cross-Clause Warning:** {xr.Details}")
                    sb.AppendLine()
                End If
            Next

            If Not cr.IsAcceptable AndAlso Not System.String.IsNullOrWhiteSpace(cr.CompromiseRedraft) Then
                sb.AppendLine("**Suggested Compromise:**")
                sb.AppendLine()
                sb.AppendLine(cr.CompromiseRedraft)
                sb.AppendLine()
            End If

            sb.AppendLine("---")
            sb.AppendLine()
        Next

        Return sb.ToString()
    End Function

    ' ========================= Chunking Helper =========================

    ''' <summary>
    ''' Splits a long contract text into overlapping chunks that respect paragraph
    ''' boundaries. Each chunk is at most maxChars long, with an overlap region
    ''' so that clauses straddling a boundary appear fully in at least one chunk.
    ''' </summary>
    ''' <param name="fullText">The full contract text.</param>
    ''' <param name="maxChars">Maximum characters per chunk.</param>
    ''' <param name="overlapChars">
    ''' Size of the overlap between adjacent chunks. Default is 10% of maxChars,
    ''' clamped between 2000 and 20000 characters.
    ''' </param>
    ''' <returns>
    ''' A list of chunk strings. For documents shorter than maxChars, returns
    ''' a single-element list containing the full text.
    ''' </returns>
    Private Function ChunkContract(ByVal fullText As System.String,
                                    ByVal maxChars As System.Int32,
                                    Optional ByVal overlapChars As System.Int32 = -1) As System.Collections.Generic.List(Of System.String)

        Dim chunks As New System.Collections.Generic.List(Of System.String)()
        If System.String.IsNullOrEmpty(fullText) Then Return chunks

        If fullText.Length <= maxChars Then
            chunks.Add(fullText)
            Return chunks
        End If

        ' Default overlap: 10% of maxChars, clamped to [2000, 20000]
        If overlapChars < 0 Then
            overlapChars = CInt(maxChars * 0.1)
        End If
        overlapChars = System.Math.Max(2000, System.Math.Min(20000, overlapChars))

        ' Step size: how far we advance between chunks
        Dim stepSize As System.Int32 = maxChars - overlapChars
        If stepSize < 1000 Then stepSize = 1000 ' Safety floor

        Dim pos As System.Int32 = 0
        While pos < fullText.Length
            Dim chunkEnd As System.Int32 = System.Math.Min(pos + maxChars, fullText.Length)

            ' Try to break at a paragraph boundary (double newline or CR+LF)
            If chunkEnd < fullText.Length Then
                ' Search backward from chunkEnd for a paragraph break
                Dim searchStart As System.Int32 = System.Math.Max(pos + stepSize, chunkEnd - 2000)
                Dim breakPos As System.Int32 = -1

                ' Look for double newline first
                Dim dblNewline As System.Int32 = fullText.LastIndexOf(vbCrLf & vbCrLf, chunkEnd - 1, chunkEnd - searchStart, System.StringComparison.Ordinal)
                If dblNewline >= searchStart Then
                    breakPos = dblNewline + 4 ' After the double CRLF
                End If

                ' Fall back to single newline
                If breakPos < 0 Then
                    Dim singleNewline As System.Int32 = fullText.LastIndexOf(vbCrLf, chunkEnd - 1, chunkEnd - searchStart, System.StringComparison.Ordinal)
                    If singleNewline >= searchStart Then
                        breakPos = singleNewline + 2
                    End If
                End If

                ' Fall back to any LF
                If breakPos < 0 Then
                    Dim lf As System.Int32 = fullText.LastIndexOf(CChar(vbLf), chunkEnd - 1, chunkEnd - searchStart)
                    If lf >= searchStart Then
                        breakPos = lf + 1
                    End If
                End If

                If breakPos > 0 Then chunkEnd = breakPos
            End If

            chunks.Add(fullText.Substring(pos, chunkEnd - pos))

            ' Advance by stepSize (not by chunk length) to ensure overlap
            pos += stepSize
            If pos >= fullText.Length Then Exit While

            ' Safety: if we'd produce a tiny final chunk, just extend the previous one
            If fullText.Length - pos < overlapChars Then
                ' Replace the last chunk with one that extends to the end
                Dim lastStart As System.Int32 = pos - stepSize
                chunks(chunks.Count - 1) = fullText.Substring(lastStart)
                Exit While
            End If
        End While

        Return chunks
    End Function

End Class