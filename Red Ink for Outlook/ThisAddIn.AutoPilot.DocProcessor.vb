' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.DocProcessor.vb
' Purpose: Ported OpenXML document processing engine from the Word add-in's
'          ThisAddIn.TranslateDocuments.vb. Operates directly on DOCX XML to modify
'          text nodes while preserving formatting.
'
' Architecture:
'  - Extracts paragraphs and text runs from OpenXML and batches them with context
'    windows for LLM processing.
'  - Uses a run boundary marker (|) for multi-run paragraphs to preserve formatting
'    boundaries when reapplying translated text.
'  - Applies processed text back into document.xml and other sub-parts (headers,
'    footers, comments, footnotes, endnotes).
'  - Does not use Word interop for core processing (compare document generation is
'    handled elsewhere).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CONSTANTS (matching Word add-in)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Number of paragraphs included before the current batch as context.
    ''' </summary>
    Private Const AP_ContextBefore As Integer = 3

    ''' <summary>
    ''' Number of paragraphs included after the current batch as context.
    ''' </summary>
    Private Const AP_ContextAfter As Integer = 2

    ''' <summary>
    ''' Maximum number of paragraphs sent to the LLM per batch.
    ''' </summary>
    Private Const AP_ParagraphsPerBatch As Integer = 10

    ''' <summary>
    ''' Maximum total characters per batch.
    ''' </summary>
    Private Const AP_MaxCharsPerBatch As Integer = 15000

    ''' <summary>Marker inserted between formatting runs so the LLM can preserve boundaries.</summary>
    Private Const AP_RunBoundaryMarker As String = "|"

    ''' <summary>
    ''' Marker inserted at footnote/endnote reference boundaries in text sent to the LLM.
    ''' U+2016 DOUBLE VERTICAL LINE — virtually never appears in legal documents.
    ''' </summary>
    Private Const AP_NoteRefMarker As String = "‖"

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASSES (local to AutoPilot, avoid collision with Word add-in)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Stores a text node and its original content for a single run.
    ''' </summary>
    Private Class APTextRunInfo
        Public Property TextNode As System.Xml.XmlNode
        Public Property OriginalText As String
        Public Property HasNoteReferenceBefore As Boolean
    End Class

    ''' <summary>
    ''' Represents a paragraph, its runs, and processed text state.
    ''' </summary>
    Private Class APParagraphInfo
        Public Property Index As Integer
        Public Property TextRuns As List(Of APTextRunInfo)
        Public Property FullText As String

        ''' <summary>Text with | markers between runs (for multi-run paragraphs).</summary>
        Public Property MarkerText As String

        Public Property TranslatedText As String
        Public Property IsEmpty As Boolean
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MAIN ENTRY POINT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes a DOCX file by applying the given instruction via LLM to all text paragraphs.
    ''' </summary>
    Private Async Function ProcessDocxForAutoPilot(inputPath As String, outputPath As String,
                                                    instruction As String, ct As CancellationToken) As Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & "xml_" & Guid.NewGuid().ToString("N"))

        Try
            ApDashboardLog("DocProcessor: starting document processing", "step")

            ' Copy input to output (we modify the copy)
            File.Copy(inputPath, outputPath, overwrite:=True)

            ' Extract DOCX (ZIP)
            ZipFile.ExtractToDirectory(outputPath, tempDir)

            ' Process document.xml
            Dim documentXmlPath = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then Return False

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            Dim paragraphs = APExtractParagraphs(xmlDoc, nsMgr)
            If paragraphs.Count = 0 OrElse paragraphs.All(Function(p) p.IsEmpty) Then Return False

            ApDashboardLog($"DocProcessor: extracted {paragraphs.Count} paragraphs from document.xml", "step")

            Dim success = Await APProcessBatches(paragraphs, instruction, ct)
            If Not success Then Return False

            APApplyTranslations(paragraphs)
            xmlDoc.Save(documentXmlPath)

            ' Process headers, footers, comments, footnotes, endnotes
            ApDashboardLog("DocProcessor: processing sub-parts (headers, footers, etc.)", "step")
            Await APProcessSubParts(tempDir, instruction, ct)

            ' Repack DOCX
            File.Delete(outputPath)
            ZipFile.CreateFromDirectory(tempDir, outputPath, CompressionLevel.Optimal, False)

            ApDashboardLog("DocProcessor: document processing complete", "success")
            Return True

        Catch ex As System.Exception
            Debug.WriteLine("ProcessDocxForAutoPilot error: " & ex.Message)
            ApDashboardLog("DocProcessor: error - " & ex.Message, "error")
            Return False
        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PARAGRAPH EXTRACTION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts paragraphs and their text runs from the given document.
    ''' Detects footnote/endnote references AND complex field boundaries (cross-references,
    ''' merge fields, etc.) so that text redistribution never shifts content across them.
    ''' </summary>
    Private Function APExtractParagraphs(xmlDoc As System.Xml.XmlDocument,
                                          nsMgr As System.Xml.XmlNamespaceManager) As List(Of APParagraphInfo)
        Dim paragraphs As New List(Of APParagraphInfo)()
        Dim paraNodes = xmlDoc.SelectNodes("//w:p", nsMgr)
        Dim paraIndex As Integer = 0

        For Each paraNode As System.Xml.XmlNode In paraNodes
            Dim paraInfo As New APParagraphInfo() With {
                .Index = paraIndex,
                .TextRuns = New List(Of APTextRunInfo)(),
                .TranslatedText = Nothing
            }

            ' Identify w:t nodes to skip: footnoteRef / endnoteRef runs and their separator spaces
            ' (inside footnotes.xml / endnotes.xml)
            Dim refNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim noteRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteRef or w:endnoteRef]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In noteRefElements
                For Each tInRef As System.Xml.XmlNode In refRunNode.SelectNodes(".//w:t", nsMgr)
                    refNodes.Add(tInRef)
                Next
                Dim nextSibling As System.Xml.XmlNode = refRunNode.NextSibling
                While nextSibling IsNot Nothing AndAlso nextSibling.NodeType <> System.Xml.XmlNodeType.Element
                    nextSibling = nextSibling.NextSibling
                End While
                If nextSibling IsNot Nothing AndAlso nextSibling.LocalName = "r" Then
                    Dim siblingTexts As System.Xml.XmlNodeList = nextSibling.SelectNodes(".//w:t", nsMgr)
                    If siblingTexts.Count = 1 Then
                        Dim sibText As String = siblingTexts(0).InnerText
                        If String.IsNullOrWhiteSpace(sibText) AndAlso sibText.Length <= 2 Then
                            refNodes.Add(siblingTexts(0))
                        End If
                    End If
                End If
            Next

            ' Identify w:r nodes containing w:footnoteReference / w:endnoteReference (document body)
            Dim bodyRefRunNodes As New HashSet(Of System.Xml.XmlNode)()
            Dim bodyRefElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:footnoteReference or w:endnoteReference]", nsMgr)
            For Each refRunNode As System.Xml.XmlNode In bodyRefElements
                bodyRefRunNodes.Add(refRunNode)
            Next

            ' ─── Complex field detection ───
            ' Complex fields (cross-references, merge fields, etc.) use w:fldChar elements.
            ' The w:r runs containing fldChar sit between text runs like footnote references.
            ' We treat them as boundaries and exclude field code text from extraction.
            Dim fldCharRuns As New HashSet(Of System.Xml.XmlNode)()
            Dim fldCharElements As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r[w:fldChar]", nsMgr)
            For Each fldCharRunNode As System.Xml.XmlNode In fldCharElements
                fldCharRuns.Add(fldCharRunNode)
            Next

            ' Track field nesting to exclude w:t inside field code regions (begin→separate)
            If fldCharRuns.Count > 0 Then
                Dim fieldDepth As Integer = 0
                Dim inFieldCode As Boolean = False

                For Each childNode As System.Xml.XmlNode In paraNode.ChildNodes
                    If childNode.NodeType <> System.Xml.XmlNodeType.Element Then Continue For
                    If childNode.LocalName <> "r" Then Continue For

                    Dim fldChar As System.Xml.XmlNode = childNode.SelectSingleNode("w:fldChar", nsMgr)
                    If fldChar IsNot Nothing Then
                        Dim fldType As String = ""
                        Dim fldTypeAttr = fldChar.Attributes("w:fldCharType")
                        If fldTypeAttr IsNot Nothing Then fldType = fldTypeAttr.Value

                        Select Case fldType
                            Case "begin"
                                fieldDepth += 1
                                inFieldCode = True
                            Case "separate"
                                inFieldCode = False
                            Case "end"
                                fieldDepth -= 1
                                If fieldDepth <= 0 Then
                                    fieldDepth = 0
                                    inFieldCode = False
                                End If
                        End Select
                        Continue For
                    End If

                    If inFieldCode Then
                        For Each tNode As System.Xml.XmlNode In childNode.SelectNodes(".//w:t", nsMgr)
                            refNodes.Add(tNode)
                        Next
                    End If
                Next
            End If

            Dim textNodes = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            For Each textNode As System.Xml.XmlNode In textNodes
                If refNodes.Contains(textNode) Then Continue For

                Dim text As String = textNode.InnerText

                ' Check if a footnoteReference/endnoteReference run OR a fldChar run
                ' sits between the previous text run and this one
                Dim hasNoteRefBefore As Boolean = False
                If paraInfo.TextRuns.Count > 0 AndAlso (bodyRefRunNodes.Count > 0 OrElse fldCharRuns.Count > 0) Then
                    Dim thisRun As System.Xml.XmlNode = textNode.ParentNode
                    Dim prevEl As System.Xml.XmlNode = thisRun.PreviousSibling
                    While prevEl IsNot Nothing
                        If prevEl.NodeType = System.Xml.XmlNodeType.Element Then
                            If prevEl.LocalName = "r" Then
                                If bodyRefRunNodes.Contains(prevEl) OrElse fldCharRuns.Contains(prevEl) Then
                                    hasNoteRefBefore = True
                                    Exit While
                                End If
                                Dim prevTexts As System.Xml.XmlNodeList = prevEl.SelectNodes(".//w:t", nsMgr)
                                Dim foundPrevTextRun As Boolean = False
                                For Each pt As System.Xml.XmlNode In prevTexts
                                    If Not refNodes.Contains(pt) Then
                                        foundPrevTextRun = True
                                        Exit For
                                    End If
                                Next
                                If foundPrevTextRun Then Exit While
                            End If
                        End If
                        prevEl = prevEl.PreviousSibling
                    End While
                End If

                paraInfo.TextRuns.Add(New APTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = text,
                    .HasNoteReferenceBefore = hasNoteRefBefore
                })
                ' Insert note-reference marker into FullText at the boundary
                If hasNoteRefBefore Then
                    fullTextBuilder.Append(AP_NoteRefMarker)
                End If

                fullTextBuilder.Append(text)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            If paraInfo.TextRuns.Count > 1 AndAlso Not paraInfo.IsEmpty Then
                paraInfo.MarkerText = APBuildMarkerAnnotatedText(paraInfo.TextRuns)
            Else
                paraInfo.MarkerText = Nothing
            End If

            paragraphs.Add(paraInfo)
            paraIndex += 1
        Next

        Return paragraphs
    End Function



    ''' <summary>
    ''' Builds text with | markers between formatting runs (matching Word add-in's BuildMarkerAnnotatedText).
    ''' Only inserts markers between non-empty runs where formatting actually changes.
    ''' Also preserves ‖ (note-reference) markers at footnote/endnote boundaries.
    ''' </summary>
    Private Shared Function APBuildMarkerAnnotatedText(textRuns As List(Of APTextRunInfo)) As String
        If textRuns Is Nothing OrElse textRuns.Count <= 1 Then Return Nothing

        Dim sb As New StringBuilder()
        Dim markerCount As Integer = 0

        For i As Integer = 0 To textRuns.Count - 1
            Dim runText As String = textRuns(i).OriginalText

            ' Insert ‖ marker BEFORE the | marker if this run has a note reference before it.
            ' The ‖ must appear in the text sent to the LLM so it can preserve it.
            If textRuns(i).HasNoteReferenceBefore Then
                sb.Append(AP_NoteRefMarker)
            End If

            ' Insert marker between runs, but only when both sides are non-empty
            ' (empty runs don't carry visible formatting, so a marker would be misleading)
            If i > 0 AndAlso runText.Length > 0 Then
                Dim prevNonEmpty As Boolean = False
                For j As Integer = i - 1 To 0 Step -1
                    If textRuns(j).OriginalText.Length > 0 Then
                        prevNonEmpty = True
                        Exit For
                    End If
                Next
                If prevNonEmpty Then
                    sb.Append(AP_RunBoundaryMarker)
                    markerCount += 1
                End If
            End If

            sb.Append(runText)
        Next

        ' If no markers were inserted, there's no benefit to using marker mode
        If markerCount = 0 Then Return Nothing

        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  BATCH PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Batches paragraphs, sends them to the LLM, and stores processed results.
    ''' </summary>
    Private Async Function APProcessBatches(paragraphs As List(Of APParagraphInfo),
                                             instruction As String, ct As CancellationToken) As Task(Of Boolean)

        Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processable.Count = 0 Then Return True

        ' Check if any paragraphs use markers
        Dim hasMarkers = processable.Any(Function(p) p.MarkerText IsNot Nothing)

        ' Build system prompt for document processing
        Dim systemPrompt As String =
            "You are a professional document processor. Apply the following instruction to the numbered paragraphs " &
            "in the [TEXTTOPROCESS] section." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            "RULES:" & vbCrLf &
            "1. Process ONLY paragraphs inside [TEXTTOPROCESS], not the context sections." & vbCrLf &
            "2. Use [CONTEXT BEFORE] and [CONTEXT AFTER] to understand meaning, tone, and terminology." & vbCrLf &
            "3. Return each processed paragraph with its [n] marker exactly as shown." & vbCrLf &
            "4. The processed text should have approximately the same number of words." & vbCrLf &
            "5. Maintain consistent terminology and style." & vbCrLf &
            "6. Return ONLY the [n] processed paragraphs, no explanations."

        If hasMarkers Then
            systemPrompt &= vbCrLf &
            "7. Some paragraphs contain pipe characters (|) that mark formatting boundaries. " &
            "IMPORTANT: Preserve these | markers in your output at approximately the same positions " &
            "relative to the text. The number of | markers in each paragraph must stay EXACTLY the same. " &
            "Do NOT add or remove any | markers."
        End If

        Dim hasNoteRefMarkers As Boolean = processable.Any(Function(p) p.FullText.Contains(AP_NoteRefMarker))
        If hasNoteRefMarkers Then
            systemPrompt &= vbCrLf &
            If(hasMarkers, "8", "7") & ". Some paragraphs contain the character ‖ (double vertical line). " &
            "This marks the position of a footnote or endnote reference. " &
            "CRITICAL: Keep each ‖ at EXACTLY the same position relative to the surrounding words. " &
            "Do NOT move, add, or remove any ‖ characters."
        End If

        Dim batchIndex As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(processable.Count / CDbl(AP_ParagraphsPerBatch)))
        Dim currentBatch As Integer = 0

        ApDashboardLog($"DocProcessor: {processable.Count} paragraphs to process (~{totalBatches} batches)", "step")

        While batchIndex < processable.Count
            ct.ThrowIfCancellationRequested()

            currentBatch += 1

            Dim batchStart = batchIndex
            Dim batchEnd = Math.Min(batchIndex + AP_ParagraphsPerBatch - 1, processable.Count - 1)

            ' Adjust for character limit
            Dim batchChars As Integer = 0
            For j = batchStart To batchEnd
                batchChars += processable(j).FullText.Length
                If batchChars > AP_MaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            ApDashboardLog($"DocProcessor: batch {currentBatch}/{totalBatches} (paragraphs {batchStart + 1}-{batchEnd + 1})", "step")

            ' Build prompt
            Dim promptBuilder As New StringBuilder()

            ' Context Before
            Dim contextBeforeStart = Math.Max(0, batchStart - AP_ContextBefore)
            If contextBeforeStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                For j = contextBeforeStart To batchStart - 1
                    Dim contextText = If(processable(j).TranslatedText, processable(j).FullText)
                    ' Strip markers from context to avoid LLM echoing them into non-marker paragraphs
                    If contextText IsNot Nothing Then
                        contextText = contextText.Replace(AP_RunBoundaryMarker, "")
                        contextText = contextText.Replace(AP_NoteRefMarker, "")
                    End If
                    promptBuilder.AppendLine(contextText)
                Next
                promptBuilder.AppendLine()
            End If

            ' Paragraphs to process — use marker text when available
            promptBuilder.AppendLine("[TEXTTOPROCESS]")
            Dim batchNumber = 1
            For j = batchStart To batchEnd
                Dim paraText = If(processable(j).MarkerText, processable(j).FullText)
                promptBuilder.AppendLine("[" & batchNumber.ToString() & "] " & paraText)
                batchNumber += 1
            Next
            promptBuilder.AppendLine("[/TEXTTOPROCESS]")
            promptBuilder.AppendLine()

            ' Context After
            Dim contextAfterEnd = Math.Min(processable.Count - 1, batchEnd + AP_ContextAfter)
            If contextAfterEnd > batchEnd Then
                promptBuilder.AppendLine("[CONTEXT AFTER - for reference only]")
                For j = batchEnd + 1 To contextAfterEnd
                    Dim ctxAfterText As String = processable(j).FullText
                    If ctxAfterText IsNot Nothing Then
                        ctxAfterText = ctxAfterText.Replace(AP_NoteRefMarker, "")
                    End If
                    promptBuilder.AppendLine(ctxAfterText)
                Next
            End If

            ' Call LLM
            Dim llmResponse = Await LLM(systemPrompt, promptBuilder.ToString(),
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResponse) Then
                ApDashboardLog($"DocProcessor: batch {currentBatch} returned empty response", "warn")
                Return False
            End If

            ' Parse response
            APParseResponse(llmResponse, processable, batchStart, batchEnd)

            batchIndex = batchEnd + 1
        End While

        ApDashboardLog($"DocProcessor: all {currentBatch} batches completed", "success")
        Return True
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  RESPONSE PARSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses LLM output and assigns processed text to paragraph entries.
    ''' </summary>
    Private Sub APParseResponse(response As String, paragraphs As List(Of APParagraphInfo),
                                 batchStart As Integer, batchEnd As Integer)
        Dim pattern As New Regex("\[(\d+)\]\s*(.*?)(?=\s*\[\d+\]|$)", RegexOptions.Singleline)
        Dim matches = pattern.Matches(response)

        For Each m As Match In matches
            Dim num As Integer
            If Integer.TryParse(m.Groups(1).Value, num) Then
                Dim absoluteIndex = batchStart + num - 1
                If absoluteIndex >= batchStart AndAlso absoluteIndex <= batchEnd AndAlso absoluteIndex < paragraphs.Count Then
                    Dim processed = m.Groups(2).Value.Trim()
                    processed = Regex.Replace(processed, "^\[/?TEXTTOPROCESS\]", "", RegexOptions.IgnoreCase).Trim()
                    processed = Regex.Replace(processed, "\[/?TEXTTOPROCESS\]$", "", RegexOptions.IgnoreCase).Trim()
                    paragraphs(absoluteIndex).TranslatedText = processed
                End If
            End If
        Next
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  APPLYING TRANSLATIONS BACK TO XML
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Applies translated text back into the XML text nodes.
    ''' Partitions runs at footnote/endnote reference boundaries so that text
    ''' redistribution never moves content across a reference anchor.
    ''' When the LLM preserves ‖ markers, splits directly at their position (they are
    ''' authoritative). Falls back to anchor + word-boundary logic only when ‖ markers
    ''' are missing. The ‖ character is purely a positional placeholder — it indicates
    ''' where a footnote/endnote/field reference sits but does not create an additional
    ''' formatting split beyond the existing run partition.
    ''' </summary>
    Private Sub APApplyTranslations(paragraphs As List(Of APParagraphInfo))
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim translatedText = para.TranslatedText

            ' Single run: simple replacement
            If para.TextRuns.Count = 1 Then
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "").Replace(AP_NoteRefMarker, "")
                APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                Continue For
            End If

            ' ─── Check if any note-reference boundaries exist ───
            Dim hasNoteRefBoundaries As Boolean = para.TextRuns.Any(Function(r) r.HasNoteReferenceBefore)

            ' Try marker-based distribution if formatting markers were used
            ' BUT only when NO footnote boundaries exist — otherwise the | markers
            ' are unaware of footnote positions and will shift text across them.
            If Not hasNoteRefBoundaries Then
                If APTryApplyMarkerBasedDistribution(para, translatedText) Then
                    Continue For
                End If
                translatedText = translatedText.Replace(AP_RunBoundaryMarker, "")
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")
                APDistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── From here: paragraph HAS footnote/endnote reference boundaries ───
            ' Strip formatting markers — the ‖ partitioning handles distribution
            translatedText = translatedText.Replace(AP_RunBoundaryMarker, "")

            ' Partition runs into segments at note-reference boundaries
            Dim segments As New List(Of List(Of Integer))()
            Dim currentSegment As New List(Of Integer)()

            For runIdx As Integer = 0 To para.TextRuns.Count - 1
                If para.TextRuns(runIdx).HasNoteReferenceBefore AndAlso currentSegment.Count > 0 Then
                    segments.Add(currentSegment)
                    currentSegment = New List(Of Integer)()
                End If
                currentSegment.Add(runIdx)
            Next
            If currentSegment.Count > 0 Then segments.Add(currentSegment)

            ' Safety: if partitioning produced only 1 segment, fall back to proportional
            If segments.Count <= 1 Then
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")
                APDistributeProportional(para, translatedText)
                Continue For
            End If

            ' ─── Split translated text across segments ───
            Dim expectedNoteRefCount As Integer = segments.Count - 1
            Dim actualNoteRefCount As Integer = 0
            For Each ch As Char In translatedText
                If ch = AP_NoteRefMarker(0) Then actualNoteRefCount += 1
            Next

            Dim segmentTexts As String() = Nothing

            If actualNoteRefCount = expectedNoteRefCount Then
                ' ─── LLM preserved ‖ markers: split directly on them ───
                ' The ‖ is a positional placeholder — it tells us where the footnote
                ' reference sits. We split the text at these positions and assign
                ' each piece to its corresponding segment of runs.
                segmentTexts = translatedText.Split(New String() {AP_NoteRefMarker}, StringSplitOptions.None)

                If segmentTexts.Length = segments.Count Then
                    ' Normalize spaces at segment boundaries: if the original text after
                    ' the footnote started with whitespace, ensure the space is on that
                    ' side (not trailing on the previous segment before the superscript).
                    For segIdx As Integer = 0 To segmentTexts.Length - 2
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            Dim curSeg As String = segmentTexts(segIdx)
                            If curSeg.Length > 0 AndAlso curSeg(curSeg.Length - 1) = " "c Then
                                Dim trimmed As String = curSeg.TrimEnd(" "c)
                                Dim movedSpaces As String = curSeg.Substring(trimmed.Length)
                                segmentTexts(segIdx) = trimmed
                                segmentTexts(segIdx + 1) = movedSpaces & segmentTexts(segIdx + 1)
                            End If
                        End If
                    Next
                Else
                    segmentTexts = Nothing
                End If
            End If

            If segmentTexts Is Nothing Then
                ' ─── Fallback: anchor-based splitting (no ‖ or wrong count) ───
                Debug.WriteLine($"Note-ref marker mismatch for paragraph {para.Index}: expected {expectedNoteRefCount}, got {actualNoteRefCount}. Using anchor fallback.")
                translatedText = translatedText.Replace(AP_NoteRefMarker, "")

                Dim totalOrigLen As Integer = para.FullText.Replace(AP_NoteRefMarker, "").Length
                Dim translatedLen As Integer = translatedText.Length

                If totalOrigLen = 0 Then
                    APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                    For idx As Integer = 1 To para.TextRuns.Count - 1
                        APSetTextNode(para.TextRuns(idx).TextNode, "")
                    Next
                    Continue For
                End If

                Dim segmentOrigTexts As New List(Of String)()
                For Each seg In segments
                    Dim sb As New StringBuilder()
                    For Each ri In seg
                        sb.Append(para.TextRuns(ri).OriginalText)
                    Next
                    segmentOrigTexts.Add(sb.ToString())
                Next

                Dim splitPoints As New List(Of Integer)()
                Dim searchFrom As Integer = 0

                For segIdx As Integer = 0 To segments.Count - 2
                    Dim segOrigText As String = segmentOrigTexts(segIdx)
                    Dim nextSegOrigText As String = segmentOrigTexts(segIdx + 1)
                    Dim cumulativeOrigLen As Integer = 0
                    For si As Integer = 0 To segIdx
                        cumulativeOrigLen += segmentOrigTexts(si).Length
                    Next
                    Dim proportion As Double = cumulativeOrigLen / CDbl(totalOrigLen)
                    Dim proportionalEnd As Integer = CInt(Math.Round(proportion * translatedLen))
                    proportionalEnd = Math.Min(proportionalEnd, translatedLen)

                    Dim splitPos As Integer = APFindNoteRefSplitPos(
                        translatedText, segOrigText, searchFrom, proportionalEnd, translatedLen, nextSegOrigText)
                    splitPos = Math.Min(splitPos, translatedLen)

                    If splitPos > searchFrom AndAlso splitPos <= translatedLen Then
                        Dim nextSegFirstRunIdx As Integer = segments(segIdx + 1)(0)
                        Dim nextRunOriginal As String = para.TextRuns(nextSegFirstRunIdx).OriginalText
                        If nextRunOriginal.Length > 0 AndAlso Char.IsWhiteSpace(nextRunOriginal(0)) Then
                            While splitPos > searchFrom AndAlso translatedText(splitPos - 1) = " "c
                                splitPos -= 1
                            End While
                        End If
                    End If

                    splitPoints.Add(splitPos)
                    searchFrom = splitPos
                Next
                splitPoints.Add(translatedLen)

                segmentTexts = New String(segments.Count - 1) {}
                Dim st As Integer = 0
                For i As Integer = 0 To segments.Count - 1
                    Dim en As Integer = splitPoints(i)
                    segmentTexts(i) = If(en > st, translatedText.Substring(st, en - st), "")
                    st = en
                Next
            End If

            ' ─── Distribute text within each segment independently ───
            ' Each segment is a group of runs between two footnote references.
            ' Within each segment, text is distributed proportionally across runs.
            For segIdx As Integer = 0 To segments.Count - 1
                Dim segText As String = segmentTexts(segIdx)
                Dim segRunIndices As List(Of Integer) = segments(segIdx)

                If segRunIndices.Count = 1 Then
                    APSetTextNode(para.TextRuns(segRunIndices(0)).TextNode, segText)
                Else
                    Dim segOrigLen As Integer = 0
                    For Each ri In segRunIndices
                        segOrigLen += para.TextRuns(ri).OriginalText.Length
                    Next
                    Dim segTransLen As Integer = segText.Length
                    Dim segCurrentPos As Integer = 0
                    Dim segCumOrig As Integer = 0

                    For ri As Integer = 0 To segRunIndices.Count - 1
                        Dim runIdx As Integer = segRunIndices(ri)
                        Dim run = para.TextRuns(runIdx)
                        segCumOrig += run.OriginalText.Length

                        If ri = segRunIndices.Count - 1 Then
                            Dim remaining As String = If(segCurrentPos < segTransLen,
                                                         segText.Substring(segCurrentPos), "")
                            APSetTextNode(run.TextNode, remaining)
                        Else
                            If segOrigLen = 0 Then
                                APSetTextNode(run.TextNode, "")
                                Continue For
                            End If

                            Dim prop As Double = segCumOrig / CDbl(segOrigLen)
                            Dim tgtEnd As Integer = CInt(Math.Round(prop * segTransLen))
                            tgtEnd = Math.Min(tgtEnd, segTransLen)

                            If tgtEnd <= segCurrentPos Then
                                APSetTextNode(run.TextNode, "")
                                Continue For
                            End If

                            Dim endPos As Integer = tgtEnd
                            Dim foundSpace As Boolean = False

                            If endPos < segTransLen AndAlso endPos > segCurrentPos Then
                                If segText(endPos) = " "c Then
                                    endPos += 1
                                    foundSpace = True
                                Else
                                    Dim searchMax As Integer = Math.Min(endPos + 15, segTransLen - 1)
                                    For sp As Integer = endPos To searchMax
                                        If segText(sp) = " "c Then
                                            endPos = sp + 1
                                            foundSpace = True
                                            Exit For
                                        End If
                                    Next
                                    If Not foundSpace Then
                                        For sp As Integer = endPos - 1 To segCurrentPos Step -1
                                            If segText(sp) = " "c Then
                                                endPos = sp + 1
                                                foundSpace = True
                                                Exit For
                                            End If
                                        Next
                                    End If
                                    If Not foundSpace Then
                                        Dim we As Integer = endPos
                                        While we < segTransLen AndAlso segText(we) <> " "c
                                            we += 1
                                        End While
                                        endPos = we
                                    End If
                                End If
                            End If

                            endPos = Math.Min(endPos, segTransLen)
                            APSetTextNode(run.TextNode, segText.Substring(segCurrentPos, endPos - segCurrentPos))
                            segCurrentPos = endPos
                        End If
                    Next
                End If
            Next

        Next
    End Sub

    ''' <summary>
    ''' Proportional distribution fallback when no note-reference boundaries exist.
    ''' </summary>
    Private Sub APDistributeProportional(para As APParagraphInfo, translatedText As String)
        Dim totalOrigLen As Integer = para.FullText.Length
        Dim translatedLen As Integer = translatedText.Length

        If totalOrigLen = 0 Then
            APSetTextNode(para.TextRuns(0).TextNode, translatedText)
            For idx As Integer = 1 To para.TextRuns.Count - 1
                APSetTextNode(para.TextRuns(idx).TextNode, "")
            Next
            Return
        End If

        Dim currentPos As Integer = 0
        Dim cumulativeOriginal As Integer = 0

        For runIdx As Integer = 0 To para.TextRuns.Count - 1
            Dim run = para.TextRuns(runIdx)
            cumulativeOriginal += run.OriginalText.Length

            If runIdx = para.TextRuns.Count - 1 Then
                Dim remaining = If(currentPos < translatedLen, translatedText.Substring(currentPos), "")
                APSetTextNode(run.TextNode, remaining)
            Else
                Dim proportion As Double = cumulativeOriginal / CDbl(totalOrigLen)
                Dim targetEndPos = CInt(Math.Round(proportion * translatedLen))
                targetEndPos = Math.Min(targetEndPos, translatedLen)

                If targetEndPos <= currentPos Then
                    APSetTextNode(run.TextNode, "")
                    Continue For
                End If

                Dim endPos = targetEndPos
                Dim foundSpace = False

                If endPos < translatedLen AndAlso endPos > currentPos Then
                    If translatedText(endPos) = " "c Then
                        endPos += 1
                        foundSpace = True
                    Else
                        Dim searchMax As Integer = Math.Min(endPos + 15, translatedLen - 1)
                        For searchPos = endPos To searchMax
                            If translatedText(searchPos) = " "c Then
                                endPos = searchPos + 1
                                foundSpace = True
                                Exit For
                            End If
                        Next
                        If Not foundSpace Then
                            For searchPos = endPos - 1 To currentPos Step -1
                                If translatedText(searchPos) = " "c Then
                                    endPos = searchPos + 1
                                    foundSpace = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not foundSpace Then
                            Dim wordEnd = endPos
                            While wordEnd < translatedLen AndAlso translatedText(wordEnd) <> " "c
                                wordEnd += 1
                            End While
                            If wordEnd < translatedLen AndAlso translatedText(wordEnd) = " "c Then
                                endPos = wordEnd + 1
                            Else
                                endPos = wordEnd
                            End If
                        End If
                    End If
                End If

                endPos = Math.Min(endPos, translatedLen)
                APSetTextNode(run.TextNode, translatedText.Substring(currentPos, endPos - currentPos))
                currentPos = endPos
            End If
        Next
    End Sub

    ''' <summary>
    ''' Finds the best split position in the translated text for a note-reference boundary
    ''' by anchoring on the last word(s) of the original segment text, or the first word(s)
    ''' of the next segment text. Falls back to proportional position with backward
    ''' word-boundary seeking if both anchoring strategies fail.
    ''' </summary>
    ''' <param name="translatedText">The full translated paragraph text.</param>
    ''' <param name="originalSegText">The original text of the segment before the note reference.</param>
    ''' <param name="currentPos">Start of the current unassigned region in translatedText.</param>
    ''' <param name="targetEndPos">Proportional end position (fallback).</param>
    ''' <param name="translatedLength">Total length of translatedText.</param>
    ''' <param name="nextSegOrigText">The original text of the segment after the note reference (Nothing if unavailable).</param>
    ''' <returns>The character index in translatedText where the split should occur.</returns>
    Private Shared Function APFindNoteRefSplitPos(
            translatedText As String,
            originalSegText As String,
            currentPos As Integer,
            targetEndPos As Integer,
            translatedLength As Integer,
            Optional nextSegOrigText As String = Nothing) As Integer

        ' ── Strategy 1: Anchor on the LAST word(s) of the current segment ──
        If Not String.IsNullOrWhiteSpace(originalSegText) AndAlso originalSegText.Length >= 2 Then
            Dim origTrimmed As String = originalSegText.TrimEnd()
            If origTrimmed.Length >= 2 Then
                Dim words As String() = origTrimmed.Split(" "c)
                Dim maxWords As Integer = Math.Min(words.Length, 3)

                For tryWords As Integer = maxWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", words, words.Length - tryWords, tryWords)
                    If fragment.Length < 2 Then Continue For

                    Dim searchRegion As String = translatedText.Substring(currentPos)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim splitPos As Integer = currentPos + fragIdx + fragment.Length
                        If splitPos >= translatedLength OrElse
                           translatedText(splitPos) = " "c OrElse
                           Char.IsPunctuation(translatedText(splitPos)) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Strategy 2: Anchor on the FIRST word(s) of the NEXT segment ──
        If Not String.IsNullOrWhiteSpace(nextSegOrigText) AndAlso nextSegOrigText.Length >= 2 Then
            Dim nextTrimmed As String = nextSegOrigText.TrimStart()
            If nextTrimmed.Length >= 2 Then
                Dim nextWords As String() = nextTrimmed.Split(" "c)
                Dim maxNextWords As Integer = Math.Min(nextWords.Length, 3)

                For tryWords As Integer = maxNextWords To 1 Step -1
                    Dim fragment As String = String.Join(" ", nextWords, 0, tryWords)
                    If fragment.Length < 2 Then Continue For

                    Dim earliestSearch As Integer = Math.Max(currentPos, CInt(currentPos + (targetEndPos - currentPos) * 0.4))
                    If earliestSearch >= translatedLength Then Continue For

                    Dim searchRegion As String = translatedText.Substring(earliestSearch)
                    Dim fragIdx As Integer = searchRegion.IndexOf(fragment, StringComparison.OrdinalIgnoreCase)

                    If fragIdx >= 0 Then
                        Dim matchStart As Integer = earliestSearch + fragIdx
                        Dim splitPos As Integer = matchStart
                        While splitPos > currentPos AndAlso translatedText(splitPos - 1) = " "c
                            splitPos -= 1
                        End While
                        If splitPos <= currentPos Then Continue For
                        If splitPos >= translatedLength OrElse
                           Char.IsWhiteSpace(translatedText(splitPos)) OrElse
                           (splitPos > 0 AndAlso (Char.IsPunctuation(translatedText(splitPos - 1)) OrElse
                                                   Char.IsWhiteSpace(translatedText(splitPos - 1)))) Then
                            Return splitPos
                        End If
                    End If
                Next
            End If
        End If

        ' ── Fallback: backward word-boundary search from proportional position ──
        Dim endPos As Integer = Math.Min(targetEndPos, translatedLength)
        If endPos > currentPos AndAlso endPos < translatedLength Then
            If translatedText(endPos) = " "c Then Return endPos
            For searchPos As Integer = endPos - 1 To currentPos Step -1
                If translatedText(searchPos) = " "c Then
                    Return searchPos + 1
                End If
            Next
        End If

        Return endPos
    End Function

    ''' <summary>
    ''' Attempts marker-based distribution (matching Word add-in's TryApplyMarkerBasedDistribution).
    ''' If the LLM returned text with exactly the right number of | markers, split on them
    ''' and assign each segment to the corresponding run. Returns True if successful.
    ''' </summary>
    Private Function APTryApplyMarkerBasedDistribution(para As APParagraphInfo, translatedText As String) As Boolean
        ' Only applicable if the paragraph was sent with markers
        If para.MarkerText Is Nothing Then Return False

        ' Count expected markers from the original marker text
        Dim expectedMarkers As Integer = 0
        For Each ch As Char In para.MarkerText
            If ch = AP_RunBoundaryMarker(0) Then expectedMarkers += 1
        Next

        If expectedMarkers = 0 Then Return False

        ' Count actual markers in translated text
        Dim actualMarkers = translatedText.Count(Function(c) c = AP_RunBoundaryMarker(0))
        If actualMarkers <> expectedMarkers Then
            Debug.WriteLine($"Marker count mismatch for paragraph {para.Index}: expected {expectedMarkers}, got {actualMarkers}. Falling back to proportional.")
            Return False
        End If

        ' Split on the marker
        Dim segments = translatedText.Split(New String() {AP_RunBoundaryMarker}, StringSplitOptions.None)

        ' Map segments back to non-empty runs
        Dim nonEmptyRunIndices As New List(Of Integer)()
        For i As Integer = 0 To para.TextRuns.Count - 1
            If para.TextRuns(i).OriginalText.Length > 0 Then
                nonEmptyRunIndices.Add(i)
            End If
        Next

        If segments.Length <> nonEmptyRunIndices.Count Then
            Debug.WriteLine($"Segment count {segments.Length} <> non-empty run count {nonEmptyRunIndices.Count} for paragraph {para.Index}. Falling back.")
            Return False
        End If

        ' Apply segments to non-empty runs, clear empty runs
        Dim segmentIdx As Integer = 0
        For runIdx As Integer = 0 To para.TextRuns.Count - 1
            Dim run = para.TextRuns(runIdx)
            If run.OriginalText.Length = 0 Then
                APSetTextNode(run.TextNode, "")
            Else
                If segmentIdx < segments.Length Then
                    APSetTextNode(run.TextNode, segments(segmentIdx))
                    segmentIdx += 1
                Else
                    APSetTextNode(run.TextNode, "")
                End If
            End If
        Next

        Return True
    End Function

    ''' <summary>
    ''' Sets a w:t text node and ensures xml:space="preserve" when needed.
    ''' </summary>
    Private Sub APSetTextNode(textNode As System.Xml.XmlNode, text As String)
        textNode.InnerText = text
        If text.Length > 0 AndAlso (text.StartsWith(" ") OrElse text.EndsWith(" ") OrElse text.Contains("  ")) Then
            Dim xmlSpaceAttr = textNode.Attributes("xml:space")
            If xmlSpaceAttr Is Nothing Then
                xmlSpaceAttr = textNode.OwnerDocument.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
                textNode.Attributes.Append(xmlSpaceAttr)
            End If
            xmlSpaceAttr.Value = "preserve"
        End If
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SUB-PART PROCESSING (headers, footers, comments, footnotes, endnotes)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes header, footer, comment, footnote, and endnote parts in a DOCX package.
    ''' </summary>
    Private Async Function APProcessSubParts(tempDir As String, instruction As String, ct As CancellationToken) As Task
        Dim wordDir = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Return

        ' Process headers and footers
        For Each pattern In {"header*.xml", "footer*.xml"}
            For Each filePath In Directory.GetFiles(wordDir, pattern)
                Try
                    Await APProcessXmlFile(filePath, instruction, ct)
                Catch ex As System.Exception
                    Debug.WriteLine("APProcessSubParts error for " & Path.GetFileName(filePath) & ": " & ex.Message)
                End Try
            Next
        Next

        ' Process comments, footnotes, endnotes
        For Each fileName In {"comments.xml", "footnotes.xml", "endnotes.xml"}
            Dim filePath = Path.Combine(wordDir, fileName)
            If File.Exists(filePath) Then
                Try
                    Await APProcessXmlFile(filePath, instruction, ct)
                Catch ex As System.Exception
                    Debug.WriteLine("APProcessSubParts error for " & fileName & ": " & ex.Message)
                End Try
            End If
        Next
    End Function

    ''' <summary>Processes a single XML file (header/footer/comments/etc.).</summary>
    Private Async Function APProcessXmlFile(filePath As String, instruction As String, ct As CancellationToken) As Task
        Dim xmlDoc As New System.Xml.XmlDocument()
        xmlDoc.PreserveWhitespace = True
        xmlDoc.Load(filePath)

        Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
        nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

        Dim paragraphs = APExtractParagraphs(xmlDoc, nsMgr)
        Dim processable = paragraphs.Where(Function(p) Not p.IsEmpty).ToList()
        If processable.Count = 0 Then Return

        Dim success = Await APProcessBatches(paragraphs, instruction, ct)
        If success Then
            APApplyTranslations(paragraphs)
            xmlDoc.Save(filePath)
        End If
    End Function

End Class