' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.PdfRedactor.vb
' Purpose:
'   AutoPilot tool implementation for LLM-driven PDF redaction (prepare + finalize).
'   Reuses the same extraction/matching/annotation logic as the Word add-in's
'   PdfRedactionService, adapted for headless AutoPilot tool execution.
'
' Architecture:
'  - Text Extraction: PdfPig-based character-level extraction (same as Word add-in).
'  - LLM Invocation: Calls LLM() with SP_Redact system prompt and user instruction.
'  - Text Matching: Exact + normalized whitespace matching for LLM-provided snippets.
'  - Annotation Creation: /Square PDF annotations via PdfSharp (removable red/black boxes).
'  - Finalization: Burns annotations into rasterized images via PdfiumViewer.
'  - Three modes: prepare-only, prepare+finalize, finalize-only.
'
' Dependencies:
'  - UglyToad.PdfPig for text extraction.
'  - PdfSharp for annotation creation and PDF manipulation.
'  - PdfiumViewer for rasterization during finalization.
'  - SharedLibrary (LLM, SanitizeLlmResult, SP_Redact, InterpolateAtRuntime).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports PDFP = UglyToad.PdfPig

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — DATA CLASSES
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Class APRedactTextPosition
        Public Property PageNumber As Integer
        Public Property X As Double
        Public Property Y As Double
        Public Property Width As Double
        Public Property Height As Double
        Public Property PageHeight As Double
    End Class

    Private Class APRedactRangeDto
        <JsonProperty("start")>
        Public Property Start As Integer
        <JsonProperty("end")>
        Public Property [End] As Integer
    End Class

    Private Class APRedactItemDto
        <JsonProperty("id")>
        Public Property Id As String
        <JsonProperty("reason")>
        Public Property Reason As String
        <JsonProperty("exact_text")>
        Public Property ExactText As String
        <JsonProperty("ranges")>
        Public Property Ranges As List(Of APRedactRangeDto)
    End Class

    Private Class APRedactResponseDto
        <JsonProperty("redactions")>
        Public Property Redactions As List(Of APRedactItemDto)
    End Class

    Private Class APRedactRectangle
        Public Property PageNumber As Integer
        Public Property X As Double
        Public Property Y As Double
        Public Property Width As Double
        Public Property Height As Double
        Public Property ReasonCode As String
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AP_RedactionPadPts As Double = 2.5
    Private Const AP_RedactionSidePad As Double = 2.0

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — MAIN ENTRY POINT (called by tool executor)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Prepares PDF redactions by extracting text, querying the LLM, and adding removable
    ''' annotation boxes. Optionally finalizes (burns in) the redactions in a second step.
    ''' </summary>
    ''' <param name="inputPath">Input PDF file path.</param>
    ''' <param name="outputPath">Output PDF file path for the redacted result.</param>
    ''' <param name="instruction">Redaction instruction for the LLM.</param>
    ''' <param name="finalize">If True, also burn annotations into rasterized images.</param>
    ''' <param name="includeReasonCodes">Include reason codes in annotation /Contents field.</param>
    ''' <param name="ct">Cancellation token.</param>
    ''' <returns>A result message string; Nothing on failure (message logged).</returns>
    Private Async Function APRedactPdf(inputPath As String, outputPath As String,
                                        instruction As String,
                                        finalize As Boolean,
                                        includeReasonCodes As Boolean,
                                        ct As CancellationToken) As Task(Of String)

        ct.ThrowIfCancellationRequested()

        ' ── Step 1: Extract text and character positions ──
        Dim fullText As String = Nothing
        Dim positions As List(Of APRedactTextPosition) = Nothing
        APRedactExtractTextAndPositions(inputPath, fullText, positions)

        If String.IsNullOrWhiteSpace(fullText) OrElse positions Is Nothing OrElse positions.Count = 0 Then
            Return Nothing ' Caller will report "no extractable text"
        End If

        ApDashboardLog($"PdfRedactor: extracted {fullText.Length} characters", "step")

        ' ── Step 2: Call LLM with redaction prompt ──
        OtherPrompt = instruction
        Dim systemPrompt As String = InterpolateAtRuntime(SP_Redact)

        Dim userPrompt As String = "<TEXTTOPROCESS>" & fullText & "</TEXTTOPROCESS>"

        ApDashboardLog("PdfRedactor: calling LLM for redaction analysis", "step")

        Dim llmJson = Await LLM(systemPrompt, userPrompt,
                                 UseSecondAPI:=If(_apConfig IsNot Nothing, _apConfig.UseSecondApi, False),
                                 HideSplash:=True, EnsureUI:=False,
                                 cancellationToken:=ct)

        llmJson = WebAgentInterpreter.SanitizeLlmResult(llmJson)

        If String.IsNullOrWhiteSpace(llmJson) Then
            Return Nothing ' Caller will report "empty response"
        End If

        ' ── Step 3: Parse LLM JSON response ──
        Dim response As APRedactResponseDto = Nothing
        Try
            response = JsonConvert.DeserializeObject(Of APRedactResponseDto)(llmJson)
        Catch ex As Exception
            ApDashboardLog($"PdfRedactor: JSON parse error: {ex.Message}", "warn")
            Return Nothing
        End Try

        If response Is Nothing OrElse response.Redactions Is Nothing OrElse response.Redactions.Count = 0 Then
            Return "no_redactions" ' LLM found nothing to redact
        End If

        ' ── Step 4: Convert exact_text to ranges ──
        For Each item In response.Redactions
            If Not String.IsNullOrEmpty(item.ExactText) Then
                item.Ranges = APRedactFindTextOccurrences(fullText, item.ExactText)
            End If
        Next

        ' Filter out items with no valid ranges
        response.Redactions = response.Redactions.Where(
            Function(r) r.Ranges IsNot Nothing AndAlso r.Ranges.Count > 0).ToList()

        If response.Redactions.Count = 0 Then
            Return Nothing ' No matching text found
        End If

        ' ── Step 5: Build redaction rectangles ──
        Dim rectangles = APRedactBuildRectangles(response, positions, fullText.Length)
        If rectangles Is Nothing OrElse rectangles.Count = 0 Then
            Return Nothing
        End If

        ' ── Step 6: Create annotated PDF ──
        Dim annotatedPath As String
        If finalize Then
            ' Write annotations to a temp file first, then burn in
            annotatedPath = outputPath & ".tmp_annotated.pdf"
        Else
            annotatedPath = outputPath
        End If

        APRedactCreateAnnotatedPdf(inputPath, annotatedPath, rectangles, includeReasonCodes,
                                    transparent:=Not finalize, removeMetadata:=False)

        Dim redactionCount = response.Redactions.Count
        Dim boxCount = rectangles.Count

        ' ── Step 7: Finalize if requested ──
        If finalize Then
            Try
                APRedactBurnIn(annotatedPath, outputPath, includeReasonCodes, dpi:=300)
            Finally
                Try : If File.Exists(annotatedPath) Then File.Delete(annotatedPath)
                Catch : End Try
            End Try
        End If

        Dim modeLabel = If(finalize, "prepared and finalized (burned in)", "prepared (removable annotation boxes)")
        Return $"{redactionCount} redaction(s) identified, {boxCount} box(es) {modeLabel}."
    End Function

    ''' <summary>
    ''' Finalizes (burns in) an existing PDF that already has redaction annotations.
    ''' Used for the finalize-only mode of the redact_pdf tool.
    ''' </summary>
    Private Sub APRedactFinalizeOnly(inputPath As String, outputPath As String,
                                      includeReasonCodes As Boolean, dpi As Integer)
        APRedactBurnIn(inputPath, outputPath, includeReasonCodes, dpi)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — TEXT EXTRACTION
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Sub APRedactExtractTextAndPositions(pdfPath As String,
                           ByRef fullText As String,
                           ByRef positions As List(Of APRedactTextPosition))

        Dim sb As New StringBuilder()
        positions = New List(Of APRedactTextPosition)()

        Using document As PDFP.PdfDocument = PDFP.PdfDocument.Open(pdfPath)
            Dim pageCount As Integer = document.NumberOfPages

            For pageNumber As Integer = 1 To pageCount
                Dim page As PDFP.Content.Page = document.GetPage(pageNumber)
                Dim pageHeight As Double = page.Height

                Dim words = page.GetWords().OrderBy(Function(w) w.BoundingBox.Bottom).
                    ThenBy(Function(w) w.BoundingBox.Left).ToList()

                Dim previousWord As PDFP.Content.Word = Nothing
                Dim lineThreshold As Double = 2.0

                For Each word As PDFP.Content.Word In words
                    Dim wordText As String = word.Text
                    If String.IsNullOrEmpty(wordText) Then Continue For

                    If previousWord IsNot Nothing Then
                        Dim sameLine As Boolean = Math.Abs(word.BoundingBox.Bottom -
                            previousWord.BoundingBox.Bottom) < lineThreshold

                        If sameLine Then
                            Dim spaceWidth As Double = word.BoundingBox.Left - previousWord.BoundingBox.Right
                            If spaceWidth > 0 Then
                                sb.Append(" ")
                                positions.Add(New APRedactTextPosition() With {
                                    .PageNumber = pageNumber,
                                    .X = previousWord.BoundingBox.Right,
                                    .Y = word.BoundingBox.Bottom,
                                    .Width = spaceWidth,
                                    .Height = word.BoundingBox.Height,
                                    .PageHeight = pageHeight
                                })
                            End If
                        Else
                            sb.Append(" ")
                            positions.Add(New APRedactTextPosition() With {
                                .PageNumber = pageNumber,
                                .X = 0, .Y = 0, .Width = 0, .Height = 0,
                                .PageHeight = pageHeight
                            })
                        End If
                    End If

                    Dim bbox As UglyToad.PdfPig.Core.PdfRectangle = word.BoundingBox
                    Dim charWidth As Double = bbox.Width / wordText.Length

                    For i As Integer = 0 To wordText.Length - 1
                        sb.Append(wordText(i))
                        positions.Add(New APRedactTextPosition() With {
                            .PageNumber = pageNumber,
                            .X = bbox.Left + (i * charWidth),
                            .Y = bbox.Bottom,
                            .Width = charWidth,
                            .Height = bbox.Height,
                            .PageHeight = pageHeight
                        })
                    Next

                    previousWord = word
                Next

                If pageNumber < pageCount Then
                    sb.Append(" ")
                    positions.Add(New APRedactTextPosition() With {
                        .PageNumber = pageNumber,
                        .X = 0, .Y = 0, .Width = 0, .Height = 0,
                        .PageHeight = pageHeight
                    })
                End If
            Next
        End Using

        fullText = sb.ToString()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — TEXT MATCHING
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function APRedactFindTextOccurrences(fullText As String,
                                                         searchText As String) As List(Of APRedactRangeDto)
        Dim ranges As New List(Of APRedactRangeDto)()
        If String.IsNullOrEmpty(searchText) Then Return ranges

        ' Try exact match first
        Dim index As Integer = 0
        While index < fullText.Length
            index = fullText.IndexOf(searchText, index, StringComparison.Ordinal)
            If index = -1 Then Exit While
            ranges.Add(New APRedactRangeDto() With {.Start = index, .[End] = index + searchText.Length})
            index += searchText.Length
        End While

        ' Fallback: normalized whitespace matching
        If ranges.Count = 0 Then
            Dim normalizedFull As String = System.Text.RegularExpressions.Regex.Replace(fullText, "\s+", " ")
            Dim normalizedSearch As String = System.Text.RegularExpressions.Regex.Replace(searchText, "\s+", " ")

            Dim normToOrigMap As New List(Of Integer)()
            Dim origIdx As Integer = 0

            For normIdx As Integer = 0 To normalizedFull.Length - 1
                While origIdx < fullText.Length AndAlso Char.IsWhiteSpace(fullText(origIdx)) AndAlso
                      (normIdx = 0 OrElse Not Char.IsWhiteSpace(normalizedFull(normIdx - 1)) OrElse
                       Not Char.IsWhiteSpace(fullText(origIdx - 1)))
                    origIdx += 1
                End While
                If origIdx < fullText.Length Then
                    normToOrigMap.Add(origIdx)
                    origIdx += 1
                End If
            Next

            index = 0
            While index < normalizedFull.Length
                index = normalizedFull.IndexOf(normalizedSearch, index, StringComparison.Ordinal)
                If index = -1 Then Exit While

                If index < normToOrigMap.Count AndAlso index + normalizedSearch.Length <= normToOrigMap.Count Then
                    Dim origStart As Integer = normToOrigMap(index)
                    Dim testText As String = fullText.Substring(origStart,
                        Math.Min(searchText.Length + 10, fullText.Length - origStart))
                    Dim actualMatch As String = System.Text.RegularExpressions.Regex.Replace(testText, "\s+", " ")

                    If actualMatch.StartsWith(normalizedSearch, StringComparison.Ordinal) Then
                        Dim pos As Integer = origStart
                        For Each c In normalizedSearch
                            If Not Char.IsWhiteSpace(c) Then
                                While pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos))
                                    pos += 1
                                End While
                                pos += 1
                            Else
                                If pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos)) Then
                                    While pos < fullText.Length AndAlso Char.IsWhiteSpace(fullText(pos))
                                        pos += 1
                                    End While
                                End If
                            End If
                        Next
                        ranges.Add(New APRedactRangeDto() With {.Start = origStart, .[End] = pos})
                    End If
                End If

                index += normalizedSearch.Length
            End While
        End If

        Return ranges
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — RECTANGLE BUILDING
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function APRedactBuildRectangles(response As APRedactResponseDto,
                                                     positions As List(Of APRedactTextPosition),
                                                     textLength As Integer) As List(Of APRedactRectangle)
        Dim rectangles As New List(Of APRedactRectangle)()
        If response Is Nothing OrElse response.Redactions Is Nothing Then Return rectangles

        Dim maxValidIndex As Integer = Math.Min(textLength, positions.Count)

        For Each item In response.Redactions
            If item Is Nothing OrElse item.Ranges Is Nothing Then Continue For

            For Each r In item.Ranges
                If r Is Nothing OrElse r.Start < 0 OrElse r.[End] <= r.Start OrElse r.[End] > maxValidIndex Then
                    Continue For
                End If

                Dim pageGroups As New Dictionary(Of Integer, List(Of APRedactTextPosition))()
                For i As Integer = r.Start To r.[End] - 1
                    Dim pos = positions(i)
                    If pos IsNot Nothing AndAlso pos.Width > 0 Then
                        If Not pageGroups.ContainsKey(pos.PageNumber) Then
                            pageGroups(pos.PageNumber) = New List(Of APRedactTextPosition)()
                        End If
                        pageGroups(pos.PageNumber).Add(pos)
                    End If
                Next

                For Each kvp In pageGroups
                    Dim pagePositions = kvp.Value
                    If pagePositions.Count = 0 Then Continue For

                    Dim lineGroups As New Dictionary(Of Integer, List(Of APRedactTextPosition))()
                    For Each pos In pagePositions
                        Dim lineKey As Integer = CInt(Math.Round(pos.Y / 5.0) * 5)
                        If Not lineGroups.ContainsKey(lineKey) Then
                            lineGroups(lineKey) = New List(Of APRedactTextPosition)()
                        End If
                        lineGroups(lineKey).Add(pos)
                    Next

                    For Each lineKvp In lineGroups
                        Dim linePositions = lineKvp.Value
                        Dim minX As Double = linePositions.Min(Function(p) p.X)
                        Dim maxX As Double = linePositions.Max(Function(p) p.X + p.Width)
                        Dim avgY As Double = linePositions.Average(Function(p) p.Y)
                        Dim maxHeight As Double = linePositions.Max(Function(p) p.Height)
                        Dim pageHeight As Double = linePositions.First().PageHeight

                        ' Convert from PdfPig (bottom-left) to PdfSharp (top-left)
                        Dim rect As New APRedactRectangle()
                        rect.PageNumber = kvp.Key
                        rect.X = minX
                        rect.Y = pageHeight - (avgY + maxHeight)
                        rect.Width = maxX - minX
                        rect.Height = maxHeight

                        ' Vertical padding
                        Dim newY As Double = Math.Max(0, rect.Y - AP_RedactionPadPts)
                        Dim bottom As Double = rect.Y + rect.Height + AP_RedactionPadPts
                        rect.Y = newY
                        rect.Height = Math.Max(0, Math.Min(pageHeight, bottom) - newY)

                        ' Horizontal padding
                        Dim newX As Double = Math.Max(0, rect.X - AP_RedactionSidePad)
                        Dim right As Double = rect.X + rect.Width + AP_RedactionSidePad
                        rect.X = newX
                        rect.Width = Math.Max(0, right - newX)

                        rect.ReasonCode = If(Not String.IsNullOrWhiteSpace(item.Reason), item.Reason, item.Id)
                        rectangles.Add(rect)
                    Next
                Next
            Next
        Next

        Return rectangles
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — ANNOTATION CREATION
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Sub APRedactCreateAnnotatedPdf(inputPath As String, outputPath As String,
                                                   rectangles As List(Of APRedactRectangle),
                                                   includeReasonCode As Boolean,
                                                   transparent As Boolean,
                                                   removeMetadata As Boolean)

        Dim document As PdfSharp.Pdf.PdfDocument =
            PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)

        Dim rectsByPage = rectangles.GroupBy(Function(r) r.PageNumber)

        For Each pageGroup In rectsByPage
            Dim pageIndex As Integer = pageGroup.Key - 1
            If pageIndex < 0 OrElse pageIndex >= document.Pages.Count Then Continue For

            Dim page As PdfSharp.Pdf.PdfPage = document.Pages(pageIndex)
            Dim pageHeightPts As Double = page.Height

            For Each r In pageGroup
                Dim yBottom As Double = pageHeightPts - (r.Y + r.Height)
                Dim yTop As Double = pageHeightPts - r.Y

                Dim rectArr As New PdfSharp.Pdf.PdfArray(document)
                rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(r.X))
                rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(yBottom))
                rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(r.X + r.Width))
                rectArr.Elements.Add(New PdfSharp.Pdf.PdfReal(yTop))

                Dim ann As New PdfSharp.Pdf.PdfDictionary(document)
                ann.Elements.SetName("/Type", "/Annot")
                ann.Elements.SetName("/Subtype", "/Square")
                ann.Elements("/Rect") = rectArr
                ann.Elements.SetInteger("/F", 4)

                If includeReasonCode AndAlso Not String.IsNullOrWhiteSpace(r.ReasonCode) Then
                    ann.Elements.SetString("/Contents", r.ReasonCode)
                End If

                Dim bs As New PdfSharp.Pdf.PdfDictionary(document)
                If transparent Then
                    bs.Elements.SetInteger("/W", 1)
                    ann.Elements("/BS") = bs
                    ann.Elements("/C") = APRedactMakeRgbArray(document, 1.0, 0.0, 0.0)
                Else
                    bs.Elements.SetInteger("/W", 0)
                    ann.Elements("/BS") = bs
                    ann.Elements("/C") = APRedactMakeRgbArray(document, 0.0, 0.0, 0.0)
                    ann.Elements("/IC") = APRedactMakeRgbArray(document, 0.0, 0.0, 0.0)
                End If

                Dim annots As PdfSharp.Pdf.PdfArray = page.Elements.GetArray("/Annots")
                If annots Is Nothing Then
                    annots = New PdfSharp.Pdf.PdfArray(document)
                    page.Elements("/Annots") = annots
                End If
                document.Internals.AddObject(ann)
                annots.Elements.Add(ann)
            Next
        Next

        If removeMetadata Then
            APRedactStripMetadata(document)
        End If

        document.Save(outputPath)
        document.Close()
    End Sub

    Private Shared Function APRedactMakeRgbArray(doc As PdfSharp.Pdf.PdfDocument,
                                                  r As Double, g As Double, b As Double) As PdfSharp.Pdf.PdfArray
        Dim arr As New PdfSharp.Pdf.PdfArray(doc)
        arr.Elements.Add(New PdfSharp.Pdf.PdfReal(r))
        arr.Elements.Add(New PdfSharp.Pdf.PdfReal(g))
        arr.Elements.Add(New PdfSharp.Pdf.PdfReal(b))
        Return arr
    End Function

    Private Shared Sub APRedactStripMetadata(doc As PdfSharp.Pdf.PdfDocument)
        If doc Is Nothing Then Return
        Try
            Dim info = doc.Info
            If info IsNot Nothing AndAlso info.Elements IsNot Nothing Then
                Dim keys As New List(Of String)
                For Each k As String In info.Elements.Keys
                    keys.Add(k)
                Next
                For Each k In keys
                    info.Elements.Remove(k)
                Next
            End If
            Dim catalog = doc.Internals.Catalog
            If catalog IsNot Nothing AndAlso catalog.Elements.ContainsKey("/Metadata") Then
                catalog.Elements.Remove("/Metadata")
            End If
            For Each p As PdfSharp.Pdf.PdfPage In doc.Pages
                If p.Elements.ContainsKey("/Metadata") Then p.Elements.Remove("/Metadata")
            Next
        Catch
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — BURN-IN (FINALIZATION)
    ' ═══════════════════════════════════════════════════════════════════════════

    <DllImport("kernel32.dll", EntryPoint:="LoadLibraryW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function APRedact_LoadLibrary(lpFileName As String) As IntPtr
    End Function

    <DllImport("kernel32.dll", EntryPoint:="SetDllDirectoryW", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function APRedact_SetDllDirectory(lpPathName As String) As Boolean
    End Function

    Private Shared _apRedactPdfiumLoaded As Boolean = False

    ''' <summary>
    ''' Ensures pdfium.dll is pre-loaded via LoadLibrary before any PdfiumViewer
    ''' type is accessed by the JIT compiler.
    ''' CRITICAL: This method must NEVER reference any PdfiumViewer type.
    ''' </summary>
    Private Shared Sub APRedactEnsurePdfiumLoaded()
        If _apRedactPdfiumLoaded Then Return
        _apRedactPdfiumLoaded = True  ' Set first to prevent re-entry

        Try
            ' Determine the platform-specific subdirectory
            Dim subDir As String = If(IntPtr.Size = 8, "x64", "x86")

            ' Strategy 1: Next to OUR assembly (most reliable for VSTO add-ins)
            ' For VSTO, Assembly.Location is the actual deployed add-in directory,
            ' NOT the Outlook.exe directory. This is where NuGet copies native DLLs.
            Dim asmLocation As String = GetType(ThisAddIn).Assembly.Location
            If Not String.IsNullOrEmpty(asmLocation) Then
                Dim asmDir As String = Path.GetDirectoryName(asmLocation)
                Debug.WriteLine($"APRedactEnsurePdfiumLoaded: Strategy 1 - Assembly.Location dir: {asmDir}")
                If TryLoadPdfiumFrom(asmDir, subDir) Then
                    Debug.WriteLine("APRedactEnsurePdfiumLoaded: SUCCESS via Assembly.Location")
                    Return
                End If
            End If

            ' Strategy 2: CodeBase URI (original location before shadow-copy, if applicable)
            Try
                Dim codeBase As String = GetType(ThisAddIn).Assembly.CodeBase
                If Not String.IsNullOrEmpty(codeBase) Then
                    Dim localPath As String = New Uri(codeBase).LocalPath
                    Dim codeDir As String = Path.GetDirectoryName(localPath)
                    Debug.WriteLine($"APRedactEnsurePdfiumLoaded: Strategy 2 - CodeBase dir: {codeDir}")
                    If Not String.IsNullOrEmpty(codeDir) Then
                        If TryLoadPdfiumFrom(codeDir, subDir) Then
                            Debug.WriteLine("APRedactEnsurePdfiumLoaded: SUCCESS via CodeBase")
                            Return
                        End If
                    End If
                End If
            Catch
            End Try

            ' Strategy 3: BaseDirectory (this is typically Outlook.exe's directory for VSTO,
            ' but may work for debug scenarios)
            Dim binDir As String = AppDomain.CurrentDomain.BaseDirectory
            Debug.WriteLine($"APRedactEnsurePdfiumLoaded: Strategy 3 - BaseDirectory: {binDir}")
            If TryLoadPdfiumFrom(binDir, subDir) Then
                Debug.WriteLine("APRedactEnsurePdfiumLoaded: SUCCESS via BaseDirectory")
                Return
            End If

            ' Strategy 4: RelativeSearchPath (probing path set by VSTO runtime)
            Dim relPath As String = AppDomain.CurrentDomain.RelativeSearchPath
            If Not String.IsNullOrEmpty(relPath) Then
                Debug.WriteLine($"APRedactEnsurePdfiumLoaded: Strategy 4 - RelativeSearchPath: {relPath}")
                For Each singlePath In relPath.Split(";"c)
                    If Not String.IsNullOrEmpty(singlePath) Then
                        If TryLoadPdfiumFrom(singlePath.Trim(), subDir) Then
                            Debug.WriteLine("APRedactEnsurePdfiumLoaded: SUCCESS via RelativeSearchPath")
                            Return
                        End If
                    End If
                Next
            End If

            ' Strategy 5: SetDllDirectory + add-in directory as fallback
            Dim fallbackDir As String = If(Not String.IsNullOrEmpty(asmLocation),
                Path.Combine(Path.GetDirectoryName(asmLocation), subDir),
                Path.Combine(binDir, subDir))
            Debug.WriteLine($"APRedactEnsurePdfiumLoaded: Strategy 5 - SetDllDirectory: {fallbackDir}")
            APRedact_SetDllDirectory(fallbackDir)

            Debug.WriteLine("APRedactEnsurePdfiumLoaded: WARNING - all strategies failed, pdfium.dll not pre-loaded")

        Catch ex As Exception
            Debug.WriteLine($"APRedactEnsurePdfiumLoaded failed: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Attempts to load pdfium.dll from a base directory, checking both the
    ''' platform subdirectory (x64/x86) and the directory root.
    ''' </summary>
    Private Shared Function TryLoadPdfiumFrom(baseDir As String, subDir As String) As Boolean
        If String.IsNullOrEmpty(baseDir) Then Return False

        ' Check platform subdirectory first (NuGet native package layout)
        Dim platformPath As String = Path.Combine(baseDir, subDir, "pdfium.dll")
        Debug.WriteLine($"  TryLoadPdfiumFrom: checking {platformPath} — exists: {File.Exists(platformPath)}")
        If File.Exists(platformPath) Then
            Dim h As IntPtr = APRedact_LoadLibrary(platformPath)
            Debug.WriteLine($"  TryLoadPdfiumFrom: LoadLibrary returned {h}")
            If h <> IntPtr.Zero Then Return True
        End If

        ' Check root directory
        Dim rootPath As String = Path.Combine(baseDir, "pdfium.dll")
        Debug.WriteLine($"  TryLoadPdfiumFrom: checking {rootPath} — exists: {File.Exists(rootPath)}")
        If File.Exists(rootPath) Then
            Dim h As IntPtr = APRedact_LoadLibrary(rootPath)
            Debug.WriteLine($"  TryLoadPdfiumFrom: LoadLibrary returned {h}")
            If h <> IntPtr.Zero Then Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Burns annotations into rasterized images. Excludes sticky notes and popups.
    ''' Optionally renders reason code labels in white text inside burned-in boxes.
    ''' </summary>
    Private Shared Sub APRedactBurnIn(inputPath As String, outputPath As String,
                                       includeReasonCodes As Boolean,
                                       Optional dpi As Integer = 300)

        ' CRITICAL: Pre-load pdfium.dll BEFORE the JIT touches any PdfiumViewer type.
        ' This must happen in a separate method from the one that references
        ' PdfiumViewer.PdfDocument, because the JIT resolves type references when
        ' compiling the method — not when execution reaches the line.
        APRedactEnsurePdfiumLoaded()
        EnsureApPdfSharpFontResolver()

        ' Delegate to NoInlining helper so PdfiumViewer types are resolved
        ' only after LoadLibrary has already succeeded.
        APRedactBurnInCore(inputPath, outputPath, includeReasonCodes, dpi)
    End Sub

    ''' <summary>
    ''' Core burn-in implementation, separated to ensure pdfium.dll is loaded
    ''' before PdfiumViewer types are JIT-resolved.
    ''' </summary>
    <Runtime.CompilerServices.MethodImpl(Runtime.CompilerServices.MethodImplOptions.NoInlining)>
    Private Shared Sub APRedactBurnInCore(inputPath As String, outputPath As String,
                                           includeReasonCodes As Boolean, dpi As Integer)

        Dim tempPath As String = Path.GetTempFileName() & ".pdf"

        Try
            Using inputDoc As PdfSharp.Pdf.PdfDocument =
                PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)

                For Each page As PdfSharp.Pdf.PdfPage In inputDoc.Pages
                    If page.Annotations IsNot Nothing AndAlso page.Annotations.Count > 0 Then
                        Dim toRemove As New List(Of PdfSharp.Pdf.Annotations.PdfAnnotation)()

                        Using gfx As PdfSharp.Drawing.XGraphics =
                            PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)

                            For Each annot As PdfSharp.Pdf.Annotations.PdfAnnotation In page.Annotations
                                Dim subtypeRaw As String =
                                    annot.Elements.GetName(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Subtype)
                                Dim subtype As String =
                                    If(String.IsNullOrEmpty(subtypeRaw), String.Empty,
                                       subtypeRaw.TrimStart("/"c).ToLowerInvariant())

                                ' Keep sticky notes and popups
                                If subtype = "text" OrElse subtype = "popup" Then Continue For

                                Dim rect As PdfSharp.Pdf.PdfRectangle =
                                    annot.Elements.GetRectangle(PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Rect)
                                If rect Is Nothing Then Continue For

                                Dim x As Double = rect.X1
                                Dim yTop As Double = rect.Y2
                                Dim width As Double = rect.X2 - rect.X1
                                Dim height As Double = rect.Y2 - rect.Y1
                                Dim yDraw As Double = page.Height.Point - yTop

                                gfx.DrawRectangle(New PdfSharp.Drawing.XSolidBrush(
                                    PdfSharp.Drawing.XColors.Black), x, yDraw, width, height)

                                ' Render reason code label if requested
                                Dim contents As String = annot.Elements.GetString(
                                    PdfSharp.Pdf.Annotations.PdfAnnotation.Keys.Contents)
                                If includeReasonCodes AndAlso Not String.IsNullOrEmpty(contents) Then
                                    Dim pad As Double = Math.Max(1.0, Math.Min(6.0, height * 0.1))
                                    Dim textRect As New PdfSharp.Drawing.XRect(
                                        x + pad, yDraw + pad,
                                        Math.Max(0, width - 2 * pad),
                                        Math.Max(0, height - 2 * pad))

                                    If textRect.Width > 2 AndAlso textRect.Height > 6 Then
                                        Dim family As String = "Arial"
                                        Dim maxFontSize As Double = Math.Max(6.0, Math.Min(22.0, textRect.Height * 0.9))
                                        Dim fittedSize As Double = APRedactFitFontSize(
                                            gfx, contents, family, textRect, 6.0, maxFontSize)

                                        Dim style As PdfSharp.Drawing.XFontStyleEx
                                        Try
                                            style = CType([Enum].Parse(GetType(PdfSharp.Drawing.XFontStyleEx),
                                                "Regular", True), PdfSharp.Drawing.XFontStyleEx)
                                        Catch
                                            style = CType([Enum].ToObject(GetType(PdfSharp.Drawing.XFontStyleEx), 0),
                                                PdfSharp.Drawing.XFontStyleEx)
                                        End Try

                                        Dim font As New PdfSharp.Drawing.XFont(family, fittedSize, style)
                                        Dim state = gfx.Save()
                                        Try
                                            gfx.IntersectClip(textRect)
                                            Dim tf As New PdfSharp.Drawing.Layout.XTextFormatter(gfx)
                                            tf.Alignment = PdfSharp.Drawing.Layout.XParagraphAlignment.Left
                                            tf.DrawString(contents, font, PdfSharp.Drawing.XBrushes.White,
                                                          textRect, PdfSharp.Drawing.XStringFormats.TopLeft)
                                        Finally
                                            gfx.Restore(state)
                                        End Try
                                    End If
                                End If

                                toRemove.Add(annot)
                            Next
                        End Using

                        For Each a In toRemove
                            page.Annotations.Remove(a)
                        Next
                    End If
                Next

                inputDoc.Save(tempPath)
            End Using

            ' Rasterize to final output
            Using pdf As PdfiumViewer.PdfDocument = PdfiumViewer.PdfDocument.Load(tempPath)
                Dim outDoc As New PdfSharp.Pdf.PdfDocument()

                ' Preserve metadata
                APRedactCopyMetadataFromPath(tempPath, outDoc)

                For pageIndex As Integer = 0 To pdf.PageCount - 1
                    Dim sizePt As System.Drawing.SizeF = pdf.PageSizes(pageIndex)
                    Dim widthPx As Integer = CInt(Math.Round(sizePt.Width / 72.0 * dpi))
                    Dim heightPx As Integer = CInt(Math.Round(sizePt.Height / 72.0 * dpi))

                    Dim renderFlags As PdfiumViewer.PdfRenderFlags =
                        PdfiumViewer.PdfRenderFlags.Annotations Or
                        PdfiumViewer.PdfRenderFlags.LcdText Or
                        PdfiumViewer.PdfRenderFlags.ForPrinting

                    Using rendered As System.Drawing.Image =
                        pdf.Render(pageIndex, widthPx, heightPx, dpi, dpi, renderFlags)
                        Dim outPage As PdfSharp.Pdf.PdfPage = outDoc.AddPage()
                        outPage.Width = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Width)
                        outPage.Height = PdfSharp.Drawing.XUnit.FromPoint(sizePt.Height)
                        Using ms As New MemoryStream()
                            rendered.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                            ms.Position = 0
                            Using xgfx As PdfSharp.Drawing.XGraphics =
                                PdfSharp.Drawing.XGraphics.FromPdfPage(outPage)
                                Using ximg As PdfSharp.Drawing.XImage = PdfSharp.Drawing.XImage.FromStream(ms)
                                    xgfx.DrawImage(ximg, 0, 0, outPage.Width.Point, outPage.Height.Point)
                                End Using
                            End Using
                        End Using
                    End Using
                Next

                outDoc.Save(outputPath)
                outDoc.Close()
            End Using
        Finally
            Try : If File.Exists(tempPath) Then File.Delete(tempPath)
            Catch : End Try
        End Try
    End Sub


    Private Shared Sub APRedactCopyMetadataFromPath(srcPath As String, dest As PdfSharp.Pdf.PdfDocument)
        If String.IsNullOrWhiteSpace(srcPath) OrElse dest Is Nothing Then Return
        Try
            Using src As PdfSharp.Pdf.PdfDocument =
                PdfSharp.Pdf.IO.PdfReader.Open(srcPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.InformationOnly)

                Dim s = src.Info : Dim d = dest.Info
                d.Title = s.Title : d.Author = s.Author : d.Subject = s.Subject
                d.Keywords = s.Keywords : d.Creator = s.Creator
                If s.CreationDate <> Date.MinValue Then
                    d.CreationDate = If(s.CreationDate.Kind = DateTimeKind.Utc,
                                        s.CreationDate.ToLocalTime(), s.CreationDate)
                End If
                If s.ModificationDate <> Date.MinValue Then
                    d.ModificationDate = If(s.ModificationDate.Kind = DateTimeKind.Utc,
                                            s.ModificationDate.ToLocalTime(), s.ModificationDate)
                End If
            End Using
        Catch
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF REDACTOR — FONT FITTING HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared Function APRedactFitFontSize(gfx As PdfSharp.Drawing.XGraphics,
                                                 text As String, family As String,
                                                 rect As PdfSharp.Drawing.XRect,
                                                 minSize As Double, maxSize As Double) As Double
        If String.IsNullOrEmpty(text) Then Return minSize
        Dim low As Double = minSize
        Dim high As Double = Math.Max(minSize, maxSize)
        Dim best As Double = low

        Dim style As PdfSharp.Drawing.XFontStyleEx
        Try
            style = CType([Enum].Parse(GetType(PdfSharp.Drawing.XFontStyleEx), "Regular", True),
                PdfSharp.Drawing.XFontStyleEx)
        Catch
            style = CType([Enum].ToObject(GetType(PdfSharp.Drawing.XFontStyleEx), 0),
                PdfSharp.Drawing.XFontStyleEx)
        End Try

        While (high - low) > 0.5
            Dim mid = (low + high) / 2.0
            Dim f As New PdfSharp.Drawing.XFont(family, mid, style)
            If APRedactTextFits(gfx, text, f, rect) Then
                best = mid : low = mid
            Else
                high = mid
            End If
        End While

        Return Math.Max(minSize, Math.Min(best, maxSize))
    End Function

    Private Shared Function APRedactTextFits(gfx As PdfSharp.Drawing.XGraphics,
                                              text As String, font As PdfSharp.Drawing.XFont,
                                              rect As PdfSharp.Drawing.XRect) As Boolean
        If String.IsNullOrEmpty(text) Then Return True
        Dim words As String() = System.Text.RegularExpressions.Regex.Split(text, "\s+")
        Dim spaceW As Double = gfx.MeasureString(" ", font).Width
        Dim lineW As Double = 0 : Dim lines As Integer = 1
        Dim lineH As Double = font.Size * 1.2

        For Each w In words
            Dim wW As Double = gfx.MeasureString(w, font).Width
            If lineW > 0 AndAlso (lineW + spaceW + wW) > rect.Width Then
                lines += 1 : lineW = wW
            Else
                lineW = If(lineW = 0, wW, lineW + spaceW + wW)
            End If
            If lines * lineH > rect.Height Then Return False
        Next
        Return True
    End Function

End Class