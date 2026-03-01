' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.PDFCommentProcessor.vb
' Purpose:
'   Adds review comments to PDF attachments by creating native PDF annotations
'   (highlight + popup), driven by LLM-generated quote/comment pairs.
'
' Architecture:
'  - Extraction:
'      * Reads PDF text and per-character coordinates using PdfPig.
'      * Builds a searchable text stream plus positional map for precise anchoring.
'  - LLM protocol:
'      * Uses the same bubble protocol as AutoPilot comment processing
'        (`text@@comment§§§...`) and reuses `APParseBubblesResponse`.
'      * Supports single-pass and batched mode for large PDFs with context windows.
'  - Matching:
'      * Resolves quoted text spans with multi-tier matching:
'        exact, case-insensitive, normalized whitespace, letters-only normalization,
'        and fuzzy prefix fallback.
'      * Converts matched spans into per-line highlight rectangles by page.
'  - Injection:
'      * Writes annotations with PdfSharp low-level dictionaries:
'          - `/Highlight` + `/Popup` for matched quotes
'          - `/Text` sticky notes for unmatched comments
'      * Adds stable metadata (`/T`, `/NM`, creation/modification timestamps, color).
'  - Output behavior:
'      * Works on a copied output PDF, preserves input unchanged.
'      * Saves modified annotations directly into the destination file.
'
' Dependencies:
'  - PdfPig (`UglyToad.PdfPig`) for text/position extraction.
'  - PdfSharp for annotation object creation and PDF write-back.
'  - AutoPilot comment pipeline helpers from:
'      * `ThisAddIn.AutoPilot.CommentProcessor.vb`
'        (`APGetCommentsFromLLM`, `APParseBubblesResponse`, prompt constants).
'
' Notes:
'  - Unmatched comments are placed as stacked sticky notes on page 1 (top-right),
'    with overlap avoidance and fallback column wrapping.
'  - Optional custom author names are supported; non-default author mode prefixes
'    comment text with the Red Ink marker.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports PDFP = UglyToad.PdfPig

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — DATA CLASSES
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents a single character's position within the PDF for text-matching.
    ''' </summary>
    Private Class PdfCharPosition
        Public Property PageIndex As Integer       ' 0-based page index
        Public Property X As Double                ' Left edge in PDF points
        Public Property Y As Double                ' Bottom edge in PDF points (PDF coordinate system)
        Public Property Width As Double            ' Character width in points
        Public Property Height As Double           ' Character height in points
        Public Property PageWidth As Double        ' Page width for fallback positioning
        Public Property PageHeight As Double       ' Page height for coordinate transforms
    End Class

    ''' <summary>
    ''' Represents one highlight rectangle on a specific page.
    ''' </summary>
    Private Class PdfHighlightRect
        Public Property PageIndex As Integer       ' 0-based
        Public Property X As Double                ' Left
        Public Property Y As Double                ' Bottom (PDF coords)
        Public Property Width As Double
        Public Property Height As Double
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Maximum characters intended for a single LLM batch when processing PDFs.
    ''' </summary>
    Private Const AP_PdfCommentMaxCharsPerBatch As Integer = 15000

    ''' <summary>
    ''' Vertical spacing between stacked unmatched sticky notes (in points).
    ''' </summary>
    Private Const AP_PdfStickyNoteSpacing As Double = 30.0

    ''' <summary>
    ''' Size of the sticky note icon rectangle (in points).
    ''' </summary>
    Private Const AP_PdfStickyNoteSize As Double = 24.0

    ''' <summary>
    ''' Right margin offset for sticky note placement (in points from right edge).
    ''' </summary>
    Private Const AP_PdfStickyNoteRightMargin As Double = 30.0

    ''' <summary>
    ''' Top margin offset for the first sticky note (in points from top edge).
    ''' </summary>
    Private Const AP_PdfStickyNoteTopMargin As Double = 30.0

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF COMMENT PROCESSOR — MAIN ENTRY POINT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes a PDF file by having the LLM add comment annotations to relevant sections
    ''' based on the given instruction.
    ''' </summary>
    ''' <param name="author">Optional author name for the comments. If Nothing or empty, defaults to AN6 ("Inky").
    ''' When a custom author is used, each comment is prefixed with "RI: " (AN5).</param>
    Private Async Function CommentPdfForAutoPilot(inputPath As String, outputPath As String,
                                                   instruction As String, ct As CancellationToken,
                                                   Optional author As String = Nothing) As Task(Of Boolean)
        ' Resolve the effective author and determine whether to prefix comments
        Dim effectiveAuthor As String = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
        Dim usePrefix As Boolean = Not effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase)

        Try
            ApDashboardLog("PdfCommentProcessor: starting PDF comment processing", "step")

            ' ─── Step 1: Extract text and character positions via PdfPig ───
            Dim fullText As String = Nothing
            Dim charPositions As List(Of PdfCharPosition) = Nothing
            APPdfExtractTextAndPositions(inputPath, fullText, charPositions)

            If String.IsNullOrWhiteSpace(fullText) OrElse charPositions Is Nothing OrElse charPositions.Count = 0 Then
                ApDashboardLog("PdfCommentProcessor: PDF contains no extractable text", "warn")
                Return False
            End If

            ApDashboardLog($"PdfCommentProcessor: extracted {fullText.Length} characters ({charPositions.Count} positioned chars)", "step")

            ' ─── Step 2: Get comments from LLM (reuses DOCX comment processor's LLM logic) ───
            Dim comments As List(Of APCommentEntry) = Nothing

            If fullText.Length <= AP_PdfCommentMaxCharsPerBatch Then
                ApDashboardLog("PdfCommentProcessor: using single-pass mode", "step")
                comments = Await APGetCommentsFromLLM(fullText, instruction, ct)
            Else
                ' For PDFs we use simple chunking of the full text since we don't have paragraph structure
                ApDashboardLog($"PdfCommentProcessor: document text ({fullText.Length} chars) exceeds single-pass limit, using batched mode", "step")
                comments = Await APPdfGetCommentsFromLLMBatched(fullText, instruction, ct)
            End If

            If comments Is Nothing OrElse comments.Count = 0 Then
                ApDashboardLog("PdfCommentProcessor: LLM returned no comments", "warn")
                Return False
            End If

            ApDashboardLog($"PdfCommentProcessor: {comments.Count} comments to insert", "step")

            ' ─── Step 3: Insert annotations into the PDF via PdfSharp ───
            EnsureApPdfSharpFontResolver()

            File.Copy(inputPath, outputPath, overwrite:=True)

            Using pdfDoc As PdfDocument = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify)
                Dim commentPrefix As String = If(usePrefix, AN5 & ": ", "")
                Dim insertedCount As Integer = 0
                Dim unmatchedEntries As New List(Of APCommentEntry)()
                Dim commentId As Integer = 0

                For Each entry In comments
                    ct.ThrowIfCancellationRequested()
                    commentId += 1

                    ' Find the quoted text in the character position map
                    Dim highlightRects = APPdfFindTextRects(fullText, charPositions, entry.QuotedText)

                    If highlightRects Is Nothing OrElse highlightRects.Count = 0 Then
                        unmatchedEntries.Add(entry)
                        Debug.WriteLine("PdfCommentProcessor: Could not find quoted text: " &
                                        entry.QuotedText.Substring(0, Math.Min(60, entry.QuotedText.Length)))
                        Continue For
                    End If

                    ' Insert highlight + popup annotations for matched text
                    Dim commentBody = commentPrefix & entry.CommentText
                    APPdfInsertHighlightAnnotation(pdfDoc, highlightRects, commentBody,
                                                    effectiveAuthor, commentId)
                    insertedCount += 1
                Next

                ' ─── Handle unmatched comments: sticky notes at top-right of first page ───
                If unmatchedEntries.Count > 0 Then
                    Dim firstPage = pdfDoc.Pages(0)
                    Dim pageW = firstPage.Width.Point
                    Dim pageH = firstPage.Height.Point
                    Dim noteX = pageW - AP_PdfStickyNoteRightMargin - AP_PdfStickyNoteSize

                    For i = 0 To unmatchedEntries.Count - 1
                        ct.ThrowIfCancellationRequested()
                        commentId += 1

                        Dim noteY = pageH - AP_PdfStickyNoteTopMargin - (i * AP_PdfStickyNoteSpacing)

                        ' Clamp to page bounds — if we run out of vertical space, wrap to next column
                        If noteY < AP_PdfStickyNoteSize Then
                            noteX -= (AP_PdfStickyNoteSize + 10)
                            noteY = pageH - AP_PdfStickyNoteTopMargin
                        End If

                        Dim quotedExcerpt = unmatchedEntries(i).QuotedText

                        Dim noteBody = commentPrefix & unmatchedEntries(i).CommentText &
                                       vbCrLf & vbCrLf & "[Referenced text not found: """ & quotedExcerpt & """]"

                        APPdfInsertStickyNoteAnnotation(pdfDoc, 0, noteX, noteY,
                                                        noteBody, effectiveAuthor, commentId)
                        insertedCount += 1
                    Next

                    ApDashboardLog($"PdfCommentProcessor: {unmatchedEntries.Count} unmatched comments placed as sticky notes", "step")
                End If

                If insertedCount = 0 Then
                    ApDashboardLog("PdfCommentProcessor: no comments could be matched or placed", "warn")
                    Return False
                End If

                ApDashboardLog($"PdfCommentProcessor: {insertedCount} annotations inserted into PDF", "step")

                pdfDoc.Save(outputPath)
            End Using

            ApDashboardLog("PdfCommentProcessor: PDF comment processing complete", "success")
            Return True

        Catch ex As System.Exception
            Debug.WriteLine("CommentPdfForAutoPilot error: " & ex.Message)
            ApDashboardLog("PdfCommentProcessor: error - " & ex.Message, "error")
            Return False
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT EXTRACTION WITH POSITIONS (via PdfPig)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts the full concatenated text from a PDF and builds a parallel list
    ''' of character positions (page, x, y, width, height) for text-matching.
    ''' Uses PdfPig's word-level extraction to correctly insert spaces between words,
    ''' since PDF content streams do not contain explicit space characters — spaces
    ''' are implied by glyph positioning gaps.
    ''' </summary>
    Private Shared Sub APPdfExtractTextAndPositions(pdfPath As String,
                                                    ByRef fullText As String,
                                                    ByRef positions As List(Of PdfCharPosition))
        Dim sb As New StringBuilder()
        positions = New List(Of PdfCharPosition)()

        Using document As PDFP.PdfDocument = PDFP.PdfDocument.Open(pdfPath)
            For pageIdx = 0 To document.NumberOfPages - 1
                Dim page = document.GetPage(pageIdx + 1) ' PdfPig is 1-based
                Dim pageWidth = page.Width
                Dim pageHeight = page.Height

                ' Use PdfPig's word extraction to get proper word boundaries.
                ' Each Word contains its constituent Letters with positions,
                ' and we insert a space between consecutive words.
                Dim words = page.GetWords().ToList()
                Dim isFirstWordOnPage As Boolean = True

                For Each word As PDFP.Content.Word In words
                    ' Insert space before this word (unless it's the first on the page
                    ' or the first word after a page separator)
                    If Not isFirstWordOnPage Then
                        ' Detect line breaks: if this word's Y position differs significantly
                        ' from what came before, insert a newline instead of a space
                        Dim lastPos = If(positions.Count > 0, positions(positions.Count - 1), Nothing)
                        Dim isNewLine As Boolean = False

                        If lastPos IsNot Nothing AndAlso lastPos.PageIndex = pageIdx Then
                            Dim yDiff = Math.Abs(word.BoundingBox.Bottom - lastPos.Y)
                            Dim lineHeight = If(lastPos.Height > 0, lastPos.Height, 12.0)
                            isNewLine = (yDiff > lineHeight * 0.5)
                        End If

                        Dim separator = If(isNewLine, vbLf, " ")
                        sb.Append(separator)
                        ' Add dummy position entry for the separator character
                        positions.Add(New PdfCharPosition() With {
                            .PageIndex = pageIdx,
                            .X = 0, .Y = 0, .Width = 0, .Height = 0,
                            .PageWidth = pageWidth,
                            .PageHeight = pageHeight
                        })
                    End If
                    isFirstWordOnPage = False

                    ' Append each letter of the word with its position
                    For Each letter As PDFP.Content.Letter In word.Letters
                        Dim ch = letter.Value
                        If String.IsNullOrEmpty(ch) Then Continue For

                        For Each c In ch
                            sb.Append(c)
                            positions.Add(New PdfCharPosition() With {
                                .PageIndex = pageIdx,
                                .X = letter.GlyphRectangle.Left,
                                .Y = letter.GlyphRectangle.Bottom,
                                .Width = letter.GlyphRectangle.Width,
                                .Height = letter.GlyphRectangle.Height,
                                .PageWidth = pageWidth,
                                .PageHeight = pageHeight
                            })
                        Next
                    Next
                Next

                ' Add a space between pages to prevent cross-page word merging
                If pageIdx < document.NumberOfPages - 1 Then
                    sb.Append(" ")
                    positions.Add(New PdfCharPosition() With {
                        .PageIndex = pageIdx,
                        .X = 0, .Y = 0, .Width = 0, .Height = 0,
                        .PageWidth = pageWidth,
                        .PageHeight = pageHeight
                    })
                End If
            Next
        End Using

        fullText = sb.ToString()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  LLM CALL — BATCHED (for large PDFs without paragraph structure)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes the PDF text in batches, sending each batch to the LLM with
    ''' surrounding context so the LLM understands the flow.
    ''' </summary>
    Private Async Function APPdfGetCommentsFromLLMBatched(fullText As String,
                                                          instruction As String,
                                                          ct As CancellationToken) As Task(Of List(Of APCommentEntry))
        Dim allComments As New List(Of APCommentEntry)()

        ' Build the system prompt once
        Dim bubblesPrompt As String = InterpolateAtRuntime(SP_Add_Bubbles)
        bubblesPrompt = bubblesPrompt.Replace("{FormatInstruction}",
            "Provide each comment as plain text without any markdown formatting.")

        Dim systemPrompt As String =
            "You are a professional document reviewer. " &
            "Apply the following instruction to the document text provided." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            bubblesPrompt & vbCrLf & vbCrLf &
            "IMPORTANT: The text between [TEXTTOPROCESS] and [/TEXTTOPROCESS] is the section you must review and comment on. " &
            "Text between [CONTEXT_BEFORE] and [/CONTEXT_BEFORE] or [CONTEXT_AFTER] and [/CONTEXT_AFTER] is provided " &
            "for context only — do NOT comment on context sections."

        Dim contextSize As Integer = 500 ' Characters of context before/after each batch
        Dim totalLen = fullText.Length
        Dim estimatedBatches As Integer = CInt(Math.Ceiling(totalLen / CDbl(AP_PdfCommentMaxCharsPerBatch)))

        ApDashboardLog($"PdfCommentProcessor: {totalLen} chars to process (~{estimatedBatches} batches)", "step")

        Dim batchStart As Integer = 0
        Dim currentBatch As Integer = 0

        While batchStart < totalLen
            ct.ThrowIfCancellationRequested()
            currentBatch += 1

            ' Determine batch end
            Dim batchEnd = Math.Min(batchStart + AP_PdfCommentMaxCharsPerBatch, totalLen)

            ' Try to break at a word boundary
            If batchEnd < totalLen Then
                Dim lastSpace = fullText.LastIndexOf(" "c, batchEnd, Math.Min(200, batchEnd - batchStart))
                If lastSpace > batchStart Then batchEnd = lastSpace
            End If

            Dim charCount = batchEnd - batchStart
            ApDashboardLog($"PdfCommentProcessor: batch {currentBatch}/{estimatedBatches} (chars {batchStart + 1}-{batchEnd}, {charCount} chars)", "step")

            ' Build context before
            Dim ctxBeforeStart = Math.Max(0, batchStart - contextSize)
            Dim contextBefore = If(ctxBeforeStart < batchStart, fullText.Substring(ctxBeforeStart, batchStart - ctxBeforeStart), "")

            ' Main text
            Dim mainText = fullText.Substring(batchStart, charCount)

            ' Build context after
            Dim ctxAfterEnd = Math.Min(totalLen, batchEnd + contextSize)
            Dim contextAfter = If(batchEnd < ctxAfterEnd, fullText.Substring(batchEnd, ctxAfterEnd - batchEnd), "")

            ' Assemble user prompt with context markers
            Dim userPrompt As New StringBuilder()
            If contextBefore.Length > 0 Then
                userPrompt.AppendLine("[CONTEXT_BEFORE]")
                userPrompt.AppendLine(contextBefore)
                userPrompt.AppendLine("[/CONTEXT_BEFORE]")
                userPrompt.AppendLine()
            End If

            userPrompt.AppendLine("[TEXTTOPROCESS]")
            userPrompt.AppendLine(mainText)
            userPrompt.AppendLine("[/TEXTTOPROCESS]")

            If contextAfter.Length > 0 Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("[CONTEXT_AFTER]")
                userPrompt.AppendLine(contextAfter)
                userPrompt.AppendLine("[/CONTEXT_AFTER]")
            End If

            ' Call LLM for this batch
            Dim llmResponse = Await LLM(systemPrompt, userPrompt.ToString(),
                                         UseSecondAPI:=False,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If Not String.IsNullOrWhiteSpace(llmResponse) Then
                Dim batchComments = APParseBubblesResponse(llmResponse)
                If batchComments IsNot Nothing AndAlso batchComments.Count > 0 Then
                    allComments.AddRange(batchComments)
                    ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned {batchComments.Count} comments", "step")
                Else
                    ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned no comments", "warn")
                End If
            Else
                ApDashboardLog($"PdfCommentProcessor: batch {currentBatch} returned empty response", "warn")
            End If

            batchStart = batchEnd
        End While

        ApDashboardLog($"PdfCommentProcessor: all {currentBatch} batches completed, {allComments.Count} total comments collected", "success")
        Return allComments
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT SPAN → HIGHLIGHT RECTANGLES
    ' ═══════════════════════════════════════════════════════════════════════════


    ''' <summary>
    ''' Finds the character positions matching the search text and groups them
    ''' into per-line highlight rectangles, one per visual line.
    ''' </summary>
    Private Shared Function APPdfFindTextRects(fullText As String,
                                               charPositions As List(Of PdfCharPosition),
                                               searchText As String) As List(Of PdfHighlightRect)
        If String.IsNullOrWhiteSpace(searchText) OrElse charPositions.Count = 0 Then Return Nothing

        ' Strip any residual surrounding quotes the LLM may have left
        Dim cleanedSearch = searchText
        If cleanedSearch.Length >= 2 Then
            Dim first = cleanedSearch(0)
            Dim last = cleanedSearch(cleanedSearch.Length - 1)
            If (first = """"c AndAlso last = """"c) OrElse
               (first = ChrW(&H201C) AndAlso (last = ChrW(&H201D) OrElse last = ChrW(&H201C))) OrElse
               (first = ChrW(&H201E) AndAlso (last = ChrW(&H201D) OrElse last = ChrW(&H201C))) OrElse
               (first = ChrW(&HAB) AndAlso last = ChrW(&HBB)) Then
                cleanedSearch = cleanedSearch.Substring(1, cleanedSearch.Length - 2).Trim()
            End If
        End If

        ' ─── Tier 1: Exact match ───
        Dim matchIdx = fullText.IndexOf(cleanedSearch, StringComparison.Ordinal)

        ' ─── Tier 2: Case-insensitive ───
        If matchIdx < 0 Then
            matchIdx = fullText.IndexOf(cleanedSearch, StringComparison.OrdinalIgnoreCase)
        End If

        ' ─── Tier 3: Normalized whitespace ───
        If matchIdx < 0 Then
            Dim normalizedFull = Regex.Replace(fullText, "\s+", " ")
            Dim normalizedSearch = Regex.Replace(cleanedSearch, "\s+", " ")
            Dim normIdx = normalizedFull.IndexOf(normalizedSearch, StringComparison.OrdinalIgnoreCase)
            If normIdx >= 0 Then
                matchIdx = APPdfMapNormalizedIndexToOriginal(fullText, normIdx)
            End If
        End If

        ' ─── Tier 4: Letters-only normalized match ───
        ' Strips all whitespace, hyphens, and digits from both texts, keeping only
        ' letters and punctuation (except hyphens). This single pass handles:
        '   - Line-break hyphenation: "interna-\ntional" / "Tat-bestand" → "international" / "Tatbestand"
        '   - Hyphen-less word splits: "Vermö\ngensdelikte" → "Vermögensdelikte"
        '   - Footnote reference numbers: "agreement1 shall" → "agreement shall"
        '   - Any combination of the above
        Dim normalizedMatchLen As Integer = -1
        If matchIdx < 0 Then
            Dim normFull As String = Nothing
            Dim normFullMap As List(Of Integer) = Nothing
            APPdfBuildLettersOnlyText(fullText, normFull, normFullMap)

            Dim normSearch As String = Nothing
            Dim normSearchMap As List(Of Integer) = Nothing
            APPdfBuildLettersOnlyText(cleanedSearch, normSearch, normSearchMap)

            If normSearch.Length > 0 Then
                Dim normIdx = normFull.IndexOf(normSearch, StringComparison.OrdinalIgnoreCase)
                If normIdx >= 0 Then
                    matchIdx = normFullMap(normIdx)
                    Dim normEndIdx = Math.Min(normIdx + normSearch.Length - 1, normFullMap.Count - 1)
                    normalizedMatchLen = normFullMap(normEndIdx) - matchIdx + 1
                End If
            End If
        End If

        ' ─── Tier 5: Try with original (unstripped quotes) search text ───
        If matchIdx < 0 AndAlso cleanedSearch <> searchText Then
            matchIdx = fullText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase)
        End If

        ' ─── Tier 6: Fuzzy prefix match against raw text ───
        If matchIdx < 0 AndAlso cleanedSearch.Length >= 50 Then
            Dim prefixLen = cleanedSearch.Length - 1
            While prefixLen >= 40 AndAlso matchIdx < 0
                Dim prefix = cleanedSearch.Substring(0, prefixLen)
                matchIdx = fullText.IndexOf(prefix, StringComparison.OrdinalIgnoreCase)
                If matchIdx >= 0 Then
                    cleanedSearch = prefix
                End If
                prefixLen -= 5
            End While
        End If

        ' ─── Tier 7: Fuzzy prefix match against letters-only normalized text ───
        If matchIdx < 0 AndAlso cleanedSearch.Length >= 50 Then
            Dim normFull As String = Nothing
            Dim normFullMap As List(Of Integer) = Nothing
            APPdfBuildLettersOnlyText(fullText, normFull, normFullMap)

            Dim normSearch As String = Nothing
            Dim normSearchMap As List(Of Integer) = Nothing
            APPdfBuildLettersOnlyText(cleanedSearch, normSearch, normSearchMap)

            If normSearch.Length >= 40 Then
                Dim prefixLen = normSearch.Length - 1
                While prefixLen >= 40 AndAlso matchIdx < 0
                    Dim prefix = normSearch.Substring(0, prefixLen)
                    Dim normIdx = normFull.IndexOf(prefix, StringComparison.OrdinalIgnoreCase)
                    If normIdx >= 0 Then
                        matchIdx = normFullMap(normIdx)
                        Dim normEndIdx = Math.Min(normIdx + prefix.Length - 1, normFullMap.Count - 1)
                        normalizedMatchLen = normFullMap(normEndIdx) - matchIdx + 1
                    End If
                    prefixLen -= 5
                End While
            End If
        End If

        If matchIdx < 0 Then Return Nothing

        ' Determine the match length in original text coordinates
        Dim matchLen As Integer
        If normalizedMatchLen > 0 Then
            matchLen = normalizedMatchLen
        Else
            matchLen = If(fullText.IndexOf(cleanedSearch, matchIdx, StringComparison.OrdinalIgnoreCase) = matchIdx,
                          cleanedSearch.Length, searchText.Length)
        End If
        Dim matchEnd = matchIdx + matchLen - 1

        ' Clamp to available positions
        If matchIdx >= charPositions.Count Then Return Nothing
        If matchEnd >= charPositions.Count Then matchEnd = charPositions.Count - 1

        ' Collect the character positions for the match range
        Dim matchedPositions As New List(Of PdfCharPosition)()
        For i = matchIdx To matchEnd
            Dim pos = charPositions(i)
            ' Skip dummy separator entries
            If pos.Width > 0 OrElse pos.Height > 0 Then
                matchedPositions.Add(pos)
            End If
        Next

        If matchedPositions.Count = 0 Then Return Nothing

        ' Group into visual lines: characters on the same page with similar Y coordinate
        Dim rects As New List(Of PdfHighlightRect)()
        Dim lineChars As New List(Of PdfCharPosition)()
        lineChars.Add(matchedPositions(0))

        For i = 1 To matchedPositions.Count - 1
            Dim prev = matchedPositions(i - 1)
            Dim curr = matchedPositions(i)

            ' Same line if same page and Y within tolerance (half the character height)
            Dim yTolerance = Math.Max(prev.Height * 0.5, 2.0)
            Dim sameLine = (curr.PageIndex = prev.PageIndex) AndAlso
                           (Math.Abs(curr.Y - prev.Y) <= yTolerance)

            If sameLine Then
                lineChars.Add(curr)
            Else
                ' Flush current line
                rects.Add(APPdfBuildLineRect(lineChars))
                lineChars.Clear()
                lineChars.Add(curr)
            End If
        Next

        ' Flush last line
        If lineChars.Count > 0 Then
            rects.Add(APPdfBuildLineRect(lineChars))
        End If

        Return rects
    End Function


    ''' <summary>
    ''' Builds a single highlight rectangle from a list of characters on the same visual line.
    ''' </summary>
    Private Shared Function APPdfBuildLineRect(lineChars As List(Of PdfCharPosition)) As PdfHighlightRect
        Dim minX = lineChars.Min(Function(c) c.X)
        Dim minY = lineChars.Min(Function(c) c.Y)
        Dim maxX = lineChars.Max(Function(c) c.X + c.Width)
        Dim maxY = lineChars.Max(Function(c) c.Y + c.Height)

        Return New PdfHighlightRect() With {
            .PageIndex = lineChars(0).PageIndex,
            .X = minX,
            .Y = minY,
            .Width = maxX - minX,
            .Height = maxY - minY
        }
    End Function



    ''' <summary>
    ''' Maps a position in normalized (collapsed-whitespace) text back to the original string.
    ''' </summary>
    Private Shared Function APPdfMapNormalizedIndexToOriginal(original As String, normalizedIdx As Integer) As Integer
        Dim normPos As Integer = 0
        Dim origPos As Integer = 0
        Dim inWhitespace As Boolean = False

        While origPos < original.Length AndAlso normPos < normalizedIdx
            If Char.IsWhiteSpace(original(origPos)) Then
                If Not inWhitespace Then
                    normPos += 1
                    inWhitespace = True
                End If
            Else
                normPos += 1
                inWhitespace = False
            End If
            origPos += 1
        End While

        Return origPos
    End Function

    ''' <summary>
    ''' Builds a "letters-only" version of the input text by stripping all whitespace,
    ''' hyphens (ASCII and Unicode), and digit characters. Only letters and non-hyphen
    ''' punctuation are preserved. An index map is built so each character in the
    ''' normalized output maps back to its position in the original text.
    ''' </summary>
    ''' <remarks>
    ''' This single normalization pass replaces the previous separate dehyphenation and
    ''' digit-stripping helpers. By reducing both the PDF text and the LLM's quoted text
    ''' to the same letters-and-punctuation skeleton, all of these mismatches are resolved
    ''' in one comparison:
    ''' <list type="bullet">
    '''   <item>Line-break hyphenation: "interna-\ntional" → "international"</item>
    '''   <item>LLM-preserved hyphens: "Tat-bestand" → "Tatbestand"</item>
    '''   <item>Hyphen-less word splits: "Vermö\ngensdelikte" → "Vermögensdelikte"</item>
    '''   <item>Footnote reference numbers: "agreement1 shall" → "agreement shall"</item>
    '''   <item>Any combination of the above</item>
    ''' </list>
    ''' Non-hyphen punctuation (commas, periods, parentheses, etc.) is kept because it
    ''' provides essential structure that prevents false-positive matches between unrelated
    ''' text passages. Stripping punctuation would make "A, B" match "AB" anywhere.
    ''' </remarks>
    Private Shared Sub APPdfBuildLettersOnlyText(fullText As String,
                                                  ByRef normalized As String,
                                                  ByRef indexMap As List(Of Integer))
        Dim sb As New StringBuilder(fullText.Length)
        indexMap = New List(Of Integer)(fullText.Length)

        For i = 0 To fullText.Length - 1
            Dim ch = fullText(i)
            Dim code = AscW(ch)

            ' Skip whitespace (spaces, newlines, tabs)
            If Char.IsWhiteSpace(ch) Then Continue For

            ' Skip digit characters (footnote reference numbers)
            If Char.IsDigit(ch) Then Continue For

            ' Skip all hyphen variants
            If code = 45 OrElse             ' ASCII hyphen-minus
               code = &H2010 OrElse         ' Unicode hyphen
               code = &H2011 OrElse         ' Non-breaking hyphen
               code = &H2012 OrElse         ' Figure dash
               code = &H2013 OrElse         ' En dash
               code = &H2014 OrElse         ' Em dash
               code = &HAD Then             ' Soft hyphen
                Continue For
            End If

            ' Keep everything else (letters and non-hyphen punctuation)
            sb.Append(ch)
            indexMap.Add(i)
        Next

        normalized = sb.ToString()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF ANNOTATION INJECTION (PdfSharp low-level dictionary manipulation)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Inserts a /Highlight annotation with an associated /Popup annotation for a matched
    ''' text span. The highlight covers all line rectangles; the popup is anchored near the
    ''' first highlight rectangle.
    ''' </summary>
    Private Shared Sub APPdfInsertHighlightAnnotation(pdfDoc As PdfDocument,
                                                       highlightRects As List(Of PdfHighlightRect),
                                                       commentText As String,
                                                       author As String,
                                                       commentId As Integer)
        ' Group rectangles by page
        Dim rectsByPage = highlightRects.GroupBy(Function(r) r.PageIndex)

        For Each pageGroup In rectsByPage
            Dim pageIdx = pageGroup.Key
            If pageIdx < 0 OrElse pageIdx >= pdfDoc.Pages.Count Then Continue For

            Dim page = pdfDoc.Pages(pageIdx)
            Dim pageRects = pageGroup.ToList()

            ' Compute the overall bounding rectangle for the /Rect entry
            Dim overallMinX = pageRects.Min(Function(r) r.X)
            Dim overallMinY = pageRects.Min(Function(r) r.Y)
            Dim overallMaxX = pageRects.Max(Function(r) r.X + r.Width)
            Dim overallMaxY = pageRects.Max(Function(r) r.Y + r.Height)

            ' Build QuadPoints array — 8 values per line rectangle
            ' PDF spec: QuadPoints are x1,y1,...,x4,y4 for each quad (top-left, top-right, bottom-left, bottom-right)
            Dim quadPointsArray As New PdfArray(pdfDoc)
            For Each rect In pageRects
                Dim x1 = rect.X                ' left
                Dim x2 = rect.X + rect.Width   ' right
                Dim y1 = rect.Y                ' bottom
                Dim y2 = rect.Y + rect.Height  ' top
                ' PDF QuadPoints order: top-left, top-right, bottom-left, bottom-right
                quadPointsArray.Elements.Add(New PdfReal(x1))   ' top-left X
                quadPointsArray.Elements.Add(New PdfReal(y2))   ' top-left Y
                quadPointsArray.Elements.Add(New PdfReal(x2))   ' top-right X
                quadPointsArray.Elements.Add(New PdfReal(y2))   ' top-right Y
                quadPointsArray.Elements.Add(New PdfReal(x1))   ' bottom-left X
                quadPointsArray.Elements.Add(New PdfReal(y1))   ' bottom-left Y
                quadPointsArray.Elements.Add(New PdfReal(x2))   ' bottom-right X
                quadPointsArray.Elements.Add(New PdfReal(y1))   ' bottom-right Y
            Next

            ' Create the /Highlight annotation
            Dim highlight As New PdfDictionary(pdfDoc)
            highlight.Elements.Add("/Type", New PdfName("/Annot"))
            highlight.Elements.Add("/Subtype", New PdfName("/Highlight"))
            highlight.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, overallMinX, overallMinY, overallMaxX, overallMaxY))
            highlight.Elements.Add("/QuadPoints", quadPointsArray)
            highlight.Elements.Add("/Contents", New PdfString(commentText))
            highlight.Elements.Add("/T", New PdfString(author))
            highlight.Elements.Add("/CreationDate", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
            highlight.Elements.Add("/M", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
            highlight.Elements.Add("/NM", New PdfString("RI_Comment_" & commentId.ToString()))
            highlight.Elements.Add("/F", New PdfInteger(4))  ' Print flag
            highlight.Elements.Add("/C", APPdfMakeColorArray(pdfDoc, 1.0, 0.85, 0.0)) ' Yellow highlight
            highlight.Elements.Add("/CA", New PdfReal(0.4))  ' Semi-transparent highlight opacity

            ' Create the /Popup annotation
            Dim popupX = Math.Min(overallMaxX + 5, page.Width.Point - 210)
            Dim popupY = Math.Max(overallMinY - 100, 5)
            Dim popup As New PdfDictionary(pdfDoc)
            popup.Elements.Add("/Type", New PdfName("/Annot"))
            popup.Elements.Add("/Subtype", New PdfName("/Popup"))
            popup.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, popupX, popupY, popupX + 200, popupY + 100))
            popup.Elements.Add("/Open", New PdfBoolean(False))

            ' Register both objects so PdfSharp assigns indirect references
            pdfDoc.Internals.AddObject(highlight)
            pdfDoc.Internals.AddObject(popup)

            ' Cross-reference highlight ↔ popup
            highlight.Elements.Add("/Popup", popup.Reference)
            popup.Elements.Add("/Parent", highlight.Reference)

            ' Add both annotations to the page's /Annots array
            Dim annots = page.Elements.GetArray("/Annots")
            If annots Is Nothing Then
                annots = New PdfArray(pdfDoc)
                page.Elements.Add("/Annots", annots)
            End If
            annots.Elements.Add(highlight.Reference)
            annots.Elements.Add(popup.Reference)
        Next
    End Sub

    ''' <summary>
    ''' Inserts a /Text (sticky note) annotation at a specific position on a page.
    ''' Used for unmatched comments that couldn't be placed on their referenced text.
    ''' Each sticky note gets a unique position so they are individually selectable.
    ''' </summary>
    Private Shared Sub APPdfInsertStickyNoteAnnotation(pdfDoc As PdfDocument,
                                                        pageIndex As Integer,
                                                        x As Double, y As Double,
                                                        commentText As String,
                                                        author As String,
                                                        commentId As Integer)
        If pageIndex < 0 OrElse pageIndex >= pdfDoc.Pages.Count Then Return

        Dim page = pdfDoc.Pages(pageIndex)

        ' Create the /Text annotation (sticky note icon)
        Dim stickyNote As New PdfDictionary(pdfDoc)
        stickyNote.Elements.Add("/Type", New PdfName("/Annot"))
        stickyNote.Elements.Add("/Subtype", New PdfName("/Text"))
        stickyNote.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, x, y, x + AP_PdfStickyNoteSize, y + AP_PdfStickyNoteSize))
        stickyNote.Elements.Add("/Contents", New PdfString(commentText))
        stickyNote.Elements.Add("/T", New PdfString(author))
        stickyNote.Elements.Add("/Name", New PdfName("/Comment"))  ' Comment icon style
        stickyNote.Elements.Add("/CreationDate", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
        stickyNote.Elements.Add("/M", New PdfString(APPdfFormatDate(DateTime.UtcNow)))
        stickyNote.Elements.Add("/NM", New PdfString("RI_Unmatched_" & commentId.ToString()))
        stickyNote.Elements.Add("/F", New PdfInteger(4))  ' Print flag
        stickyNote.Elements.Add("/C", APPdfMakeColorArray(pdfDoc, 1.0, 0.6, 0.0))  ' Orange for unmatched
        stickyNote.Elements.Add("/Open", New PdfBoolean(False))

        ' Create associated popup
        Dim popup As New PdfDictionary(pdfDoc)
        popup.Elements.Add("/Type", New PdfName("/Annot"))
        popup.Elements.Add("/Subtype", New PdfName("/Popup"))
        popup.Elements.Add("/Rect", APPdfMakeRect(pdfDoc, x - 200, y - 100, x, y))
        popup.Elements.Add("/Open", New PdfBoolean(False))

        pdfDoc.Internals.AddObject(stickyNote)
        pdfDoc.Internals.AddObject(popup)

        stickyNote.Elements.Add("/Popup", popup.Reference)
        popup.Elements.Add("/Parent", stickyNote.Reference)

        Dim annots = page.Elements.GetArray("/Annots")
        If annots Is Nothing Then
            annots = New PdfArray(pdfDoc)
            page.Elements.Add("/Annots", annots)
        End If
        annots.Elements.Add(stickyNote.Reference)
        annots.Elements.Add(popup.Reference)
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PDF HELPER FUNCTIONS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Creates a PDF /Rect array [x1 y1 x2 y2].
    ''' </summary>
    Private Shared Function APPdfMakeRect(pdfDoc As PdfDocument,
                                          x1 As Double, y1 As Double,
                                          x2 As Double, y2 As Double) As PdfArray
        Dim arr As New PdfArray(pdfDoc)
        arr.Elements.Add(New PdfReal(x1))
        arr.Elements.Add(New PdfReal(y1))
        arr.Elements.Add(New PdfReal(x2))
        arr.Elements.Add(New PdfReal(y2))
        Return arr
    End Function

    ''' <summary>
    ''' Creates a PDF color array [r g b] with values in 0.0–1.0 range.
    ''' </summary>
    Private Shared Function APPdfMakeColorArray(pdfDoc As PdfDocument,
                                                r As Double, g As Double, b As Double) As PdfArray
        Dim arr As New PdfArray(pdfDoc)
        arr.Elements.Add(New PdfReal(r))
        arr.Elements.Add(New PdfReal(g))
        arr.Elements.Add(New PdfReal(b))
        Return arr
    End Function

    ''' <summary>
    ''' Formats a DateTime as a PDF date string: D:YYYYMMDDHHmmSSOHH'mm'
    ''' </summary>
    Private Shared Function APPdfFormatDate(dt As DateTime) As String
        Return "D:" & dt.ToString("yyyyMMddHHmmss") & "Z"
    End Function

End Class