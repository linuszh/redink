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

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASSES (local to AutoPilot, avoid collision with Word add-in)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Stores a text node and its original content for a single run.
    ''' </summary>
    Private Class APTextRunInfo
        Public Property TextNode As System.Xml.XmlNode
        Public Property OriginalText As String
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

            Dim success = Await APProcessBatches(paragraphs, instruction, ct)
            If Not success Then Return False

            APApplyTranslations(paragraphs)
            xmlDoc.Save(documentXmlPath)

            ' Process headers, footers, comments, footnotes, endnotes
            Await APProcessSubParts(tempDir, instruction, ct)

            ' Repack DOCX
            File.Delete(outputPath)
            ZipFile.CreateFromDirectory(tempDir, outputPath, CompressionLevel.Optimal, False)

            Return True

        Catch ex As System.Exception
            Debug.WriteLine("ProcessDocxForAutoPilot error: " & ex.Message)
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

            Dim textNodes = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim fullTextBuilder As New StringBuilder()

            For Each textNode As System.Xml.XmlNode In textNodes
                paraInfo.TextRuns.Add(New APTextRunInfo() With {
                    .TextNode = textNode,
                    .OriginalText = textNode.InnerText
                })
                fullTextBuilder.Append(textNode.InnerText)
            Next

            paraInfo.FullText = fullTextBuilder.ToString()
            paraInfo.IsEmpty = String.IsNullOrWhiteSpace(paraInfo.FullText)

            ' Build marker-annotated text for multi-run paragraphs
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
    ''' Example: "Hello" + " world" → "Hello| world"
    ''' </summary>
    Private Shared Function APBuildMarkerAnnotatedText(textRuns As List(Of APTextRunInfo)) As String
        Dim sb As New StringBuilder()
        For i As Integer = 0 To textRuns.Count - 1
            If i > 0 Then sb.Append(AP_RunBoundaryMarker)
            sb.Append(textRuns(i).OriginalText)
        Next
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

        Dim batchIndex As Integer = 0

        While batchIndex < processable.Count
            ct.ThrowIfCancellationRequested()

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

            ' Build prompt
            Dim promptBuilder As New StringBuilder()

            ' Context Before
            Dim contextBeforeStart = Math.Max(0, batchStart - AP_ContextBefore)
            If contextBeforeStart < batchStart Then
                promptBuilder.AppendLine("[CONTEXT BEFORE - for reference only]")
                For j = contextBeforeStart To batchStart - 1
                    Dim contextText = If(processable(j).TranslatedText, processable(j).FullText)
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
                    promptBuilder.AppendLine(processable(j).FullText)
                Next
            End If

            ' Call LLM
            Dim llmResponse = Await LLM(systemPrompt, promptBuilder.ToString(),
                                         UseSecondAPI:=_apUseSecondApi,
                                         HideSplash:=True, EnsureUI:=False,
                                         cancellationToken:=ct)

            If String.IsNullOrWhiteSpace(llmResponse) Then Return False

            ' Parse response
            APParseResponse(llmResponse, processable, batchStart, batchEnd)

            batchIndex = batchEnd + 1
        End While

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
    ''' </summary>
    Private Sub APApplyTranslations(paragraphs As List(Of APParagraphInfo))
        For Each para In paragraphs
            If para.IsEmpty OrElse String.IsNullOrEmpty(para.TranslatedText) Then Continue For
            If para.TextRuns.Count = 0 Then Continue For

            Dim translatedText = para.TranslatedText

            ' Single run: simple replacement
            If para.TextRuns.Count = 1 Then
                APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                Continue For
            End If

            ' Multi-run: try marker-based distribution first (matching Word add-in)
            If APTryApplyMarkerBasedDistribution(para, translatedText) Then
                Continue For
            End If

            ' Fallback: proportional distribution
            Dim totalOrigLen = para.FullText.Length
            If totalOrigLen = 0 Then
                APSetTextNode(para.TextRuns(0).TextNode, translatedText)
                For idx = 1 To para.TextRuns.Count - 1
                    APSetTextNode(para.TextRuns(idx).TextNode, "")
                Next
                Continue For
            End If

            Dim translatedLen = translatedText.Length
            Dim currentPos = 0
            Dim cumulativeOriginal = 0

            For runIdx = 0 To para.TextRuns.Count - 1
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

                    ' Try to break at word boundary
                    Dim endPos = targetEndPos
                    If endPos < translatedLen AndAlso endPos > currentPos Then
                        If translatedText(endPos) = " "c Then
                            endPos += 1
                        Else
                            Dim foundSpace = False
                            For searchPos = endPos To Math.Min(endPos + 15, translatedLen - 1)
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
        Next
    End Sub

    ''' <summary>
    ''' Attempts marker-based distribution (matching Word add-in's TryApplyMarkerBasedDistribution).
    ''' If the LLM returned text with exactly the right number of | markers, split on them
    ''' and assign each segment to the corresponding run. Returns True if successful.
    ''' </summary>
    Private Function APTryApplyMarkerBasedDistribution(para As APParagraphInfo, translatedText As String) As Boolean
        ' Only applicable if the paragraph was sent with markers
        If para.MarkerText Is Nothing Then Return False

        ' Count expected markers (one fewer than runs)
        Dim expectedMarkers = para.TextRuns.Count - 1
        If expectedMarkers <= 0 Then Return False

        ' Count actual markers in translated text
        Dim markerCount = translatedText.Count(Function(c) c = AP_RunBoundaryMarker(0))
        If markerCount <> expectedMarkers Then Return False

        ' Split on the marker
        Dim segments = translatedText.Split(AP_RunBoundaryMarker(0))
        If segments.Length <> para.TextRuns.Count Then Return False

        ' Apply each segment to its run
        For i As Integer = 0 To para.TextRuns.Count - 1
            APSetTextNode(para.TextRuns(i).TextNode, segments(i))
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