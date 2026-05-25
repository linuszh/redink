' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.SandboxedReaders.vb
' Purpose: Sandboxed (COM-free) text extraction from Office documents and
'          embedded mail files. Uses ZIP + XML parsing only — no Office interop.
'
' Output Compatibility:
'  - DOCX: paragraph-per-line plain text (matches ReadWordDocument / ExtractWordText output)
'  - XLSX: "{addr}\tFORMULA:={formula}\tVALUE: {value}" per cell with "=== Sheet: {name} ==="
'           headers (matches ExtractExcelText output format exactly)
'  - PPTX: "=== Slide {n} ===" + shape text + "--- Notes ---" per slide
'           (matches ExtractPowerPointText output format exactly)
'  - EML:  structured header + body text (matches ParseEmlAsText output)
'  - MSG:  structured text via callback, or heuristic UTF-16LE extraction fallback
'
' Supported Formats:
'  - .docx  — WordprocessingML (OpenXML ZIP → word/document.xml)
'  - .xlsx  — SpreadsheetML (OpenXML ZIP → shared strings + sheet XML)
'  - .pptx  — PresentationML (OpenXML ZIP → slide XML via DrawingML)
'  - .msg   — Outlook message (callback for MAPI, with heuristic fallback)
'  - .eml   — RFC 2822 mail (plain text parsing with MIME boundary handling)
'
' External Dependencies:
'  - System.IO.Compression (ZipArchive)
'  - System.Xml (XmlDocument / XmlNamespaceManager)
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Cached value for allowing legacy document files. Set during LoadConfig.
        ''' </summary>
        Public Shared INI_AllowLegacyDocFiles_Cached As Boolean = False

        ' ═══════════════════════════════════════════════════════════════════════
        '  DOCX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Private Const SB_WordNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        ''' <summary>
        ''' Extracts plain text from a .docx file without COM interop.
        ''' Output: one line per paragraph from the main body.
        ''' When <see cref="DocxIncludeHeaderFooterFootnotes"/> is <c>True</c>, headers, footers,
        ''' footnotes and endnotes are appended as delimited sections.
        ''' </summary>
        ''' <param name="docxPath">Absolute path to the .docx file.</param>
        ''' <returns>Extracted plain text, or an error string on failure.</returns>
        Public Shared Function oldReadDocxSandboxed(docxPath As String) As String
            If String.IsNullOrWhiteSpace(docxPath) OrElse Not File.Exists(docxPath) Then
                Return "Error: File not found."
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_docx_" & Guid.NewGuid().ToString("N"))
            Try
                ZipFile.ExtractToDirectory(docxPath, tempDir)

                Dim wordDir = Path.Combine(tempDir, "word")
                Dim documentXmlPath = Path.Combine(wordDir, "document.xml")
                If Not File.Exists(documentXmlPath) Then Return "Error: Not a valid .docx file (missing word/document.xml)."

                Dim xmlDoc As New XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(documentXmlPath)

                Dim nsMgr As New XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("w", SB_WordNs)

                Dim sb As New StringBuilder(4096)

                ' ── Main body text ──
                ' Collect footnote/endnote reference IDs inline so we can insert markers
                Dim paragraphs = xmlDoc.SelectNodes("//w:body/w:p", nsMgr)
                If paragraphs IsNot Nothing Then
                    For Each para As XmlNode In paragraphs
                        Dim paraText As New StringBuilder()
                        Dim runs = para.SelectNodes(".//w:r", nsMgr)
                        If runs IsNot Nothing Then
                            For Each run As XmlNode In runs
                                ' Check for footnote / endnote references within this run
                                If DocxIncludeHeaderFooterFootnotes Then
                                    Dim fnRef = run.SelectSingleNode("w:footnoteReference", nsMgr)
                                    If fnRef IsNot Nothing Then
                                        Dim fnId = fnRef.Attributes?("w:id")?.Value
                                        If Not String.IsNullOrWhiteSpace(fnId) AndAlso fnId <> "0" Then
                                            paraText.Append($" [Footnote {fnId}]")
                                        End If
                                    End If
                                    Dim enRef = run.SelectSingleNode("w:endnoteReference", nsMgr)
                                    If enRef IsNot Nothing Then
                                        Dim enId = enRef.Attributes?("w:id")?.Value
                                        If Not String.IsNullOrWhiteSpace(enId) AndAlso enId <> "0" Then
                                            paraText.Append($" [Endnote {enId}]")
                                        End If
                                    End If
                                End If

                                ' Collect <w:t> text nodes from this run
                                Dim tNodes = run.SelectNodes("w:t", nsMgr)
                                If tNodes IsNot Nothing Then
                                    For Each tNode As XmlNode In tNodes
                                        paraText.Append(tNode.InnerText)
                                    Next
                                End If
                            Next
                        End If

                        If paraText.Length > 0 Then
                            sb.AppendLine(paraText.ToString())
                        Else
                            ' Empty paragraph = blank line (matches COM Content.Text behavior)
                            sb.AppendLine()
                        End If
                    Next
                End If

                ' ── Optional: headers, footers, footnotes, endnotes ──
                If DocxIncludeHeaderFooterFootnotes AndAlso Directory.Exists(wordDir) Then
                    ' Headers
                    ExtractDocxSubParts(wordDir, "header*.xml", "Header", sb)

                    ' Footers
                    ExtractDocxSubParts(wordDir, "footer*.xml", "Footer", sb)

                    ' Footnotes (skip id="0" = separator, id="-1" = continuation separator)
                    ExtractDocxNotes(wordDir, "footnotes.xml", "w:footnote", "Footnote", sb)

                    ' Endnotes (skip id="0" and id="-1")
                    ExtractDocxNotes(wordDir, "endnotes.xml", "w:endnote", "Endnote", sb)
                End If

                Dim result = sb.ToString().TrimEnd()
                Return If(String.IsNullOrWhiteSpace(result), "Error: No text content found in .docx.", result)

            Catch ex As Exception
                Return $"Error reading .docx: {ex.Message}"
            Finally
                Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                Catch : End Try
            End Try
        End Function


        Public Shared Function ReadDocxSandboxed(docxPath As String) As String
            If System.String.IsNullOrWhiteSpace(docxPath) OrElse Not System.IO.File.Exists(docxPath) Then
                Return "Error: File not found."
            End If

            Dim tempDir As String =
        System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "ri_docx_" & System.Guid.NewGuid().ToString("N")
        )

            Try
                System.IO.Compression.ZipFile.ExtractToDirectory(docxPath, tempDir)

                Dim wordDir As String = System.IO.Path.Combine(tempDir, "word")
                Dim documentXmlPath As String = System.IO.Path.Combine(wordDir, "document.xml")

                If Not System.IO.File.Exists(documentXmlPath) Then
                    Return "Error: Not a valid .docx file (missing word/document.xml)."
                End If

                Dim xmlDoc As New System.Xml.XmlDocument()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.Load(documentXmlPath)

                Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
                nsMgr.AddNamespace("w", SB_WordNs)

                Dim sb As New System.Text.StringBuilder(4096)

                ' ── Main body content in original order: paragraphs and tables ──
                Dim bodyNode As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
                Dim tableIndex As Integer = 0

                If bodyNode IsNot Nothing Then
                    For Each childNode As System.Xml.XmlNode In bodyNode.ChildNodes

                        If childNode.NamespaceURI <> SB_WordNs Then
                            Continue For
                        End If

                        Select Case childNode.LocalName

                            Case "p"
                                Dim paraText As String = ExtractDocxParagraphText(childNode, nsMgr)

                                If paraText.Length > 0 Then
                                    sb.AppendLine(paraText)
                                Else
                                    ' Empty paragraph = blank line
                                    sb.AppendLine()
                                End If

                            Case "tbl"
                                tableIndex += 1
                                AppendDocxTableForLlm(childNode, nsMgr, sb, tableIndex, 0)

                        End Select
                    Next
                End If

                ' ── Optional: headers, footers, footnotes, endnotes ──
                If DocxIncludeHeaderFooterFootnotes AndAlso System.IO.Directory.Exists(wordDir) Then

                    ' Headers
                    ExtractDocxSubParts(wordDir, "header*.xml", "Header", sb)

                    ' Footers
                    ExtractDocxSubParts(wordDir, "footer*.xml", "Footer", sb)

                    ' Footnotes, skip id="0" = separator, id="-1" = continuation separator
                    ExtractDocxNotes(wordDir, "footnotes.xml", "w:footnote", "Footnote", sb)

                    ' Endnotes, skip id="0" and id="-1"
                    ExtractDocxNotes(wordDir, "endnotes.xml", "w:endnote", "Endnote", sb)
                End If

                Dim result As String = sb.ToString().TrimEnd()

                Return If(
            System.String.IsNullOrWhiteSpace(result),
            "Error: No text content found in .docx.",
            result
        )

            Catch ex As System.Exception
                Return "Error reading .docx: " & ex.Message

            Finally
                Try
                    If System.IO.Directory.Exists(tempDir) Then
                        System.IO.Directory.Delete(tempDir, True)
                    End If
                Catch
                End Try
            End Try
        End Function


        Private Shared Function ExtractDocxParagraphText(
    paraNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As String

            Dim paraText As New System.Text.StringBuilder()

            Dim runs As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:r", nsMgr)

            If runs IsNot Nothing Then
                For Each runNode As System.Xml.XmlNode In runs

                    If DocxIncludeHeaderFooterFootnotes Then
                        Dim fnRef As System.Xml.XmlNode = runNode.SelectSingleNode("w:footnoteReference", nsMgr)

                        If fnRef IsNot Nothing Then
                            Dim fnId As String = GetWordAttributeValue(fnRef, "id")

                            If Not System.String.IsNullOrWhiteSpace(fnId) AndAlso fnId <> "0" Then
                                paraText.Append(" [Footnote " & fnId & "]")
                            End If
                        End If

                        Dim enRef As System.Xml.XmlNode = runNode.SelectSingleNode("w:endnoteReference", nsMgr)

                        If enRef IsNot Nothing Then
                            Dim enId As String = GetWordAttributeValue(enRef, "id")

                            If Not System.String.IsNullOrWhiteSpace(enId) AndAlso enId <> "0" Then
                                paraText.Append(" [Endnote " & enId & "]")
                            End If
                        End If
                    End If

                    For Each runChild As System.Xml.XmlNode In runNode.ChildNodes

                        If runChild.NamespaceURI <> SB_WordNs Then
                            Continue For
                        End If

                        Select Case runChild.LocalName

                            Case "t"
                                paraText.Append(runChild.InnerText)

                            Case "tab"
                                paraText.Append(vbTab)

                            Case "br", "cr"
                                paraText.AppendLine()

                            Case "noBreakHyphen"
                                paraText.Append(ChrW(&H2011))

                            Case "softHyphen"
                                paraText.Append(ChrW(&HAD))

                        End Select
                    Next
                Next
            End If

            Return paraText.ToString().Trim()
        End Function


        Private Shared Sub AppendDocxTableForLlm(
    tableNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager,
    sb As System.Text.StringBuilder,
    tableIndex As Integer,
    nestingLevel As Integer
)

            Dim indent As String = New System.String(" "c, nestingLevel * 2)
            Dim tableNumberText As String = tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)

            sb.AppendLine()
            sb.AppendLine(indent & "[Table " & tableNumberText & "]")

            Dim rows As System.Xml.XmlNodeList = tableNode.SelectNodes("w:tr", nsMgr)

            If rows Is Nothing OrElse rows.Count = 0 Then
                sb.AppendLine(indent & "[Empty table]")
                sb.AppendLine(indent & "[/Table " & tableNumberText & "]")
                sb.AppendLine()
                Return
            End If

            Dim rowIndex As Integer = 0

            For Each rowNode As System.Xml.XmlNode In rows
                rowIndex += 1

                Dim cells As System.Xml.XmlNodeList = rowNode.SelectNodes("w:tc", nsMgr)

                If cells Is Nothing OrElse cells.Count = 0 Then
                    sb.AppendLine(
                indent &
                "Row " &
                rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                ": [empty row]"
            )

                    Continue For
                End If

                Dim visualColumnIndex As Integer = 1
                Dim physicalCellIndex As Integer = 0

                For Each cellNode As System.Xml.XmlNode In cells
                    physicalCellIndex += 1

                    Dim gridSpan As Integer = GetDocxGridSpan(cellNode, nsMgr)
                    Dim verticalMerge As String = GetDocxVerticalMerge(cellNode, nsMgr)
                    Dim cellText As String = ExtractDocxCellText(cellNode, nsMgr, tableIndex, nestingLevel + 1)

                    Dim rowText As String = rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim physicalCellText As String = physicalCellIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim startColumnText As String = visualColumnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim endColumnText As String = (visualColumnIndex + gridSpan - 1).ToString(System.Globalization.CultureInfo.InvariantCulture)

                    Dim cellLabel As String =
                indent &
                "Row " &
                rowText &
                ", Cell " &
                physicalCellText &
                ", Column " &
                startColumnText

                    If gridSpan > 1 Then
                        cellLabel &= "-" & endColumnText & " spanning " & gridSpan.ToString(System.Globalization.CultureInfo.InvariantCulture) & " columns"
                    End If

                    If Not System.String.IsNullOrWhiteSpace(verticalMerge) Then
                        If verticalMerge = "restart" Then
                            cellLabel &= ", starts vertical merge"
                        Else
                            cellLabel &= ", continues vertical merge from row above"
                        End If
                    End If

                    sb.AppendLine(cellLabel & ": " & cellText)

                    visualColumnIndex += gridSpan
                Next
            Next

            sb.AppendLine(indent & "[/Table " & tableNumberText & "]")
            sb.AppendLine()
        End Sub


        Private Shared Function ExtractDocxCellText(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager,
    tableIndex As Integer,
    nestingLevel As Integer
) As String

            Dim parts As New System.Collections.Generic.List(Of String)()
            Dim nestedTableIndex As Integer = 0

            For Each childNode As System.Xml.XmlNode In cellNode.ChildNodes

                If childNode.NamespaceURI <> SB_WordNs Then
                    Continue For
                End If

                Select Case childNode.LocalName

                    Case "p"
                        Dim paragraphText As String = ExtractDocxParagraphText(childNode, nsMgr)

                        If Not System.String.IsNullOrWhiteSpace(paragraphText) Then
                            parts.Add(paragraphText)
                        End If

                    Case "tbl"
                        nestedTableIndex += 1

                        Dim nestedBuilder As New System.Text.StringBuilder()

                        AppendDocxTableForLlm(
                    childNode,
                    nsMgr,
                    nestedBuilder,
                    tableIndex * 1000 + nestedTableIndex,
                    nestingLevel
                )

                        parts.Add(nestedBuilder.ToString().Trim())

                End Select
            Next

            If parts.Count = 0 Then
                Return "[empty]"
            End If

            Return System.String.Join(" | ", parts)
        End Function


        Private Shared Function GetDocxGridSpan(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As Integer

            Dim gridSpanNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:gridSpan", nsMgr)

            If gridSpanNode Is Nothing Then
                Return 1
            End If

            Dim valueText As String = GetWordAttributeValue(gridSpanNode, "val")
            Dim result As Integer

            If System.Int32.TryParse(valueText, result) AndAlso result > 1 Then
                Return result
            End If

            Return 1
        End Function


        Private Shared Function GetDocxVerticalMerge(
    cellNode As System.Xml.XmlNode,
    nsMgr As System.Xml.XmlNamespaceManager
) As String

            Dim vMergeNode As System.Xml.XmlNode = cellNode.SelectSingleNode("w:tcPr/w:vMerge", nsMgr)

            If vMergeNode Is Nothing Then
                Return System.String.Empty
            End If

            Dim valueText As String = GetWordAttributeValue(vMergeNode, "val")

            If System.String.IsNullOrWhiteSpace(valueText) Then
                Return "continue"
            End If

            Return valueText
        End Function


        Private Shared Function GetWordAttributeValue(
    node As System.Xml.XmlNode,
    localName As String
) As String

            If node Is Nothing OrElse node.Attributes Is Nothing Then
                Return System.String.Empty
            End If

            Dim attr As System.Xml.XmlNode = node.Attributes.GetNamedItem(localName, SB_WordNs)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            attr = node.Attributes.GetNamedItem("w:" & localName)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            attr = node.Attributes.GetNamedItem(localName)

            If attr IsNot Nothing Then
                Return attr.Value
            End If

            Return System.String.Empty
        End Function


        ''' <summary>
        ''' Extracts text from DOCX sub-parts matching a file pattern (e.g., <c>header*.xml</c>, <c>footer*.xml</c>).
        ''' Each file produces a labeled section with its paragraphs.
        ''' </summary>
        Private Shared Sub ExtractDocxSubParts(wordDir As String, filePattern As String,
                                               sectionLabel As String, sb As StringBuilder)
            Try
                Dim files = Directory.GetFiles(wordDir, filePattern).OrderBy(Function(f) f).ToArray()
                If files.Length = 0 Then Return

                For Each filePath In files
                    Dim partDoc As New XmlDocument()
                    partDoc.PreserveWhitespace = True
                    partDoc.Load(filePath)

                    Dim partNs As New XmlNamespaceManager(partDoc.NameTable)
                    partNs.AddNamespace("w", SB_WordNs)

                    Dim partParagraphs = partDoc.SelectNodes("//w:p", partNs)
                    If partParagraphs Is Nothing OrElse partParagraphs.Count = 0 Then Continue For

                    ' Collect text first to check if there's any real content
                    Dim partText As New StringBuilder()
                    For Each para As XmlNode In partParagraphs
                        Dim tNodes = para.SelectNodes(".//w:t", partNs)
                        If tNodes IsNot Nothing AndAlso tNodes.Count > 0 Then
                            Dim paraLine As New StringBuilder()
                            For Each tNode As XmlNode In tNodes
                                paraLine.Append(tNode.InnerText)
                            Next
                            Dim lineText = paraLine.ToString().Trim()
                            If lineText.Length > 0 Then
                                partText.AppendLine(lineText)
                            End If
                        End If
                    Next

                    If partText.Length > 0 Then
                        ' Derive a display label: "Header 1", "Footer 2", etc.
                        Dim fileLabel = Path.GetFileNameWithoutExtension(filePath)
                        ' Extract trailing number from e.g. "header1" → "1"
                        Dim numPart = ""
                        For i = fileLabel.Length - 1 To 0 Step -1
                            If Char.IsDigit(fileLabel(i)) Then
                                numPart = fileLabel(i) & numPart
                            Else
                                Exit For
                            End If
                        Next

                        sb.AppendLine()
                        sb.AppendLine($"--- {sectionLabel}{If(numPart.Length > 0, " " & numPart, "")} ---")
                        sb.Append(partText.ToString().TrimEnd())
                        sb.AppendLine()
                    End If
                Next
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Extracts text from DOCX footnotes or endnotes XML. Skips system notes (id 0 and -1)
        ''' which are separator/continuation placeholders. Each note is output with its ID
        ''' so it can be cross-referenced with <c>[Footnote n]</c> / <c>[Endnote n]</c> markers
        ''' inserted into the body text.
        ''' </summary>
        Private Shared Sub ExtractDocxNotes(wordDir As String, fileName As String,
                                             noteElementName As String, sectionLabel As String,
                                             sb As StringBuilder)
            Dim filePath = Path.Combine(wordDir, fileName)
            If Not File.Exists(filePath) Then Return

            Try
                Dim notesDoc As New XmlDocument()
                notesDoc.PreserveWhitespace = True
                notesDoc.Load(filePath)

                Dim notesNs As New XmlNamespaceManager(notesDoc.NameTable)
                notesNs.AddNamespace("w", SB_WordNs)

                Dim noteNodes = notesDoc.SelectNodes($"//{ noteElementName}", notesNs)
                If noteNodes Is Nothing OrElse noteNodes.Count = 0 Then Return

                Dim notesCollected As New StringBuilder()
                Dim noteCount As Integer = 0

                For Each noteNode As XmlNode In noteNodes
                    ' Skip system separator/continuation notes (id="0" or id="-1")
                    Dim noteId = noteNode.Attributes?("w:id")?.Value
                    If noteId = "0" OrElse noteId = "-1" Then Continue For

                    Dim noteParagraphs = noteNode.SelectNodes("w:p", notesNs)
                    If noteParagraphs Is Nothing OrElse noteParagraphs.Count = 0 Then Continue For

                    Dim noteText As New StringBuilder()
                    For Each para As XmlNode In noteParagraphs
                        ' Skip the footnoteRef/endnoteRef marker runs (the auto-number "1", "2" etc.)
                        Dim tNodes = para.SelectNodes(".//w:r[not(w:footnoteRef) and not(w:endnoteRef)]/w:t", notesNs)
                        If tNodes IsNot Nothing AndAlso tNodes.Count > 0 Then
                            For Each tNode As XmlNode In tNodes
                                noteText.Append(tNode.InnerText)
                            Next
                        End If
                    Next

                    Dim trimmedNote = noteText.ToString().Trim()
                    If trimmedNote.Length > 0 Then
                        notesCollected.AppendLine($"  [{sectionLabel} {If(noteId, "?")}] {trimmedNote}")
                        noteCount += 1
                    End If
                Next

                If noteCount > 0 Then
                    sb.AppendLine()
                    sb.AppendLine($"--- {sectionLabel}s ---")
                    sb.Append(notesCollected.ToString().TrimEnd())
                    sb.AppendLine()
                End If
            Catch
            End Try
        End Sub


        ' ═══════════════════════════════════════════════════════════════════════
        '  XLSX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Private Const SB_XlsxNs As String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        Private Const SB_XlsxRelNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ''' <summary>
        ''' Represents a readable worksheet discovered in an .xlsx workbook.
        ''' </summary>
        Private Structure XlsxSheetEntry
            Public Name As String
            Public SheetXmlPath As String
        End Structure

        ''' <summary>
        ''' Represents a worksheet declared in workbook.xml, whether or not its XML part
        ''' could be resolved immediately.
        ''' </summary>
        Private Structure XlsxDeclaredSheetEntry
            Public Name As String
            Public SheetXmlPath As String
        End Structure

        Public Const XlsxSelectionCancelledMarker As String = "__RI_XLSX_SELECTION_CANCELLED__"

        ''' <summary>
        ''' Extracts text from an .xlsx file without COM interop.
        ''' Output format matches <c>ExtractExcelText</c>: <c>{addr}\tFORMULA:={formula}\tVALUE: {value}</c> per cell,
        ''' with <c>=== Sheet: {name} ===</c> headers.
        ''' </summary>
        ''' <param name="xlsxPath">Absolute path to the .xlsx file.</param>
        ''' <param name="silent">
        ''' When <c>True</c>, suppresses worksheet-selection UI and always loads all readable worksheets.
        ''' </param>
        ''' <param name="askWorksheetSelection">
        ''' When <c>True</c> and <paramref name="silent"/> is <c>False</c>, prompts the user via
        ''' <see cref="SelectValue(IEnumerable(Of SelectionItem), Integer, String, String)"/> to choose
        ''' either all readable worksheets or one specific worksheet when the workbook contains multiple sheets.
        ''' </param>
        ''' <returns>
        ''' Extracted text representation of the selected worksheet set, or an error string on failure.
        ''' Returns an empty string when worksheet selection is canceled.
        ''' </returns>
        Public Shared Function ReadXlsxSandboxed(xlsxPath As String,
                                                 Optional silent As Boolean = True,
                                                 Optional askWorksheetSelection As Boolean = False) As String
            If String.IsNullOrWhiteSpace(xlsxPath) OrElse Not File.Exists(xlsxPath) Then
                Return "Error: File not found."
            End If

            If Not EnsureClosedWorkbookForSandboxedRead(xlsxPath, silent) Then
                Return "Error: The workbook is open in Excel."
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_xlsx_" & Guid.NewGuid().ToString("N"))

            Try
                ZipFile.ExtractToDirectory(xlsxPath, tempDir)

                ' ── Load shared strings ──
                Dim sharedStrings As New List(Of String)()
                Dim sstPath = Path.Combine(tempDir, "xl", "sharedStrings.xml")
                If File.Exists(sstPath) Then
                    Dim sstDoc As New XmlDocument()
                    sstDoc.Load(sstPath)
                    Dim sstNs As New XmlNamespaceManager(sstDoc.NameTable)
                    sstNs.AddNamespace("x", SB_XlsxNs)
                    Dim siNodes = sstDoc.SelectNodes("//x:si", sstNs)
                    If siNodes IsNot Nothing Then
                        For Each si As XmlNode In siNodes
                            Dim tNodes = si.SelectNodes(".//x:t", sstNs)
                            Dim cellText As New StringBuilder()
                            If tNodes IsNot Nothing Then
                                For Each tNode As XmlNode In tNodes
                                    cellText.Append(tNode.InnerText)
                                Next
                            End If
                            sharedStrings.Add(cellText.ToString())
                        Next
                    End If
                End If

                ' ── Discover sheet names from workbook.xml ──
                Dim wbPath = Path.Combine(tempDir, "xl", "workbook.xml")
                If Not File.Exists(wbPath) Then Return "Error: Not a valid .xlsx file (missing xl/workbook.xml)."

                Dim wbDoc As New XmlDocument()
                wbDoc.Load(wbPath)
                Dim wbNs As New XmlNamespaceManager(wbDoc.NameTable)
                wbNs.AddNamespace("x", SB_XlsxNs)
                wbNs.AddNamespace("r", SB_XlsxRelNs)

                Dim sheetNodes = wbDoc.SelectNodes("//x:sheets/x:sheet", wbNs)
                If sheetNodes Is Nothing OrElse sheetNodes.Count = 0 Then Return "Error: No sheets found in workbook."

                ' ── Map rId → file path via workbook.xml.rels ──
                Dim relsPath = Path.Combine(tempDir, "xl", "_rels", "workbook.xml.rels")
                Dim ridMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                If File.Exists(relsPath) Then
                    Dim relsDoc As New XmlDocument()
                    relsDoc.Load(relsPath)
                    If relsDoc.DocumentElement IsNot Nothing Then
                        For Each relNode As XmlNode In relsDoc.DocumentElement.ChildNodes
                            If relNode.Attributes Is Nothing Then Continue For
                            Dim id = relNode.Attributes("Id")
                            Dim target = relNode.Attributes("Target")
                            If id IsNot Nothing AndAlso target IsNot Nothing Then
                                ridMap(id.Value) = target.Value
                            End If
                        Next
                    End If
                End If

                Dim declaredSheets As New List(Of XlsxDeclaredSheetEntry)()
                Dim availableSheets As New List(Of XlsxSheetEntry)()
                Dim sheetIdx As Integer = 0

                For Each sheetNode As XmlNode In sheetNodes
                    sheetIdx += 1
                    Dim sheetName = If(sheetNode.Attributes("name")?.Value, "Sheet")
                    Dim rId = If(sheetNode.Attributes("r:id")?.Value, "")

                    Dim sheetXmlPath = ResolveSheetPath(tempDir, rId, ridMap, sheetIdx)

                    declaredSheets.Add(New XlsxDeclaredSheetEntry With {
                        .Name = sheetName,
                        .SheetXmlPath = sheetXmlPath
                    })

                    If Not String.IsNullOrWhiteSpace(sheetXmlPath) AndAlso File.Exists(sheetXmlPath) Then
                        availableSheets.Add(New XlsxSheetEntry With {
                            .Name = sheetName,
                            .SheetXmlPath = sheetXmlPath
                        })
                    End If
                Next

                If availableSheets.Count = 0 Then
                    Return "Error: No readable sheets found in workbook."
                End If

                Dim sheetsToRead As New List(Of XlsxSheetEntry)()
                For Each s In availableSheets
                    sheetsToRead.Add(s)
                Next

                If askWorksheetSelection AndAlso Not silent AndAlso declaredSheets.Count > 1 Then
                    Dim items As New List(Of SelectionItem) From {
                        New SelectionItem("All worksheets", -1)
                    }

                    For i As Integer = 0 To declaredSheets.Count - 1
                        items.Add(New SelectionItem(
                            declaredSheets(i).Name & " (worksheet " & (i + 1).ToString(Globalization.CultureInfo.InvariantCulture) & ")",
                            i + 1))
                    Next

                    Dim selectedValue As Integer = SelectValue(
                        items,
                        -1,
                        "This workbook contains multiple worksheets. Load all worksheets or only one worksheet?",
                        AN & " - Select Worksheet")

                    If selectedValue = 0 Then
                        Return XlsxSelectionCancelledMarker
                    End If

                    If selectedValue > 0 AndAlso selectedValue <= declaredSheets.Count Then
                        Dim selectedSheet = declaredSheets(selectedValue - 1)

                        If String.IsNullOrWhiteSpace(selectedSheet.SheetXmlPath) OrElse
                           Not File.Exists(selectedSheet.SheetXmlPath) Then
                            Return "Error: The selected worksheet could not be read."
                        End If

                        sheetsToRead.Clear()
                        sheetsToRead.Add(New XlsxSheetEntry With {
                            .Name = selectedSheet.Name,
                            .SheetXmlPath = selectedSheet.SheetXmlPath
                        })
                    End If
                End If

                Dim sb As New StringBuilder(4096)

                For Each sheet In sheetsToRead
                    sb.AppendLine("=== Sheet: " & sheet.Name & " ===")

                    Dim sheetDoc As New XmlDocument()
                    sheetDoc.Load(sheet.SheetXmlPath)
                    Dim sheetNs As New XmlNamespaceManager(sheetDoc.NameTable)
                    sheetNs.AddNamespace("x", SB_XlsxNs)

                    Dim cellNodes = sheetDoc.SelectNodes("//x:sheetData/x:row/x:c", sheetNs)
                    If cellNodes IsNot Nothing Then
                        For Each cellNode As XmlElement In cellNodes
                            Dim cellRef = cellNode.GetAttribute("r")
                            If String.IsNullOrEmpty(cellRef) Then Continue For

                            Dim cellType = cellNode.GetAttribute("t")
                            Dim vNode = cellNode.SelectSingleNode("x:v", sheetNs)
                            Dim fNode = cellNode.SelectSingleNode("x:f", sheetNs)

                            Dim displayValue As String = ""
                            Dim formulaStr As String = ""

                            If fNode IsNot Nothing Then
                                formulaStr = fNode.InnerText
                                displayValue = If(vNode?.InnerText, "")
                            ElseIf cellType = "s" Then
                                If vNode IsNot Nothing Then
                                    Dim ssIndex As Integer
                                    If Integer.TryParse(vNode.InnerText, ssIndex) AndAlso
                                       ssIndex >= 0 AndAlso ssIndex < sharedStrings.Count Then
                                        displayValue = sharedStrings(ssIndex)
                                    End If
                                End If
                            ElseIf cellType = "b" Then
                                displayValue = If(vNode?.InnerText = "1", "TRUE", "FALSE")
                            Else
                                displayValue = If(vNode?.InnerText, "")
                            End If

                            If String.IsNullOrEmpty(formulaStr) AndAlso String.IsNullOrEmpty(displayValue) Then Continue For

                            sb.Append(cellRef)
                            sb.Append(vbTab)
                            sb.Append("FORMULA:")
                            If Not String.IsNullOrEmpty(formulaStr) Then
                                sb.Append("=")
                                sb.Append(formulaStr)
                            End If
                            sb.Append(vbTab)
                            sb.Append("VALUE: ")
                            sb.AppendLine(displayValue)
                        Next
                    End If

                    sb.AppendLine()
                Next

                Dim result = sb.ToString().Trim()
                Return If(String.IsNullOrWhiteSpace(result), "Error: No data found in .xlsx.", result)

            Catch ex As IOException
                If Not silent AndAlso IsWorkbookOpenInExcel(xlsxPath) Then
                    If EnsureClosedWorkbookForSandboxedRead(xlsxPath, silent) Then
                        Return ReadXlsxSandboxed(xlsxPath, silent, askWorksheetSelection)
                    End If

                    Return "Error: The workbook is open in Excel."
                End If

                Return $"Error reading .xlsx: {ex.Message}"

            Catch ex As Exception
                Return $"Error reading .xlsx: {ex.Message}"

            Finally
                Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                Catch : End Try
            End Try
        End Function


        ''' <summary>
        ''' Returns <c>True</c> when the workbook cannot currently be opened with exclusive read access,
        ''' which usually means that Excel still has the file open.
        ''' </summary>
        Private Shared Function IsWorkbookOpenInExcel(xlsxPath As String) As Boolean
            Try
                Using fs As New FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.None)
                End Using

                Return False

            Catch ex As IOException
                Return True
            End Try
        End Function

        ''' <summary>
        ''' In interactive mode, prompts the user to close an open workbook and retry.
        ''' </summary>
        Private Shared Function EnsureClosedWorkbookForSandboxedRead(xlsxPath As String,
                                                                     silent As Boolean) As Boolean
            If silent Then
                Return True
            End If

            Do While IsWorkbookOpenInExcel(xlsxPath)
                Dim answer As Integer = ShowCustomYesNoBox(
                    "The Excel workbook '" & Path.GetFileName(xlsxPath) & "' appears to be open. " &
                    "Please close it, then click Retry to try reading it again.",
                    "Retry",
                    "Cancel",
                    AN & " - Workbook Open",
                    nonModal:=True)

                If answer <> 1 Then
                    Return False
                End If

                System.Threading.Thread.Sleep(250)
            Loop

            Return True
        End Function


        ''' <summary>
        ''' Resolves the full path to a sheet XML file from a relationship ID or positional fallback.
        ''' </summary>
        Private Shared Function ResolveSheetPath(tempDir As String, rId As String,
                                                  ridMap As Dictionary(Of String, String),
                                                  positionalIndex As Integer) As String
            If Not String.IsNullOrWhiteSpace(rId) AndAlso ridMap.ContainsKey(rId) Then
                Dim rel = ridMap(rId)
                Dim resolved = Path.Combine(tempDir, "xl", rel.Replace("/"c, Path.DirectorySeparatorChar))
                If File.Exists(resolved) Then Return resolved
            End If
            ' Positional fallback
            Dim fallback = Path.Combine(tempDir, "xl", "worksheets", $"sheet{positionalIndex}.xml")
            Return If(File.Exists(fallback), fallback, Nothing)
        End Function


        ' ═══════════════════════════════════════════════════════════════════════
        '  PPTX — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        Private Const SB_PptxNs As String = "http://schemas.openxmlformats.org/presentationml/2006/main"
        Private Const SB_DrawNs As String = "http://schemas.openxmlformats.org/drawingml/2006/main"
        Private Const SB_PptxRelNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ''' <summary>
        ''' Extracts plain text from a .pptx file without COM interop.
        ''' Output format matches <c>ExtractPowerPointText</c>: <c>=== Slide {n} ===</c> + shape text +
        ''' <c>--- Notes ---</c> per slide.
        ''' </summary>
        ''' <param name="pptxPath">Absolute path to the .pptx file.</param>
        ''' <returns>Extracted text content, or an error string on failure.</returns>
        Public Shared Function ReadPptxSandboxed(pptxPath As String) As String
            If String.IsNullOrWhiteSpace(pptxPath) OrElse Not File.Exists(pptxPath) Then
                Return "Error: File not found."
            End If

            Dim tempDir As String = Path.Combine(Path.GetTempPath(), "ri_pptx_" & Guid.NewGuid().ToString("N"))
            Try
                ZipFile.ExtractToDirectory(pptxPath, tempDir)

                Dim presPath = Path.Combine(tempDir, "ppt", "presentation.xml")
                If Not File.Exists(presPath) Then Return "Error: Not a valid .pptx file (missing ppt/presentation.xml)."

                Dim presDoc As New XmlDocument()
                presDoc.Load(presPath)
                Dim presNs As New XmlNamespaceManager(presDoc.NameTable)
                presNs.AddNamespace("p", SB_PptxNs)
                presNs.AddNamespace("r", SB_PptxRelNs)

                ' Map rId → file path
                Dim relsPath = Path.Combine(tempDir, "ppt", "_rels", "presentation.xml.rels")
                Dim ridMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                If File.Exists(relsPath) Then
                    Dim relsDoc As New XmlDocument()
                    relsDoc.Load(relsPath)
                    If relsDoc.DocumentElement IsNot Nothing Then
                        For Each relNode As XmlNode In relsDoc.DocumentElement.ChildNodes
                            If relNode.Attributes Is Nothing Then Continue For
                            Dim id = relNode.Attributes("Id")
                            Dim target = relNode.Attributes("Target")
                            If id IsNot Nothing AndAlso target IsNot Nothing Then
                                ridMap(id.Value) = target.Value
                            End If
                        Next
                    End If
                End If

                Dim slideIdNodes = presDoc.SelectNodes("//p:sldIdLst/p:sldId", presNs)
                Dim sb As New StringBuilder(2048)
                Dim slideIndex As Integer = 0

                If slideIdNodes IsNot Nothing Then
                    For Each slideIdNode As XmlNode In slideIdNodes
                        slideIndex += 1
                        Dim rId = slideIdNode.Attributes("r:id")?.Value

                        Dim slidePath = ResolveSlidePath(tempDir, rId, ridMap, slideIndex)
                        If slidePath Is Nothing OrElse Not File.Exists(slidePath) Then Continue For

                        sb.AppendLine("=== Slide " & slideIndex.ToString(Globalization.CultureInfo.InvariantCulture) & " ===")

                        ' Extract shape text via DrawingML <a:t> nodes, grouped by <a:p>
                        ExtractDrawingText(slidePath, sb)

                        ' Extract notes
                        ExtractSlideNotes(slidePath, sb)

                        sb.AppendLine()
                    Next
                End If

                Dim result = sb.ToString().Trim()
                Return If(String.IsNullOrWhiteSpace(result), "Error: No text content found in .pptx.", result)

            Catch ex As Exception
                Return $"Error reading .pptx: {ex.Message}"
            Finally
                Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                Catch : End Try
            End Try
        End Function

        Private Shared Function ResolveSlidePath(tempDir As String, rId As String,
                                                  ridMap As Dictionary(Of String, String),
                                                  positionalIndex As Integer) As String
            If Not String.IsNullOrWhiteSpace(rId) AndAlso ridMap.ContainsKey(rId) Then
                Dim rel = ridMap(rId)
                Dim resolved = Path.Combine(tempDir, "ppt", rel.Replace("/"c, Path.DirectorySeparatorChar))
                If File.Exists(resolved) Then Return resolved
            End If
            Dim fallback = Path.Combine(tempDir, "ppt", "slides", $"slide{positionalIndex}.xml")
            Return If(File.Exists(fallback), fallback, Nothing)
        End Function

        ''' <summary>
        ''' Extracts text from DrawingML shapes in a slide XML, one line per text shape.
        ''' Matches <c>ExtractPowerPointText</c> output: each shape's text on its own line.
        ''' </summary>
        Private Shared Sub ExtractDrawingText(slideXmlPath As String, sb As StringBuilder)
            Dim slideDoc As New XmlDocument()
            slideDoc.Load(slideXmlPath)
            Dim slideNs As New XmlNamespaceManager(slideDoc.NameTable)
            slideNs.AddNamespace("a", SB_DrawNs)
            slideNs.AddNamespace("p", SB_PptxNs)

            ' Group by <p:sp> or <p:txBody> shapes — each shape gets one output line
            Dim spNodes = slideDoc.SelectNodes("//p:sp", slideNs)
            If spNodes IsNot Nothing Then
                For Each sp As XmlNode In spNodes
                    Dim txBody = sp.SelectSingleNode(".//p:txBody", slideNs)
                    If txBody Is Nothing Then Continue For

                    Dim shapeSb As New StringBuilder()
                    Dim paragraphs = txBody.SelectNodes("a:p", slideNs)
                    If paragraphs IsNot Nothing Then
                        Dim first As Boolean = True
                        For Each para As XmlNode In paragraphs
                            Dim tNodes = para.SelectNodes(".//a:t", slideNs)
                            If tNodes Is Nothing OrElse tNodes.Count = 0 Then Continue For
                            Dim paraText As New StringBuilder()
                            For Each tNode As XmlNode In tNodes
                                paraText.Append(tNode.InnerText)
                            Next
                            Dim text = paraText.ToString()
                            If Not String.IsNullOrWhiteSpace(text) Then
                                If Not first Then shapeSb.AppendLine()
                                shapeSb.Append(text)
                                first = False
                            End If
                        Next
                    End If

                    Dim shapeText = shapeSb.ToString().Trim()
                    If Not String.IsNullOrWhiteSpace(shapeText) Then
                        sb.AppendLine(shapeText)
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' Extracts notes from a slide's associated notesSlide, if present.
        ''' Matches <c>ExtractPowerPointText</c> output: <c>--- Notes ---</c> header.
        ''' </summary>
        Private Shared Sub ExtractSlideNotes(slideXmlPath As String, sb As StringBuilder)
            Dim notesRelsPath = Path.Combine(Path.GetDirectoryName(slideXmlPath), "_rels",
                                             Path.GetFileName(slideXmlPath) & ".rels")
            If Not File.Exists(notesRelsPath) Then Return

            Try
                Dim relsDoc As New XmlDocument()
                relsDoc.Load(notesRelsPath)
                If relsDoc.DocumentElement Is Nothing Then Return

                For Each rel As XmlNode In relsDoc.DocumentElement.ChildNodes
                    If rel.Attributes Is Nothing Then Continue For
                    Dim relType = rel.Attributes("Type")?.Value
                    If relType Is Nothing OrElse Not relType.Contains("notesSlide") Then Continue For

                    Dim notesTarget = rel.Attributes("Target")?.Value
                    If String.IsNullOrWhiteSpace(notesTarget) Then Continue For

                    Dim notesPath = Path.Combine(Path.GetDirectoryName(slideXmlPath),
                                                 notesTarget.Replace("/"c, Path.DirectorySeparatorChar))
                    If Not File.Exists(notesPath) Then Continue For

                    Dim notesDoc As New XmlDocument()
                    notesDoc.Load(notesPath)
                    Dim notesNs As New XmlNamespaceManager(notesDoc.NameTable)
                    notesNs.AddNamespace("a", SB_DrawNs)

                    Dim noteTexts = notesDoc.SelectNodes("//a:t", notesNs)
                    If noteTexts Is Nothing OrElse noteTexts.Count = 0 Then Continue For

                    Dim hasNotes = False
                    For Each nt As XmlNode In noteTexts
                        If Not String.IsNullOrWhiteSpace(nt.InnerText) Then
                            If Not hasNotes Then
                                sb.AppendLine("--- Notes ---")
                                hasNotes = True
                            End If
                            sb.AppendLine(nt.InnerText.Trim())
                        End If
                    Next
                Next
            Catch
            End Try
        End Sub


        ' ═══════════════════════════════════════════════════════════════════════
        '  EML — Sandboxed
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Extracts plain text from a .eml (RFC 2822) file by parsing headers and body.
        ''' Output matches <c>ParseEmlAsText</c> format in AutoPilot.
        ''' </summary>
        Public Shared Function ReadEmlSandboxed(emlPath As String) As String
            If String.IsNullOrWhiteSpace(emlPath) OrElse Not File.Exists(emlPath) Then
                Return "Error: File not found."
            End If

            Try
                Dim emlContent = File.ReadAllText(emlPath, Encoding.UTF8)
                If String.IsNullOrWhiteSpace(emlContent) Then Return "Error: Empty .eml file."

                Dim sb As New StringBuilder()
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine($"EMAIL MESSAGE (from {Path.GetFileName(emlPath)})")
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine()

                Dim headerEnd = emlContent.IndexOf(vbCrLf & vbCrLf, StringComparison.Ordinal)
                If headerEnd < 0 Then headerEnd = emlContent.IndexOf(vbLf & vbLf, StringComparison.Ordinal)

                Dim headerSection As String
                Dim bodySection As String

                If headerEnd >= 0 Then
                    headerSection = emlContent.Substring(0, headerEnd)
                    bodySection = emlContent.Substring(headerEnd).TrimStart({CChar(vbCr), CChar(vbLf)})
                Else
                    headerSection = emlContent
                    bodySection = ""
                End If

                ' Extract key headers
                For Each headerName In {"From", "To", "CC", "Subject", "Date"}
                    Dim pattern = $"(?m)^{headerName}:\s*(.+?)(?=\r?\n\S|\r?\n\r?\n|$)"
                    Dim m = Text.RegularExpressions.Regex.Match(headerSection, pattern,
                        Text.RegularExpressions.RegexOptions.IgnoreCase Or
                        Text.RegularExpressions.RegexOptions.Singleline)
                    If m.Success Then
                        sb.AppendLine($"{headerName}: {m.Groups(1).Value.Trim()}")
                    End If
                Next

                sb.AppendLine()
                sb.AppendLine("───────────────────────────────────────────────────")
                sb.AppendLine()

                ' Body extraction with MIME boundary handling
                If Not String.IsNullOrWhiteSpace(bodySection) Then
                    Dim boundaryMatch = Text.RegularExpressions.Regex.Match(headerSection,
                        "boundary=""?([^"";\s]+)""?", Text.RegularExpressions.RegexOptions.IgnoreCase)
                    If boundaryMatch.Success Then
                        Dim boundary = boundaryMatch.Groups(1).Value
                        Dim parts = bodySection.Split({$"--{boundary}"}, StringSplitOptions.RemoveEmptyEntries)
                        Dim textPartFound = False
                        For Each part In parts
                            If part.StartsWith("--") Then Continue For
                            If part.IndexOf("Content-Type: text/plain", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                Dim partHeaderEnd = part.IndexOf(vbCrLf & vbCrLf, StringComparison.Ordinal)
                                If partHeaderEnd < 0 Then partHeaderEnd = part.IndexOf(vbLf & vbLf, StringComparison.Ordinal)
                                If partHeaderEnd >= 0 Then
                                    Dim textBody = part.Substring(partHeaderEnd).Trim()
                                    If textBody.Length > 50000 Then textBody = textBody.Substring(0, 50000) & vbCrLf & "[... truncated ...]"
                                    sb.Append(textBody)
                                    textPartFound = True
                                    Exit For
                                End If
                            End If
                        Next
                        If Not textPartFound Then
                            If bodySection.Length > 50000 Then bodySection = bodySection.Substring(0, 50000) & vbCrLf & "[... truncated ...]"
                            sb.Append(bodySection)
                        End If
                    Else
                        If bodySection.Length > 50000 Then bodySection = bodySection.Substring(0, 50000) & vbCrLf & "[... truncated ...]"
                        sb.Append(bodySection)
                    End If
                End If

                Return sb.ToString().Trim()

            Catch ex As Exception
                Return $"Error reading .eml: {ex.Message}"
            End Try
        End Function


        ' ═══════════════════════════════════════════════════════════════════════
        '  MSG — Sandboxed (with callback for MAPI and nested file support)
        ' ═══════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Delegate for extracting text and nested file paths from a .msg file via Outlook COM.
        ''' Returns the mail body text; populates <paramref name="nestedFiles"/> with paths
        ''' to any attachments saved to the temp directory for recursive processing.
        ''' </summary>
        ''' <param name="msgPath">Path to the .msg file.</param>
        ''' <param name="tempDir">Temp directory for saving nested attachments.</param>
        ''' <param name="nestedFiles">Output: list of saved nested attachment paths.</param>
        ''' <returns>Extracted mail text, or Nothing on failure.</returns>
        Public Delegate Function MsgReadCallback(msgPath As String, tempDir As String,
                                                  ByRef nestedFiles As List(Of String)) As String

        ''' <summary>
        ''' Extracts plain text from a .msg file.
        ''' When <paramref name="msgReadFunc"/> is provided, uses Outlook COM to extract body and nested attachments.
        ''' Nested attachments (docx, xlsx, pptx, pdf, eml, msg, text files) are recursively read and appended.
        ''' Falls back to heuristic UTF-16LE extraction when no callback is available.
        ''' </summary>
        ''' <param name="msgPath">Absolute path to the .msg file.</param>
        ''' <param name="msgReadFunc">Optional callback for MAPI-based extraction with nested file support.</param>
        ''' <param name="depth">Current recursion depth (for nested msg/eml within msg). Max 5.</param>
        ''' <returns>Extracted mail text with inline nested attachment content, or an error string on failure.</returns>
        Public Shared Function ReadMsgSandboxed(msgPath As String,
                                                Optional msgReadFunc As MsgReadCallback = Nothing,
                                                Optional depth As Integer = 0) As String
            If String.IsNullOrWhiteSpace(msgPath) OrElse Not File.Exists(msgPath) Then
                Return "Error: File not found."
            End If

            If depth > 5 Then
                Return $"[Skipped: max nesting depth reached for {Path.GetFileName(msgPath)}]"
            End If

            Dim sb As New StringBuilder()

            ' Prefer the Outlook-based callback when available
            If msgReadFunc IsNot Nothing Then
                Dim tempDir = Path.Combine(Path.GetTempPath(), "ri_msg_" & Guid.NewGuid().ToString("N"))
                Try
                    Directory.CreateDirectory(tempDir)
                    Dim nestedFiles As New List(Of String)()
                    Dim bodyText = msgReadFunc(msgPath, tempDir, nestedFiles)

                    If Not String.IsNullOrWhiteSpace(bodyText) Then
                        sb.Append(bodyText)

                        ' Recursively read nested attachments
                        If nestedFiles IsNot Nothing AndAlso nestedFiles.Count > 0 Then
                            For Each nestedPath In nestedFiles
                                If Not File.Exists(nestedPath) Then Continue For
                                Dim nestedExt = Path.GetExtension(nestedPath).ToLowerInvariant()
                                Dim nestedText As String = Nothing

                                Try
                                    Select Case nestedExt
                                        Case ".docx"
                                            nestedText = ReadDocxSandboxed(nestedPath)
                                        Case ".xlsx"
                                            nestedText = ReadXlsxSandboxed(nestedPath)
                                        Case ".pptx"
                                            nestedText = ReadPptxSandboxed(nestedPath)
                                        Case ".eml"
                                            nestedText = ReadEmlSandboxed(nestedPath)
                                        Case ".msg"
                                            nestedText = ReadMsgSandboxed(nestedPath, msgReadFunc, depth + 1)
                                        Case ".pdf"
                                            Try
                                                nestedText = ReadPdfAsText(nestedPath, True, False, False, Nothing).Result
                                            Catch
                                            End Try
                                        Case ".txt", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                                             ".md", ".yaml", ".yml", ".ini"
                                            nestedText = ReadTextFile(nestedPath, False)
                                        Case ".doc"
                                            If SharedMethods.INI_AllowLegacyDocFiles_Cached Then
                                                nestedText = ReadWordDocument(nestedPath, False)
                                            Else
                                                nestedText = "[Skipped: .doc format disabled for security]"
                                            End If
                                    End Select
                                Catch ex As Exception
                                    nestedText = $"[Error reading {Path.GetFileName(nestedPath)}: {ex.Message}]"
                                End Try

                                If Not String.IsNullOrWhiteSpace(nestedText) AndAlso
                                   Not nestedText.StartsWith("Error") Then
                                    sb.AppendLine()
                                    sb.AppendLine()
                                    sb.AppendLine($"═══ Attachment: {Path.GetFileName(nestedPath)} ═══")
                                    sb.AppendLine()
                                    sb.Append(nestedText)
                                End If
                            Next
                        End If

                        Return sb.ToString().Trim()
                    End If
                Catch
                    ' Fall through to heuristic extraction
                Finally
                    Try : If Directory.Exists(tempDir) Then Directory.Delete(tempDir, True)
                    Catch : End Try
                End Try
            End If

            ' Heuristic fallback: scan raw bytes for UTF-16LE text runs
            Try
                Dim rawBytes = File.ReadAllBytes(msgPath)
                sb.Clear()
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine($"EMAIL MESSAGE (from {Path.GetFileName(msgPath)})")
                sb.AppendLine("═══════════════════════════════════════════════════")
                sb.AppendLine()

                Dim unicodeText = Encoding.Unicode.GetString(rawBytes)
                Dim runs As New List(Of String)()
                Dim currentRun As New StringBuilder()
                For Each ch In unicodeText
                    If Char.IsLetterOrDigit(ch) OrElse Char.IsPunctuation(ch) OrElse
                       Char.IsWhiteSpace(ch) OrElse Char.IsSymbol(ch) Then
                        currentRun.Append(ch)
                    Else
                        If currentRun.Length >= 20 Then
                            runs.Add(currentRun.ToString().Trim())
                        End If
                        currentRun.Clear()
                    End If
                Next
                If currentRun.Length >= 20 Then runs.Add(currentRun.ToString().Trim())

                If runs.Count > 0 Then
                    Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
                    For Each run In runs
                        If seen.Add(run) Then sb.AppendLine(run)
                    Next
                Else
                    sb.AppendLine("[No readable text could be extracted from .msg without Outlook]")
                End If

                Return sb.ToString().Trim()

            Catch ex As Exception
                Return $"Error reading .msg: {ex.Message}"
            End Try
        End Function

    End Class
End Namespace