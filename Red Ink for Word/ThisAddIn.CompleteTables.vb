' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.CompleteTables.vb
' Purpose: Completes Word document tables using AI by extracting table structure
'          to JSON, having the LLM fill/expand them, then patching OpenXML.
'
' Architecture / Key Ideas:
'  - OpenXML Processing: Operates directly on DOCX XML, modifying only <w:t>
'    nodes within <w:tc> cells and cloning <w:tr> rows for insertions.
'  - JSON Round-Trip: Extracts table structure (headings, row labels, cell text,
'    merged cell spans, column widths) into compact JSON, sends to LLM, parses
'    response JSON, and applies changes back to XML preserving all formatting.
'  - Form Field Support: Detects and exposes structured document tags (SDTs)
'    including dropdown/comboBox selections and checkbox states. The LLM can
'    select dropdown items by value and toggle checkboxes. Content control
'    plain-text fields are also extracted and modifiable.
'    Legacy form fields (w:ffData) with dropdown lists and checkboxes are
'    also extracted and can be set by the LLM.
'  - Full Document Context: The LLM receives the entire document body text
'    (paragraphs between and around tables) so it can understand document
'    purpose, fill tables consistently, and detect placeholders anywhere.
'  - Placeholder Detection: Cells containing patterns like [...], ___, TBD,
'    <enter ...>, N/A, or similar are flagged with "ph":true so the LLM
'    knows they require completion.
'  - Column Width Hints: Relative column widths (as percentages) are included
'    so the LLM can gauge how much text fits in each cell.
'  - Row Cloning: New rows are created by deep-cloning the last data row of the
'    table, preserving cell properties (w:tcPr), paragraph properties (w:pPr),
'    run properties (w:rPr), borders, widths, and shading.
'  - Batch Processing: Large documents with many tables are processed in batches
'    to stay within LLM token limits.
'  - Formatting Preservation: Only text content changes; all styles, borders,
'    widths, merged cell definitions, and document structure are untouched.
'  - Prompt Persistence: The user's last instructions are saved/restored via
'    My.Settings.TableComplete_LastInstructions.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Maximum characters of JSON per LLM batch for table completion.
    ''' </summary>
    Private Const TableCompleteMaxCharsPerBatch As Integer = 12000

    ''' <summary>
    ''' Maximum tables per LLM batch.
    ''' </summary>
    Private Const TableCompleteMaxTablesPerBatch As Integer = 3

    ''' <summary>
    ''' Maximum characters of document context to include per batch.
    ''' </summary>
    Private Const TableCompleteMaxContextChars As Integer = 4000

    ''' <summary>
    ''' XML namespace URI for w14 extensions (checkbox content controls).
    ''' </summary>
    Private Const W14Ns As String = "http://schemas.microsoft.com/office/word/2010/wordml"

    ''' <summary>
    ''' XML namespace URI for the main wordprocessingml namespace.
    ''' </summary>
    Private Const WNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    ''' <summary>
    ''' Regex pattern matching common placeholder content in table cells.
    ''' Matches: [...], [TBD], ___, …, TBD, N/A, n/a, &lt;enter ...&gt;, {placeholder}, etc.
    ''' </summary>
    Private Shared ReadOnly PlaceholderPattern As New Regex(
        "^\s*(\[.*\]|\{.*\}|<[^>]*>|_{3,}|\.{3,}|…+|TBD|TBC|N/?A|TODO|FIXME|XXX|INSERT|ENTER)\s*$",
        RegexOptions.IgnoreCase Or RegexOptions.Compiled)

#Region "Data Classes"

    ''' <summary>
    ''' Represents a single cell extracted from a table for JSON serialization.
    ''' </summary>
    Private Class TableCellInfo
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property Text As String
        Public Property GridSpan As Integer
        Public Property VMerge As String
        Public Property TextNodes As List(Of System.Xml.XmlNode)
        Public Property IsHeader As Boolean
        Public Property IsPlaceholder As Boolean
        Public Property WidthPct As Double
        ''' <summary>Form fields (SDT and legacy) found inside this cell.</summary>
        Public Property FormFields As List(Of FormFieldInfo)
    End Class

    ''' <summary>
    ''' Represents a table row extracted from the document.
    ''' </summary>
    Private Class TableRowInfo
        Public Property RowIndex As Integer
        Public Property Cells As List(Of TableCellInfo)
        Public Property XmlNode As System.Xml.XmlNode
    End Class

    ''' <summary>
    ''' Represents a complete table extracted from the document.
    ''' </summary>
    Private Class TableInfo
        Public Property TableIndex As Integer
        Public Property Rows As List(Of TableRowInfo)
        Public Property XmlNode As System.Xml.XmlNode
        Public Property ColumnCount As Integer
        Public Property ContextBefore As String
        Public Property ColumnWidths As List(Of Double)
    End Class

    ''' <summary>
    ''' Represents a cell update or new row from the LLM response.
    ''' </summary>
    Private Class TableCellUpdate
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property Text As String
    End Class

    ''' <summary>
    ''' Represents the LLM's response for a single table.
    ''' </summary>
    Private Class TableUpdateResponse
        Public Property TableIndex As Integer
        Public Property Updates As List(Of TableCellUpdate)
        Public Property NewRows As List(Of List(Of String))
        Public Property InsertAfterRow As Integer
        Public Property FieldUpdates As List(Of FormFieldUpdate)
    End Class

    ''' <summary>
    ''' The type of a form field found in the document.
    ''' </summary>
    Private Enum FormFieldType
        ''' <summary>SDT dropdown or comboBox content control.</summary>
        Dropdown
        ''' <summary>SDT checkbox content control (w14:checkbox).</summary>
        Checkbox
        ''' <summary>SDT plain-text or rich-text content control.</summary>
        TextInput
        ''' <summary>Legacy form field dropdown (w:ffData/w:ddList).</summary>
        LegacyDropdown
        ''' <summary>Legacy form field checkbox (w:ffData/w:checkBox).</summary>
        LegacyCheckbox
    End Enum

    ''' <summary>
    ''' Extracted information about a form field (SDT or legacy).
    ''' </summary>
    Private Class FormFieldInfo
        ''' <summary>Unique ID within the cell for LLM round-trip ("f0", "f1", ...).</summary>
        Public Property FieldId As String
        Public Property FieldType As FormFieldType
        ''' <summary>Tag or alias of the content control (w:sdtPr/w:tag or w:sdtPr/w:alias), or legacy field name.</summary>
        Public Property Name As String
        ''' <summary>Current value: selected item text for dropdowns, "true"/"false" for checkboxes, text for text inputs.</summary>
        Public Property CurrentValue As String
        ''' <summary>Available options for dropdown fields (display values).</summary>
        Public Property Options As List(Of String)
        ''' <summary>The w:sdt node (for SDTs) or the w:fldChar node (for legacy fields).</summary>
        Public Property XmlNode As System.Xml.XmlNode
    End Class

    ''' <summary>
    ''' An LLM-requested change to a form field.
    ''' </summary>
    Private Class FormFieldUpdate
        Public Property Row As Integer
        Public Property Col As Integer
        Public Property FieldId As String
        Public Property Value As String
    End Class

    ''' <summary>
    ''' Represents a section of body-level content (paragraphs between tables) that
    ''' may contain placeholders, empty text, or form fields that need completion.
    ''' </summary>
    Private Class BodySectionInfo
        ''' <summary>Sequential index among all body sections.</summary>
        Public Property SectionIndex As Integer
        ''' <summary>Heading or context text preceding the section paragraphs.</summary>
        Public Property HeadingText As String
        ''' <summary>Paragraphs in this section that are candidates for completion.</summary>
        Public Property Paragraphs As List(Of BodyParagraphInfo)
        ''' <summary>Body-level SDT form fields in this section.</summary>
        Public Property FormFields As List(Of FormFieldInfo)
    End Class

    ''' <summary>
    ''' A single body-level paragraph that may be empty or contain placeholder text.
    ''' </summary>
    Private Class BodyParagraphInfo
        Public Property Index As Integer
        Public Property Text As String
        Public Property IsPlaceholder As Boolean
        Public Property IsEmpty As Boolean
        Public Property TextNodes As List(Of System.Xml.XmlNode)
        Public Property XmlNode As System.Xml.XmlNode
    End Class

    ''' <summary>
    ''' LLM response for body section updates.
    ''' </summary>
    Private Class BodySectionUpdateResponse
        Public Property SectionIndex As Integer
        Public Property ParagraphUpdates As List(Of BodyParagraphUpdate)
        Public Property FieldUpdates As List(Of BodyFieldUpdate)
    End Class

    Private Class BodyParagraphUpdate
        Public Property Index As Integer
        Public Property Text As String
    End Class

    Private Class BodyFieldUpdate
        Public Property FieldId As String
        Public Property Value As String
    End Class

#End Region

#Region "Entry Point and Pipeline"

    ''' <summary>
    ''' Entry point: prompts for file and user instructions, then completes tables.
    ''' </summary>
    Public Async Sub CompleteWordDocumentTables()
        Dim selectedPath As String = ""

        If INI_AllowLegacyDocFiles Then
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents|*.doc;*.docx|Word Document (*.docx)|*.docx|Word 97-2003 (*.doc)|*.doc"
        Else
            Globals.ThisAddIn.DragDropFormLabel = "Select a Word document with tables to complete"
            Globals.ThisAddIn.DragDropFormFilter = "Word Documents (*.docx)|*.docx"
        End If

        Try
            Using frm As New DragDropForm(DragDropMode.FileOrDirectory)
                If frm.ShowDialog() = DialogResult.OK Then
                    selectedPath = frm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        If String.IsNullOrWhiteSpace(selectedPath) Then Exit Sub
        If Not File.Exists(selectedPath) Then
            ShowCustomMessageBox("The selected file does not exist.")
            Exit Sub
        End If

        Dim ext As String = Path.GetExtension(selectedPath).ToLowerInvariant()
        If ext = ".doc" AndAlso Not INI_AllowLegacyDocFiles Then
            ShowCustomMessageBox("The .doc format is disabled for security. Please convert to .docx first.")
            Exit Sub
        End If
        If ext <> ".doc" AndAlso ext <> ".docx" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Please select a .docx file.")
            Exit Sub
        End If

        ' Restore last instructions from My.Settings
        Dim lastInstructions As String = ""
        Try
            lastInstructions = If(My.Settings.TableComplete_LastInstructions, "")
        Catch
        End Try

        ' Get user instructions for table completion
        Dim userInstructions As String = ShowCustomInputBox(
            "Describe how the tables should be completed:" & vbCrLf &
            "(e.g., 'Fill in the empty cells based on headings and row labels. " &
            "Add rows where the table appears incomplete.')",
            AN & " Complete Tables", False, lastInstructions)

        If String.IsNullOrWhiteSpace(userInstructions) Then Exit Sub
        userInstructions = userInstructions.Trim()

        ' Save instructions to My.Settings
        Try
            My.Settings.TableComplete_LastInstructions = userInstructions
            My.Settings.Save()
        Catch
        End Try

        ' Determine output path
        Dim dir As String = Path.GetDirectoryName(selectedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputPath As String = Path.Combine(dir, $"{nameWithoutExt}_completed.docx")

        If File.Exists(outputPath) Then
            Dim overwrite As Integer = ShowCustomYesNoBox(
                $"Output file already exists:{vbCrLf}{Path.GetFileName(outputPath)}{vbCrLf}{vbCrLf}Overwrite?",
                "Yes, overwrite", "No, cancel")
            If overwrite <> 1 Then Exit Sub
        End If

        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = 1
        ProgressBarModule.GlobalProgressLabel = "Analyzing tables..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread(AN & " Complete Tables", "Starting...")

        Try
            Dim success As Boolean = Await ProcessTableCompletion(selectedPath, outputPath, userInstructions)

            If success Then
                ShowCustomMessageBox($"Tables completed successfully.{vbCrLf}Output: {Path.GetFileName(outputPath)}", AN & " Complete Tables")
            Else
                ShowCustomMessageBox("Table completion failed or was cancelled.", AN & " Complete Tables")
            End If
        Catch ex As Exception
            ShowCustomMessageBox($"Error: {ex.Message}", AN & " Complete Tables")
        Finally
            ProgressBarModule.CancelOperation = True
        End Try
    End Sub

    Private Async Function ProcessTableCompletion(inputPath As String, outputPath As String, userInstructions As String) As Task(Of Boolean)
        Dim tempDocxPath As String = Nothing
        Dim wordApp As Word.Application = Nothing
        Dim doc As Word.Document = Nothing

        Try
            If Path.GetExtension(inputPath).ToLowerInvariant() = ".doc" Then
                tempDocxPath = Path.Combine(Path.GetTempPath(), $"{AN2}_tblconv_{Guid.NewGuid():N}.docx")
                wordApp = Globals.ThisAddIn.Application
                wordApp.ScreenUpdating = False
                doc = wordApp.Documents.Open(inputPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                doc.SaveAs2(tempDocxPath, WdSaveFormat.wdFormatXMLDocument)
                doc.Close(WdSaveOptions.wdDoNotSaveChanges)
                doc = Nothing
                wordApp.ScreenUpdating = True
            Else
                tempDocxPath = inputPath
            End If

            File.Copy(tempDocxPath, outputPath, overwrite:=True)
            Dim success As Boolean = Await ProcessDocxTables(outputPath, userInstructions)
            Return success

        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            End If
            If tempDocxPath IsNot Nothing AndAlso tempDocxPath <> inputPath AndAlso File.Exists(tempDocxPath) Then
                Try : File.Delete(tempDocxPath) : Catch : End Try
            End If
        End Try
    End Function

    Private Async Function ProcessDocxTables(docxPath As String, userInstructions As String) As Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), $"{AN2}_tbl_{Guid.NewGuid():N}")

        Try
            ZipFile.ExtractToDirectory(docxPath, tempDir)

            Dim documentXmlPath As String = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then
                ShowCustomMessageBox("Invalid DOCX structure - document.xml not found.")
                Return False
            End If

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", WNs)
            nsMgr.AddNamespace("w14", W14Ns)

            ' Phase 1: Extract full document context + all tables
            Dim documentContext As String = ExtractDocumentBodyContext(xmlDoc, nsMgr)
            Dim tables As List(Of TableInfo) = ExtractTablesFromXml(xmlDoc, nsMgr)

            If tables.Count = 0 Then
                ShowCustomMessageBox("No tables found in the document.")
                Return False
            End If

            Dim placeholderCount As Integer = 0
            Dim formFieldCount As Integer = 0
            For Each tbl In tables
                For Each row In tbl.Rows
                    For Each cell In row.Cells
                        If cell.IsPlaceholder Then placeholderCount += 1
                        If cell.FormFields IsNot Nothing Then formFieldCount += cell.FormFields.Count
                    Next
                Next
            Next

            Dim confirmMsg As String = $"Found {tables.Count} table(s)"
            If placeholderCount > 0 Then confirmMsg &= $", {placeholderCount} placeholder(s)"
            If formFieldCount > 0 Then confirmMsg &= $", {formFieldCount} form field(s)"
            confirmMsg &= ". Continue with AI completion?"
            Dim confirm As Integer = ShowCustomYesNoBox(confirmMsg, "Yes, continue", "No, cancel")
            If confirm <> 1 Then Return False

            ' Phase 2: Send to LLM in batches
            Dim allUpdates As List(Of TableUpdateResponse) = Await ProcessTableBatches(tables, userInstructions, documentContext)
            If allUpdates Is Nothing Then Return False

            ' Phase 3: Apply changes back to XML
            ApplyTableUpdates(tables, allUpdates, nsMgr)

            ' Save and repack
            xmlDoc.Save(documentXmlPath)
            File.Delete(docxPath)
            ZipFile.CreateFromDirectory(tempDir, docxPath, CompressionLevel.Optimal, False)

            Return True

        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

#End Region

#Region "Extraction"

    Private Function ExtractDocumentBodyContext(xmlDoc As System.Xml.XmlDocument, nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sb As New StringBuilder()
        Dim bodyNode As System.Xml.XmlNode = xmlDoc.SelectSingleNode("//w:body", nsMgr)
        If bodyNode Is Nothing Then Return ""

        Dim tableCounter As Integer = 0

        For Each child As System.Xml.XmlNode In bodyNode.ChildNodes
            If child.NodeType <> System.Xml.XmlNodeType.Element Then Continue For

            If child.LocalName = "p" Then
                Dim text As String = GetParagraphPlainText(child, nsMgr)
                If Not String.IsNullOrWhiteSpace(text) Then
                    sb.AppendLine(text)
                End If
            ElseIf child.LocalName = "tbl" Then
                sb.AppendLine($"[Table {tableCounter}]")
                tableCounter += 1
            End If

            If sb.Length > TableCompleteMaxContextChars * 2 Then Exit For
        Next

        Dim result As String = sb.ToString()
        If result.Length > TableCompleteMaxContextChars Then
            Dim halfLen As Integer = TableCompleteMaxContextChars \ 2
            result = result.Substring(0, halfLen) &
                     vbCrLf & "[... document content truncated ...]" & vbCrLf &
                     result.Substring(result.Length - halfLen)
        End If

        Return result
    End Function

    Private Function ExtractTablesFromXml(xmlDoc As System.Xml.XmlDocument, nsMgr As System.Xml.XmlNamespaceManager) As List(Of TableInfo)
        Dim tables As New List(Of TableInfo)()
        Dim tblNodes As System.Xml.XmlNodeList = xmlDoc.SelectNodes("//w:tbl", nsMgr)
        Dim tableIndex As Integer = 0

        For Each tblNode As System.Xml.XmlNode In tblNodes
            Dim tblInfo As New TableInfo() With {
                .TableIndex = tableIndex,
                .Rows = New List(Of TableRowInfo)(),
                .XmlNode = tblNode,
                .ColumnCount = 0,
                .ContextBefore = "",
                .ColumnWidths = New List(Of Double)()
            }

            ' Extract column widths from w:tblGrid/w:gridCol
            Dim gridCols As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tblGrid/w:gridCol", nsMgr)
            Dim rawWidths As New List(Of Double)()
            Dim totalWidth As Double = 0

            For Each gridCol As System.Xml.XmlNode In gridCols
                Dim wAttr As System.Xml.XmlNode = gridCol.Attributes("w:w")
                Dim colWidth As Double = 0
                If wAttr IsNot Nothing Then Double.TryParse(wAttr.Value, colWidth)
                rawWidths.Add(colWidth)
                totalWidth += colWidth
            Next

            If totalWidth > 0 AndAlso rawWidths.Count > 0 Then
                For Each w In rawWidths
                    tblInfo.ColumnWidths.Add(Math.Round(w / totalWidth * 100, 1))
                Next
            End If

            ' Get context: text from preceding sibling paragraph(s)
            Dim prevSibling As System.Xml.XmlNode = tblNode.PreviousSibling
            Dim contextParts As New List(Of String)()
            Dim contextCount As Integer = 0
            While prevSibling IsNot Nothing AndAlso contextCount < 3
                If prevSibling.NodeType = System.Xml.XmlNodeType.Element AndAlso prevSibling.LocalName = "p" Then
                    Dim paraText As String = GetParagraphPlainText(prevSibling, nsMgr)
                    If Not String.IsNullOrWhiteSpace(paraText) Then
                        contextParts.Insert(0, paraText)
                        contextCount += 1
                    End If
                End If
                prevSibling = prevSibling.PreviousSibling
            End While
            tblInfo.ContextBefore = String.Join(" | ", contextParts)

            ' Extract rows
            Dim trNodes As System.Xml.XmlNodeList = tblNode.SelectNodes("w:tr", nsMgr)
            Dim rowIndex As Integer = 0

            For Each trNode As System.Xml.XmlNode In trNodes
                Dim rowInfo As New TableRowInfo() With {
                    .RowIndex = rowIndex,
                    .Cells = New List(Of TableCellInfo)(),
                    .XmlNode = trNode
                }

                Dim isHeaderRow As Boolean = (trNode.SelectSingleNode("w:trPr/w:tblHeader", nsMgr) IsNot Nothing)
                Dim tcNodes As System.Xml.XmlNodeList = trNode.SelectNodes("w:tc", nsMgr)
                Dim colIndex As Integer = 0

                For Each tcNode As System.Xml.XmlNode In tcNodes
                    Dim cellInfo As New TableCellInfo() With {
                        .Row = rowIndex,
                        .Col = colIndex,
                        .TextNodes = New List(Of System.Xml.XmlNode)(),
                        .IsHeader = isHeaderRow OrElse rowIndex = 0,
                        .GridSpan = 1,
                        .VMerge = "",
                        .IsPlaceholder = False,
                        .WidthPct = 0,
                        .FormFields = New List(Of FormFieldInfo)()
                    }

                    ' gridSpan
                    Dim gridSpanNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:gridSpan/@w:val", nsMgr)
                    If gridSpanNode IsNot Nothing Then
                        Dim span As Integer
                        If Integer.TryParse(gridSpanNode.Value, span) Then cellInfo.GridSpan = span
                    End If

                    ' cell width
                    If tblInfo.ColumnWidths.Count > 0 Then
                        Dim widthSum As Double = 0
                        For gi As Integer = colIndex To Math.Min(colIndex + cellInfo.GridSpan - 1, tblInfo.ColumnWidths.Count - 1)
                            widthSum += tblInfo.ColumnWidths(gi)
                        Next
                        cellInfo.WidthPct = widthSum
                    End If

                    ' vMerge
                    Dim vMergeNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:tcPr/w:vMerge", nsMgr)
                    If vMergeNode IsNot Nothing Then
                        Dim vMergeVal As System.Xml.XmlNode = vMergeNode.Attributes("w:val")
                        cellInfo.VMerge = If(vMergeVal IsNot Nothing AndAlso vMergeVal.Value = "restart", "restart", "continue")
                    End If

                    ' ── Extract form fields FIRST so we know which SDT nodes exist ──
                    ExtractFormFieldsFromCell(tcNode, cellInfo, nsMgr)

                    ' ── Collect all w:sdt nodes owned by form fields ──
                    Dim sdtAncestors As New HashSet(Of System.Xml.XmlNode)()
                    If cellInfo.FormFields IsNot Nothing Then
                        For Each ff In cellInfo.FormFields
                            If ff.XmlNode IsNot Nothing AndAlso ff.XmlNode.LocalName = "sdt" Then
                                sdtAncestors.Add(ff.XmlNode)
                            End If
                        Next
                    End If

                    ' ── Extract text from w:t nodes, EXCLUDING those inside any SDT ──
                    Dim textBuilder As New StringBuilder()
                    Dim allTextNodes As System.Xml.XmlNodeList = tcNode.SelectNodes(".//w:t", nsMgr)
                    For Each tNode As System.Xml.XmlNode In allTextNodes
                        Dim insideSdt As Boolean = False
                        If sdtAncestors.Count > 0 Then
                            Dim ancestor As System.Xml.XmlNode = tNode.ParentNode
                            While ancestor IsNot Nothing AndAlso ancestor IsNot tcNode
                                If ancestor.LocalName = "sdt" AndAlso sdtAncestors.Contains(ancestor) Then
                                    insideSdt = True
                                    Exit While
                                End If
                                ancestor = ancestor.ParentNode
                            End While
                        End If

                        If Not insideSdt Then
                            cellInfo.TextNodes.Add(tNode)
                            textBuilder.Append(tNode.InnerText)
                        End If
                    Next
                    cellInfo.Text = textBuilder.ToString()

                    ' Detect placeholder content
                    If Not String.IsNullOrWhiteSpace(cellInfo.Text) Then
                        cellInfo.IsPlaceholder = PlaceholderPattern.IsMatch(cellInfo.Text)
                    End If

                    rowInfo.Cells.Add(cellInfo)
                    colIndex += cellInfo.GridSpan
                Next

                If colIndex > tblInfo.ColumnCount Then tblInfo.ColumnCount = colIndex
                tblInfo.Rows.Add(rowInfo)
                rowIndex += 1
            Next

            tables.Add(tblInfo)
            tableIndex += 1
        Next

        Return tables
    End Function

    ''' <summary>
    ''' Extracts SDT content controls and legacy form fields from a table cell node.
    ''' For SDT dropdowns/comboBoxes: if the current display text does not match any
    ''' defined listItem option, the field is treated as unselected (empty).
    ''' </summary>
    Private Sub ExtractFormFieldsFromCell(tcNode As System.Xml.XmlNode, cellInfo As TableCellInfo, nsMgr As System.Xml.XmlNamespaceManager)
        Dim fieldIndex As Integer = 0

        ' === SDT content controls (w:sdt) ===
        Dim sdtNodes As System.Xml.XmlNodeList = tcNode.SelectNodes(".//w:sdt", nsMgr)
        For Each sdtNode As System.Xml.XmlNode In sdtNodes
            Dim sdtPr As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr", nsMgr)
            If sdtPr Is Nothing Then Continue For

            Dim ff As New FormFieldInfo() With {
                .FieldId = $"f{fieldIndex}",
                .XmlNode = sdtNode,
                .Options = New List(Of String)(),
                .Name = "",
                .CurrentValue = ""
            }

            ' Get tag or alias as name
            Dim tagNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:tag/@w:val", nsMgr)
            Dim aliasNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:alias/@w:val", nsMgr)
            ff.Name = If(tagNode?.Value, If(aliasNode?.Value, ""))

            ' Detect whether this SDT is still showing its placeholder text
            Dim showingPlcHdr As Boolean = (sdtPr.SelectSingleNode("w:showingPlcHdr", nsMgr) IsNot Nothing)

            ' Determine type
            Dim dropDownNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
            Dim comboBoxNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)
            Dim checkboxNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w14:checkbox", nsMgr)
            Dim textNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:text", nsMgr)
            Dim dateNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:date", nsMgr)

            If dropDownNode IsNot Nothing OrElse comboBoxNode IsNot Nothing Then
                ' --- Dropdown / ComboBox ---
                ff.FieldType = FormFieldType.Dropdown
                Dim listParent As System.Xml.XmlNode = If(dropDownNode, comboBoxNode)

                ' Extract options
                Dim listItems As System.Xml.XmlNodeList = listParent.SelectNodes("w:listItem", nsMgr)
                For Each li As System.Xml.XmlNode In listItems
                    Dim displayVal As System.Xml.XmlNode = li.Attributes("w:displayText")
                    Dim valueVal As System.Xml.XmlNode = li.Attributes("w:value")
                    Dim optText As String = If(displayVal?.Value, If(valueVal?.Value, ""))
                    If Not String.IsNullOrEmpty(optText) Then ff.Options.Add(optText)
                Next

                ' Current value = text inside w:sdtContent
                Dim rawValue As String = GetSdtContentText(sdtNode, nsMgr)

                ' The only reliable test: if the display text matches one of the
                ' defined options, it was explicitly selected. Everything else
                ' (locale-specific placeholder, showingPlcHdr flag, or any other
                ' text that isn't an option) means "not yet selected".
                If showingPlcHdr Then
                    ff.CurrentValue = ""
                ElseIf ff.Options.Count > 0 AndAlso
                       ff.Options.Any(Function(o) o.Equals(rawValue, StringComparison.OrdinalIgnoreCase)) Then
                    ff.CurrentValue = rawValue
                Else
                    ' Display text doesn't match any option → unselected
                    ff.CurrentValue = ""
                End If

            ElseIf checkboxNode IsNot Nothing Then
                ' --- Checkbox ---
                ff.FieldType = FormFieldType.Checkbox
                ff.Options.Add("true")
                ff.Options.Add("false")

                Dim checkedNode As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checked/@w14:val", nsMgr)
                If checkedNode IsNot Nothing Then
                    ff.CurrentValue = If(checkedNode.Value = "1" OrElse checkedNode.Value.ToLowerInvariant() = "true", "true", "false")
                Else
                    ff.CurrentValue = "false"
                End If

            ElseIf textNode IsNot Nothing OrElse dateNode IsNot Nothing Then
                ' --- Plain text input or date picker ---
                ff.FieldType = FormFieldType.TextInput
                Dim rawValue As String = GetSdtContentText(sdtNode, nsMgr)
                If showingPlcHdr Then
                    ff.CurrentValue = ""
                Else
                    ff.CurrentValue = rawValue
                End If

            Else
                ' --- Rich-text or untyped content control ---
                ff.FieldType = FormFieldType.TextInput
                Dim rawValue As String = GetSdtContentText(sdtNode, nsMgr)
                If showingPlcHdr Then
                    ff.CurrentValue = ""
                Else
                    ff.CurrentValue = rawValue
                End If
            End If

            cellInfo.FormFields.Add(ff)
            fieldIndex += 1
        Next

        ' === Legacy form fields (w:ffData inside w:fldChar) ===
        Dim fldCharNodes As System.Xml.XmlNodeList = tcNode.SelectNodes(".//w:r[w:fldChar[@w:fldCharType='begin']]", nsMgr)
        For Each fldRunNode As System.Xml.XmlNode In fldCharNodes
            Dim fldChar As System.Xml.XmlNode = fldRunNode.SelectSingleNode("w:fldChar", nsMgr)
            If fldChar Is Nothing Then Continue For

            Dim ffData As System.Xml.XmlNode = fldChar.SelectSingleNode("w:ffData", nsMgr)
            If ffData Is Nothing Then Continue For

            Dim ff As New FormFieldInfo() With {
                .FieldId = $"f{fieldIndex}",
                .XmlNode = fldChar,
                .Options = New List(Of String)(),
                .CurrentValue = ""
            }

            ' Get field name
            Dim nameNode As System.Xml.XmlNode = ffData.SelectSingleNode("w:name/@w:val", nsMgr)
            ff.Name = If(nameNode?.Value, "")

            ' Check for legacy dropdown (w:ddList)
            Dim ddList As System.Xml.XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
            If ddList IsNot Nothing Then
                ff.FieldType = FormFieldType.LegacyDropdown

                Dim listEntries As System.Xml.XmlNodeList = ddList.SelectNodes("w:listEntry/@w:val", nsMgr)
                For Each entry As System.Xml.XmlNode In listEntries
                    ff.Options.Add(entry.Value)
                Next

                Dim resultNode As System.Xml.XmlNode = ddList.SelectSingleNode("w:result/@w:val", nsMgr)
                Dim selectedIdx As Integer = 0
                If resultNode IsNot Nothing Then Integer.TryParse(resultNode.Value, selectedIdx)
                If selectedIdx >= 0 AndAlso selectedIdx < ff.Options.Count Then
                    ff.CurrentValue = ff.Options(selectedIdx)
                End If

                cellInfo.FormFields.Add(ff)
                fieldIndex += 1
                Continue For
            End If

            ' Check for legacy checkbox (w:checkBox)
            Dim checkBox As System.Xml.XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
            If checkBox IsNot Nothing Then
                ff.FieldType = FormFieldType.LegacyCheckbox
                ff.Options.Add("true")
                ff.Options.Add("false")

                Dim checkedNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:checked/@w:val", nsMgr)
                If checkedNode IsNot Nothing Then
                    ff.CurrentValue = If(checkedNode.Value = "1" OrElse checkedNode.Value.ToLowerInvariant() = "true", "true", "false")
                Else
                    Dim defaultNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:default/@w:val", nsMgr)
                    ff.CurrentValue = If(defaultNode IsNot Nothing AndAlso (defaultNode.Value = "1" OrElse defaultNode.Value.ToLowerInvariant() = "true"), "true", "false")
                End If

                cellInfo.FormFields.Add(ff)
                fieldIndex += 1
            End If
        Next
    End Sub

    ''' <summary>
    ''' Gets the plain text content from inside a w:sdt/w:sdtContent node.
    ''' </summary>
    Private Function GetSdtContentText(sdtNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
        If sdtContent Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        For Each ct As System.Xml.XmlNode In sdtContent.SelectNodes(".//w:t", nsMgr)
            sb.Append(ct.InnerText)
        Next
        Return sb.ToString()
    End Function


    Private Function GetParagraphPlainText(paraNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sb As New StringBuilder()
        Dim textNodes As System.Xml.XmlNodeList = paraNode.SelectNodes(".//w:t", nsMgr)
        For Each tNode As System.Xml.XmlNode In textNodes
            sb.Append(tNode.InnerText)
        Next
        Return sb.ToString()
    End Function

#End Region

#Region "JSON Serialization"

    Private Shared Function TableToJson(tbl As TableInfo) As String
        Dim jObj As New JObject()
        jObj("tableIndex") = tbl.TableIndex

        If Not String.IsNullOrWhiteSpace(tbl.ContextBefore) Then
            jObj("contextBefore") = tbl.ContextBefore
        End If

        jObj("columns") = tbl.ColumnCount

        If tbl.ColumnWidths IsNot Nothing AndAlso tbl.ColumnWidths.Count > 0 Then
            Dim jWidths As New JArray()
            For Each w In tbl.ColumnWidths
                jWidths.Add(w)
            Next
            jObj("colWidthsPct") = jWidths
        End If

        Dim jRows As New JArray()
        For Each row In tbl.Rows
            Dim jRow As New JObject()
            jRow("r") = row.RowIndex

            Dim jCells As New JArray()
            For Each cell In row.Cells
                Dim jCell As New JObject()
                jCell("c") = cell.Col
                jCell("t") = If(cell.Text, "")
                If cell.GridSpan > 1 Then jCell("span") = cell.GridSpan
                If cell.VMerge <> "" Then jCell("vm") = cell.VMerge
                If cell.IsHeader Then jCell("hdr") = True
                If cell.IsPlaceholder Then jCell("ph") = True
                If cell.WidthPct > 0 Then jCell("w") = cell.WidthPct

                ' Serialize form fields
                If cell.FormFields IsNot Nothing AndAlso cell.FormFields.Count > 0 Then
                    Dim jFields As New JArray()
                    For Each ff In cell.FormFields
                        Dim jFF As New JObject()
                        jFF("id") = ff.FieldId
                        jFF("type") = ff.FieldType.ToString().ToLowerInvariant()
                        If Not String.IsNullOrWhiteSpace(ff.Name) Then jFF("name") = ff.Name
                        jFF("val") = If(ff.CurrentValue, "")
                        If String.IsNullOrWhiteSpace(ff.CurrentValue) Then jFF("empty") = True
                        If ff.Options IsNot Nothing AndAlso ff.Options.Count > 0 Then
                            Dim jOpts As New JArray()
                            For Each opt In ff.Options
                                jOpts.Add(opt)
                            Next
                            jFF("opts") = jOpts
                        End If
                        jFields.Add(jFF)
                    Next
                    jCell("fields") = jFields
                End If

                jCells.Add(jCell)
            Next

            jRow("cells") = jCells
            jRows.Add(jRow)
        Next

        jObj("rows") = jRows
        Return jObj.ToString(Newtonsoft.Json.Formatting.None)
    End Function

#End Region

#Region "LLM Batching"

    Private Async Function ProcessTableBatches(tables As List(Of TableInfo), userInstructions As String, documentContext As String) As Task(Of List(Of TableUpdateResponse))
        Dim allUpdates As New List(Of TableUpdateResponse)()

        Dim systemPrompt As String =
            "You are a professional document assistant that completes and fills in tables. " &
            "You will receive:" & vbCrLf &
            "- DOCUMENT CONTEXT: the full document body text with [Table N] markers showing where each table sits." & vbCrLf &
            "- TABLE DATA: one or more tables as JSON." & vbCrLf & vbCrLf &
            "JSON SCHEMA for each table:" & vbCrLf &
            "- ""tableIndex"": which table in the document (matches [Table N] in context)" & vbCrLf &
            "- ""colWidthsPct"": array of relative column widths as percentages — narrow columns (< 15%) need short text" & vbCrLf &
            "- ""contextBefore"": paragraphs immediately before the table" & vbCrLf &
            "- Each cell has: ""c"" (column index), ""t"" (text), optionally ""span"" (column span), " &
            """vm"" (vertical merge), ""hdr"" (header), ""ph"" (placeholder to replace), ""w"" (cell width %)" & vbCrLf &
            "- Cells may contain ""fields"": an array of form fields:" & vbCrLf &
            "  - ""id"": field identifier for your response (e.g. ""f0"")" & vbCrLf &
            "  - ""type"": ""dropdown"", ""checkbox"", ""textinput"", ""legacydropdown"", or ""legacycheckbox""" & vbCrLf &
            "  - ""name"": field label/tag (may be empty)" & vbCrLf &
            "  - ""val"": current value" & vbCrLf &
            "  - ""opts"": available options (for dropdowns: selectable items; for checkboxes: [""true"",""false""])" & vbCrLf & vbCrLf &
            "YOUR TASK:" & vbCrLf &
            "1. Use the DOCUMENT CONTEXT to understand the document's purpose, terminology, and style." & vbCrLf &
            "2. Fill in empty cells (""t"": """") with appropriate content based on column headings, row labels, and document context." & vbCrLf &
            "3. Replace ALL placeholder cells (""ph"": true) with real content." & vbCrLf &
            "4. For form fields: select the most appropriate dropdown option or set the correct checkbox state based on context." & vbCrLf &
            "   - For dropdowns: set ""val"" to one of the strings from ""opts""." & vbCrLf &
            "   - For checkboxes: set ""val"" to ""true"" or ""false""." & vbCrLf &
            "   - For text inputs: set ""val"" to the appropriate text." & vbCrLf &
            "   - Fields with ""empty"":true MUST be filled — they are blank content controls waiting for a value." & vbCrLf &
            "   - IMPORTANT: Every field that has ""empty"":true or whose ""val"" is a placeholder like ""Choose an item."" or ""Click or tap here..."" MUST appear in your ""fieldUpdates"" response." & vbCrLf & "5. Keep text length proportional to column width (""w"" field)." & vbCrLf &
            "6. If the table appears incomplete, add new rows." & vbCrLf &
            "7. Do NOT modify cells that already have meaningful non-placeholder content unless the user specifically asks." & vbCrLf &
            "8. Maintain consistent terminology, tone, and style with the rest of the document." & vbCrLf & vbCrLf &
            "RESPONSE FORMAT — return ONLY a JSON array. Each element:" & vbCrLf &
            "{" & vbCrLf &
            "  ""tableIndex"": <int>," & vbCrLf &
            "  ""updates"": [{""r"": <row>, ""c"": <col>, ""t"": ""new text""}]," & vbCrLf &
            "  ""fieldUpdates"": [{""r"": <row>, ""c"": <col>, ""id"": ""f0"", ""val"": ""selected option""}]," & vbCrLf &
            "  ""newRows"": [[""cell1"", ""cell2"", ...]]," & vbCrLf &
            "  ""insertAfterRow"": <int>    // -1 = append at end" & vbCrLf &
            "}" & vbCrLf & vbCrLf &
            "Include ""fieldUpdates"" only for fields you want to change." & vbCrLf &
            "If a table needs no changes, include it with empty arrays." & vbCrLf &
            "Return ONLY the JSON array, no explanations or markdown fences."

        Dim batchStart As Integer = 0
        Dim totalBatches As Integer = CInt(Math.Ceiling(tables.Count / TableCompleteMaxTablesPerBatch))
        ProgressBarModule.GlobalProgressMax = totalBatches
        ProgressBarModule.GlobalProgressValue = 0

        While batchStart < tables.Count
            If ProgressBarModule.CancelOperation Then Return Nothing

            Dim batchEnd As Integer = Math.Min(batchStart + TableCompleteMaxTablesPerBatch - 1, tables.Count - 1)

            Dim batchChars As Integer = 0
            For j As Integer = batchStart To batchEnd
                batchChars += TableToJson(tables(j)).Length
                If batchChars > TableCompleteMaxCharsPerBatch AndAlso j > batchStart Then
                    batchEnd = j - 1
                    Exit For
                End If
            Next

            Dim promptBuilder As New StringBuilder()
            promptBuilder.AppendLine($"USER INSTRUCTIONS: {userInstructions}")
            promptBuilder.AppendLine()

            If Not String.IsNullOrWhiteSpace(documentContext) Then
                promptBuilder.AppendLine("[DOCUMENT CONTEXT]")
                promptBuilder.AppendLine(documentContext)
                promptBuilder.AppendLine("[/DOCUMENT CONTEXT]")
                promptBuilder.AppendLine()
            End If

            promptBuilder.AppendLine("[TABLES]")
            For j As Integer = batchStart To batchEnd
                promptBuilder.AppendLine(TableToJson(tables(j)))
                If j < batchEnd Then promptBuilder.AppendLine(",")
            Next
            promptBuilder.AppendLine("[/TABLES]")

            Dim currentBatch As Integer = CInt(Math.Floor(batchStart / TableCompleteMaxTablesPerBatch)) + 1
            ProgressBarModule.GlobalProgressLabel = $"Processing table batch {currentBatch}/{totalBatches}"

            Dim response As String = Await SharedMethods.LLM(
                _context, systemPrompt, promptBuilder.ToString(),
                "", "", 0, False, True)

            If String.IsNullOrWhiteSpace(response) Then
                ShowCustomMessageBox("LLM returned empty response. Table completion incomplete.")
                Return Nothing
            End If

            Dim batchUpdates As List(Of TableUpdateResponse) = ParseTableCompletionResponse(response)
            If batchUpdates IsNot Nothing Then
                allUpdates.AddRange(batchUpdates)
            End If

            batchStart = batchEnd + 1
            ProgressBarModule.GlobalProgressValue = currentBatch
        End While

        Return allUpdates
    End Function

#End Region

#Region "Response Parsing"

    Private Function ParseTableCompletionResponse(response As String) As List(Of TableUpdateResponse)
        Dim results As New List(Of TableUpdateResponse)()

        Try
            response = response.Trim()
            If response.StartsWith("```") Then
                Dim firstNewline As Integer = response.IndexOf(vbLf)
                If firstNewline > 0 Then response = response.Substring(firstNewline + 1)
                If response.EndsWith("```") Then response = response.Substring(0, response.Length - 3)
                response = response.Trim()
            End If

            Dim parsed As JArray = JArray.Parse(response)

            For Each item As JToken In parsed
                If item.Type <> JTokenType.Object Then Continue For
                Dim jObj As JObject = CType(item, JObject)

                Dim update As New TableUpdateResponse() With {
                    .Updates = New List(Of TableCellUpdate)(),
                    .NewRows = New List(Of List(Of String))(),
                    .FieldUpdates = New List(Of FormFieldUpdate)(),
                    .InsertAfterRow = -1
                }

                If jObj("tableIndex") IsNot Nothing Then update.TableIndex = CInt(jObj("tableIndex"))
                If jObj("insertAfterRow") IsNot Nothing Then update.InsertAfterRow = CInt(jObj("insertAfterRow"))

                ' Parse cell text updates
                Dim updatesToken As JToken = jObj("updates")
                If updatesToken IsNot Nothing AndAlso updatesToken.Type = JTokenType.Array Then
                    For Each updItem As JToken In CType(updatesToken, JArray)
                        If updItem.Type <> JTokenType.Object Then Continue For
                        Dim updObj As JObject = CType(updItem, JObject)
                        Dim cellUpdate As New TableCellUpdate()
                        If updObj("r") IsNot Nothing Then cellUpdate.Row = CInt(updObj("r"))
                        If updObj("c") IsNot Nothing Then cellUpdate.Col = CInt(updObj("c"))
                        If updObj("t") IsNot Nothing Then cellUpdate.Text = CStr(updObj("t"))
                        update.Updates.Add(cellUpdate)
                    Next
                End If

                ' Parse form field updates
                Dim fieldUpdatesToken As JToken = jObj("fieldUpdates")
                If fieldUpdatesToken IsNot Nothing AndAlso fieldUpdatesToken.Type = JTokenType.Array Then
                    For Each fuItem As JToken In CType(fieldUpdatesToken, JArray)
                        If fuItem.Type <> JTokenType.Object Then Continue For
                        Dim fuObj As JObject = CType(fuItem, JObject)
                        Dim fu As New FormFieldUpdate()
                        If fuObj("r") IsNot Nothing Then fu.Row = CInt(fuObj("r"))
                        If fuObj("c") IsNot Nothing Then fu.Col = CInt(fuObj("c"))
                        If fuObj("id") IsNot Nothing Then fu.FieldId = CStr(fuObj("id"))
                        If fuObj("val") IsNot Nothing Then fu.Value = CStr(fuObj("val"))
                        update.FieldUpdates.Add(fu)
                    Next
                End If

                ' Parse new rows
                Dim newRowsToken As JToken = jObj("newRows")
                If newRowsToken IsNot Nothing AndAlso newRowsToken.Type = JTokenType.Array Then
                    For Each rowItem As JToken In CType(newRowsToken, JArray)
                        If rowItem.Type <> JTokenType.Array Then Continue For
                        Dim rowTexts As New List(Of String)()
                        For Each cellVal As JToken In CType(rowItem, JArray)
                            rowTexts.Add(If(cellVal IsNot Nothing, cellVal.ToString(), ""))
                        Next
                        update.NewRows.Add(rowTexts)
                    Next
                End If

                results.Add(update)
            Next

        Catch ex As Exception
            Debug.WriteLine($"ParseTableCompletionResponse error: {ex.Message}")
            ShowCustomMessageBox($"Failed to parse LLM response as JSON: {ex.Message}")
            Return Nothing
        End Try

        Return results
    End Function

#End Region

#Region "Apply Updates"

    Private Sub ApplyTableUpdates(tables As List(Of TableInfo), updates As List(Of TableUpdateResponse), nsMgr As System.Xml.XmlNamespaceManager)
        For Each update In updates
            Dim tbl As TableInfo = tables.FirstOrDefault(Function(t) t.TableIndex = update.TableIndex)
            If tbl Is Nothing Then Continue For

            ' Apply cell text updates
            For Each cellUpdate In update.Updates
                Dim row As TableRowInfo = tbl.Rows.FirstOrDefault(Function(r) r.RowIndex = cellUpdate.Row)
                If row Is Nothing Then Continue For
                Dim cell As TableCellInfo = row.Cells.FirstOrDefault(Function(c) c.Col = cellUpdate.Col)
                If cell Is Nothing Then Continue For

                If cell.TextNodes.Count = 0 Then
                    Dim tcNode As System.Xml.XmlNode = Nothing
                    Dim tcNodes As System.Xml.XmlNodeList = row.XmlNode.SelectNodes("w:tc", nsMgr)
                    Dim cellIdx As Integer = row.Cells.IndexOf(cell)
                    If cellIdx >= 0 AndAlso cellIdx < tcNodes.Count Then tcNode = tcNodes(cellIdx)

                    If tcNode IsNot Nothing Then
                        Dim pNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
                        If pNode Is Nothing Then
                            pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
                            tcNode.AppendChild(pNode)
                        End If
                        Dim rNode As System.Xml.XmlNode = pNode.SelectSingleNode("w:r", nsMgr)
                        If rNode Is Nothing Then
                            rNode = tcNode.OwnerDocument.CreateElement("w", "r", WNs)
                            pNode.AppendChild(rNode)
                        End If
                        Dim tNode As System.Xml.XmlNode = tcNode.OwnerDocument.CreateElement("w", "t", WNs)
                        rNode.AppendChild(tNode)
                        SetTextNodeWithSpacePreserve(tNode, cellUpdate.Text)
                    End If
                ElseIf cell.TextNodes.Count = 1 Then
                    SetTextNodeWithSpacePreserve(cell.TextNodes(0), cellUpdate.Text)
                Else
                    SetTextNodeWithSpacePreserve(cell.TextNodes(0), cellUpdate.Text)
                    For i As Integer = 1 To cell.TextNodes.Count - 1
                        SetTextNodeWithSpacePreserve(cell.TextNodes(i), "")
                    Next
                End If
            Next

            ' Apply form field updates
            If update.FieldUpdates IsNot Nothing Then
                For Each fu In update.FieldUpdates
                    Dim row As TableRowInfo = tbl.Rows.FirstOrDefault(Function(r) r.RowIndex = fu.Row)
                    If row Is Nothing Then Continue For
                    Dim cell As TableCellInfo = row.Cells.FirstOrDefault(Function(c) c.Col = fu.Col)
                    If cell Is Nothing OrElse cell.FormFields Is Nothing Then Continue For

                    Dim ff As FormFieldInfo = cell.FormFields.FirstOrDefault(Function(f) f.FieldId = fu.FieldId)
                    If ff Is Nothing Then Continue For

                    ApplyFormFieldUpdate(ff, fu.Value, nsMgr)
                Next
            End If

            ' Insert new rows
            If update.NewRows IsNot Nothing AndAlso update.NewRows.Count > 0 Then
                InsertNewTableRows(tbl, update.NewRows, update.InsertAfterRow, nsMgr)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Removes the w:showingPlcHdr element from w:sdtPr and strips the PlaceholderText
    ''' run style from all runs inside w:sdtContent. This is REQUIRED when writing a real
    ''' value into an SDT that was previously showing placeholder text, otherwise Word
    ''' continues to render the grey placeholder and ignores the new value.
    ''' </summary>
    Private Sub ClearSdtPlaceholderState(sdtNode As System.Xml.XmlNode, nsMgr As System.Xml.XmlNamespaceManager)
        ' 1. Remove w:showingPlcHdr from w:sdtPr
        Dim showingNode As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w:showingPlcHdr", nsMgr)
        If showingNode IsNot Nothing Then
            showingNode.ParentNode.RemoveChild(showingNode)
        End If

        ' 2. Strip PlaceholderText run style from all runs inside w:sdtContent
        '    The placeholder content has: <w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>
        '    If we leave this, the text renders in grey even after showingPlcHdr is removed.
        Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
        If sdtContent Is Nothing Then Exit Sub

        Dim rStyleNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(
            ".//w:rPr/w:rStyle[@w:val='PlaceholderText' or @w:val='Platzhaltertext']", nsMgr)
        For Each rStyleNode As System.Xml.XmlNode In rStyleNodes
            Dim rPr As System.Xml.XmlNode = rStyleNode.ParentNode
            rPr.RemoveChild(rStyleNode)
            ' If rPr is now empty, remove it too
            If rPr.ChildNodes.Count = 0 Then
                rPr.ParentNode.RemoveChild(rPr)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Applies a value change to a single form field (SDT or legacy).
    ''' For SDTs, also clears the placeholder state so Word renders the new value.
    ''' </summary>
    Private Sub ApplyFormFieldUpdate(ff As FormFieldInfo, newValue As String, nsMgr As System.Xml.XmlNamespaceManager)
        If String.IsNullOrEmpty(newValue) Then Exit Sub

        Select Case ff.FieldType

            Case FormFieldType.Dropdown
                Dim sdtNode As System.Xml.XmlNode = ff.XmlNode
                Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
                If sdtContent Is Nothing Then Exit Sub

                ' Validate: only accept a value that exists in the options
                If ff.Options.Count > 0 AndAlso Not ff.Options.Any(Function(o) o.Equals(newValue, StringComparison.OrdinalIgnoreCase)) Then
                    Exit Sub
                End If

                ' CRITICAL: Remove placeholder state before writing new value
                ClearSdtPlaceholderState(sdtNode, nsMgr)

                ' Update the display text inside sdtContent
                Dim contentTextNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(".//w:t", nsMgr)
                If contentTextNodes.Count > 0 Then
                    SetTextNodeWithSpacePreserve(contentTextNodes(0), newValue)
                    For i As Integer = 1 To contentTextNodes.Count - 1
                        SetTextNodeWithSpacePreserve(contentTextNodes(i), "")
                    Next
                Else
                    ' No w:t node exists yet — create w:r > w:t inside sdtContent
                    ' First clear any existing placeholder content
                    sdtContent.InnerXml = ""
                    Dim rNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "r", WNs)
                    sdtContent.AppendChild(rNode)
                    Dim tNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "t", WNs)
                    rNode.AppendChild(tNode)
                    SetTextNodeWithSpacePreserve(tNode, newValue)
                End If

                ' Update w:sdtPr to reflect the selected item via lastValue attribute
                Dim sdtPr As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr", nsMgr)
                If sdtPr IsNot Nothing Then
                    Dim listNode As System.Xml.XmlNode = sdtPr.SelectSingleNode("w:dropDownList", nsMgr)
                    If listNode Is Nothing Then listNode = sdtPr.SelectSingleNode("w:comboBox", nsMgr)
                    If listNode IsNot Nothing Then
                        ' Find matching listItem by displayText or value
                        For Each li As System.Xml.XmlNode In listNode.SelectNodes("w:listItem", nsMgr)
                            Dim displayAttr As System.Xml.XmlNode = li.Attributes("w:displayText")
                            Dim valueAttr As System.Xml.XmlNode = li.Attributes("w:value")
                            Dim displayText As String = If(displayAttr?.Value, "")
                            Dim valueText As String = If(valueAttr?.Value, "")
                            If displayText.Equals(newValue, StringComparison.OrdinalIgnoreCase) OrElse
                               valueText.Equals(newValue, StringComparison.OrdinalIgnoreCase) Then
                                ' Set w:lastValue on the dropDownList/comboBox element
                                Dim lastValAttr As System.Xml.XmlAttribute = listNode.OwnerDocument.CreateAttribute("w", "lastValue", WNs)
                                lastValAttr.Value = valueText
                                listNode.Attributes.SetNamedItem(lastValAttr)
                                Exit For
                            End If
                        Next
                    End If
                End If

            Case FormFieldType.Checkbox
                Dim sdtNode As System.Xml.XmlNode = ff.XmlNode
                Dim checkboxNode As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtPr/w14:checkbox", nsMgr)
                If checkboxNode Is Nothing Then Exit Sub

                Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")
                Dim xmlVal As String = If(isChecked, "1", "0")

                ' Set w14:checked
                Dim checkedEl As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checked", nsMgr)
                If checkedEl Is Nothing Then
                    checkedEl = checkboxNode.OwnerDocument.CreateElement("w14", "checked", W14Ns)
                    checkboxNode.AppendChild(checkedEl)
                End If
                Dim valAttr As System.Xml.XmlAttribute = checkedEl.OwnerDocument.CreateAttribute("w14", "val", W14Ns)
                valAttr.Value = xmlVal
                checkedEl.Attributes.SetNamedItem(valAttr)

                ' Update the display character in w:sdtContent
                ClearSdtPlaceholderState(sdtNode, nsMgr)
                Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
                If sdtContent IsNot Nothing Then
                    Dim contentTextNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(".//w:t", nsMgr)
                    If contentTextNodes.Count > 0 Then
                        Dim charCode As String
                        If isChecked Then
                            Dim checkedState As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:checkedState/@w14:val", nsMgr)
                            charCode = If(checkedState?.Value, "2612")
                        Else
                            Dim uncheckedState As System.Xml.XmlNode = checkboxNode.SelectSingleNode("w14:uncheckedState/@w14:val", nsMgr)
                            charCode = If(uncheckedState?.Value, "2610")
                        End If
                        Dim displayChar As String = ChrW(System.Convert.ToInt32(charCode, 16)).ToString()
                        SetTextNodeWithSpacePreserve(contentTextNodes(0), displayChar)
                    End If
                End If

            Case FormFieldType.TextInput
                Dim sdtNode As System.Xml.XmlNode = ff.XmlNode

                ' CRITICAL: Remove placeholder state before writing new value
                ClearSdtPlaceholderState(sdtNode, nsMgr)

                Dim sdtContent As System.Xml.XmlNode = sdtNode.SelectSingleNode("w:sdtContent", nsMgr)
                If sdtContent Is Nothing Then Exit Sub
                Dim contentTextNodes As System.Xml.XmlNodeList = sdtContent.SelectNodes(".//w:t", nsMgr)
                If contentTextNodes.Count > 0 Then
                    SetTextNodeWithSpacePreserve(contentTextNodes(0), newValue)
                    For i As Integer = 1 To contentTextNodes.Count - 1
                        SetTextNodeWithSpacePreserve(contentTextNodes(i), "")
                    Next
                Else
                    ' No w:t exists — create structure
                    sdtContent.InnerXml = ""
                    ' Check if this is block-level (needs w:p > w:r > w:t) or inline (needs w:r > w:t)
                    ' Block-level SDTs are direct children of w:tc or w:body
                    Dim parentName As String = If(sdtNode.ParentNode?.LocalName, "")
                    If parentName = "tc" OrElse parentName = "body" Then
                        ' Block-level: needs paragraph wrapper
                        Dim pNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "p", WNs)
                        sdtContent.AppendChild(pNode)
                        Dim rNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "r", WNs)
                        pNode.AppendChild(rNode)
                        Dim tNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "t", WNs)
                        rNode.AppendChild(tNode)
                        SetTextNodeWithSpacePreserve(tNode, newValue)
                    Else
                        ' Inline: just w:r > w:t
                        Dim rNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "r", WNs)
                        sdtContent.AppendChild(rNode)
                        Dim tNode As System.Xml.XmlNode = sdtContent.OwnerDocument.CreateElement("w", "t", WNs)
                        rNode.AppendChild(tNode)
                        SetTextNodeWithSpacePreserve(tNode, newValue)
                    End If
                End If

            Case FormFieldType.LegacyDropdown
                Dim fldCharNode As System.Xml.XmlNode = ff.XmlNode
                Dim ffData As System.Xml.XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                If ffData Is Nothing Then Exit Sub
                Dim ddList As System.Xml.XmlNode = ffData.SelectSingleNode("w:ddList", nsMgr)
                If ddList Is Nothing Then Exit Sub

                Dim targetIdx As Integer = -1
                For i As Integer = 0 To ff.Options.Count - 1
                    If ff.Options(i).Equals(newValue, StringComparison.OrdinalIgnoreCase) Then
                        targetIdx = i
                        Exit For
                    End If
                Next
                If targetIdx < 0 Then Exit Sub

                Dim resultNode As System.Xml.XmlNode = ddList.SelectSingleNode("w:result", nsMgr)
                If resultNode Is Nothing Then
                    resultNode = ddList.OwnerDocument.CreateElement("w", "result", WNs)
                    ddList.AppendChild(resultNode)
                End If
                Dim resultValAttr As System.Xml.XmlAttribute = resultNode.OwnerDocument.CreateAttribute("w", "val", WNs)
                resultValAttr.Value = targetIdx.ToString()
                resultNode.Attributes.SetNamedItem(resultValAttr)

                UpdateLegacyFieldDisplayText(fldCharNode, newValue, nsMgr)

            Case FormFieldType.LegacyCheckbox
                Dim fldCharNode As System.Xml.XmlNode = ff.XmlNode
                Dim ffData As System.Xml.XmlNode = fldCharNode.SelectSingleNode("w:ffData", nsMgr)
                If ffData Is Nothing Then Exit Sub
                Dim checkBox As System.Xml.XmlNode = ffData.SelectSingleNode("w:checkBox", nsMgr)
                If checkBox Is Nothing Then Exit Sub

                Dim isChecked As Boolean = (newValue.ToLowerInvariant() = "true" OrElse newValue = "1")

                Dim checkedNode As System.Xml.XmlNode = checkBox.SelectSingleNode("w:checked", nsMgr)
                If checkedNode Is Nothing Then
                    checkedNode = checkBox.OwnerDocument.CreateElement("w", "checked", WNs)
                    checkBox.AppendChild(checkedNode)
                End If
                Dim checkedValAttr As System.Xml.XmlAttribute = checkedNode.OwnerDocument.CreateAttribute("w", "val", WNs)
                checkedValAttr.Value = If(isChecked, "1", "0")
                checkedNode.Attributes.SetNamedItem(checkedValAttr)

        End Select
    End Sub

    ''' <summary>
    ''' Updates the display text for a legacy form field dropdown.
    ''' Walks sibling runs after the fldChar begin until fldChar separate/end is found.
    ''' </summary>
    Private Sub UpdateLegacyFieldDisplayText(fldCharBeginRun As System.Xml.XmlNode, newText As String, nsMgr As System.Xml.XmlNamespaceManager)
        ' fldCharBeginRun is the w:fldChar node itself; its parent is the w:r containing begin
        Dim beginRun As System.Xml.XmlNode = fldCharBeginRun.ParentNode
        If beginRun Is Nothing Then Exit Sub

        Dim currentNode As System.Xml.XmlNode = beginRun.NextSibling
        Dim foundSeparate As Boolean = False
        Dim textUpdated As Boolean = False

        While currentNode IsNot Nothing
            If currentNode.LocalName = "r" Then
                ' Check for fldChar separate or end
                Dim fc As System.Xml.XmlNode = currentNode.SelectSingleNode("w:fldChar", nsMgr)
                If fc IsNot Nothing Then
                    Dim fcType As System.Xml.XmlNode = fc.Attributes("w:fldCharType")
                    If fcType IsNot Nothing Then
                        If fcType.Value = "separate" Then
                            foundSeparate = True
                        ElseIf fcType.Value = "end" Then
                            Exit While
                        End If
                    End If
                ElseIf foundSeparate Then
                    ' This run contains the display text — update it
                    Dim tNodes As System.Xml.XmlNodeList = currentNode.SelectNodes("w:t", nsMgr)
                    If tNodes.Count > 0 AndAlso Not textUpdated Then
                        SetTextNodeWithSpacePreserve(tNodes(0), newText)
                        textUpdated = True
                        For i As Integer = 1 To tNodes.Count - 1
                            SetTextNodeWithSpacePreserve(tNodes(i), "")
                        Next
                    ElseIf tNodes.Count > 0 Then
                        ' Clear additional runs after the first
                        For Each tn As System.Xml.XmlNode In tNodes
                            SetTextNodeWithSpacePreserve(tn, "")
                        Next
                    End If
                End If
            End If
            currentNode = currentNode.NextSibling
        End While
    End Sub

    Private Sub InsertNewTableRows(tbl As TableInfo, newRows As List(Of List(Of String)), insertAfterRow As Integer, nsMgr As System.Xml.XmlNamespaceManager)
        Dim templateRow As TableRowInfo = Nothing
        For i As Integer = tbl.Rows.Count - 1 To 0 Step -1
            If Not tbl.Rows(i).Cells.All(Function(c) c.IsHeader) Then
                templateRow = tbl.Rows(i)
                Exit For
            End If
        Next
        If templateRow Is Nothing AndAlso tbl.Rows.Count > 0 Then
            templateRow = tbl.Rows(tbl.Rows.Count - 1)
        End If
        If templateRow Is Nothing Then Exit Sub

        Dim refNode As System.Xml.XmlNode
        If insertAfterRow >= 0 AndAlso insertAfterRow < tbl.Rows.Count Then
            refNode = tbl.Rows(insertAfterRow).XmlNode
        Else
            refNode = tbl.Rows(tbl.Rows.Count - 1).XmlNode
        End If

        For Each rowTexts In newRows
            Dim clonedRow As System.Xml.XmlNode = templateRow.XmlNode.CloneNode(deep:=True)

            Dim tblHeaderNode As System.Xml.XmlNode = clonedRow.SelectSingleNode("w:trPr/w:tblHeader", nsMgr)
            If tblHeaderNode IsNot Nothing Then tblHeaderNode.ParentNode.RemoveChild(tblHeaderNode)

            Dim vMergeNodes As System.Xml.XmlNodeList = clonedRow.SelectNodes(".//w:tcPr/w:vMerge", nsMgr)
            For Each vmNode As System.Xml.XmlNode In vMergeNodes
                vmNode.ParentNode.RemoveChild(vmNode)
            Next

            Dim tcNodes As System.Xml.XmlNodeList = clonedRow.SelectNodes("w:tc", nsMgr)
            For ci As Integer = 0 To tcNodes.Count - 1
                Dim tcNode As System.Xml.XmlNode = tcNodes(ci)
                Dim cellText As String = If(ci < rowTexts.Count, rowTexts(ci), "")

                Dim tNodes As System.Xml.XmlNodeList = tcNode.SelectNodes(".//w:t", nsMgr)
                If tNodes.Count > 0 Then
                    SetTextNodeWithSpacePreserve(tNodes(0), cellText)
                    For ti As Integer = 1 To tNodes.Count - 1
                        SetTextNodeWithSpacePreserve(tNodes(ti), "")
                    Next
                Else
                    Dim pNode As System.Xml.XmlNode = tcNode.SelectSingleNode("w:p", nsMgr)
                    If pNode Is Nothing Then
                        pNode = tcNode.OwnerDocument.CreateElement("w", "p", WNs)
                        tcNode.AppendChild(pNode)
                    End If
                    Dim rNode As System.Xml.XmlNode = pNode.SelectSingleNode("w:r", nsMgr)
                    If rNode Is Nothing Then
                        rNode = tcNode.OwnerDocument.CreateElement("w", "r", WNs)
                        pNode.AppendChild(rNode)
                    End If
                    Dim tNode As System.Xml.XmlNode = tcNode.OwnerDocument.CreateElement("w", "t", WNs)
                    rNode.AppendChild(tNode)
                    SetTextNodeWithSpacePreserve(tNode, cellText)
                End If
            Next

            tbl.XmlNode.InsertAfter(clonedRow, refNode)
            refNode = clonedRow
        Next
    End Sub

#End Region

End Class