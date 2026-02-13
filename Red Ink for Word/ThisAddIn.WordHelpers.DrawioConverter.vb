' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WordHelpers.DrawioConverter.vb
' Purpose: Converts a .drawio diagram file into a self-contained interactive
'          HTML flowchart navigator.
'
' Architecture / Key Ideas:
'  - XML Parsing: Reads .drawio XML and extracts mxCell elements, separating
'    vertices (nodes) and edges (arrows) by their attributes.
'  - Object Wrappers: Handles <object>/<UserObject> elements that wrap mxCell
'    children — the id/value live on the wrapper, not the inner cell.
'  - Compressed Diagrams: Handles both raw <mxGraphModel> XML and compressed
'    (base64 + deflate) content stored inside <diagram> elements.
'  - Multi-Page Support: Processes all <diagram> pages in a single .drawio file,
'    merging vertices and edges across pages.
'  - Classification: Classifies each vertex as start (no incoming), decision
'    (incoming + outgoing), end (incoming only), or isolated — matching the
'    original Python reference implementation. Filters unlabeled arrows when
'    labeled ones exist for a source.
'  - Smart Buttons: Single-outgoing nodes show "Start" or "Next" instead of
'    repeating the upcoming node text, avoiding duplicate statements.
'  - Edge Cases: Self-loops, dangling edges, cyclic graphs, multiple start
'    nodes, empty node text, duplicate edges.
'  - HTML Generation: Standalone HTML with embedded JSON, Red Ink logo at
'    bottom, user-provided title, optional Home link, consistent square
'    styling with #B91312 accent.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' =========================================================================
    ' Draw.io Data Models
    ' =========================================================================

    Private Class DrawioVertex
        Public Property Id As String
        Public Property Text As String
    End Class

    Private Class DrawioEdge
        Public Property Source As String
        Public Property Target As String
        Public Property Label As String
    End Class

    Private Class DrawioOutgoingArrow
        Public Property TargetId As String
        Public Property Label As String
    End Class

    Private Class DrawioIncomingArrow
        Public Property SourceId As String
        Public Property Label As String
    End Class

    Private Class DrawioElement
        Public Property Id As String
        Public Property Text As String
        Public Property ElementType As String  ' "start", "decision", "end", "isolated"
        Public Property Incoming As List(Of DrawioIncomingArrow)
        Public Property Outgoing As List(Of DrawioOutgoingArrow)
    End Class

    ' =========================================================================
    ' Entry Point
    ' =========================================================================

    ''' <summary>
    ''' Entry point for converting a single .drawio file to interactive HTML.
    ''' Prompts for file selection via DragDropForm, asks for a title and an
    ''' optional Home link, then converts.
    ''' </summary>
    Public Sub ConvertDrawioToHtml()
        Dim selectedPath As String = ""

        Globals.ThisAddIn.DragDropFormLabel = "Select a .drawio file to convert"
        Globals.ThisAddIn.DragDropFormFilter = "Draw.io Diagrams|*.drawio|All Files (*.*)|*.*"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOnly)
                If frm.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
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
        If ext <> ".drawio" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Please select a .drawio file.")
            Exit Sub
        End If

        ' Ask user for a title
        Dim defaultTitle As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim userTitle As String = ShowCustomInputBox(
            "Enter a title for the flowchart navigator:",
            AN & " Draw.io Converter", True, defaultTitle)
        If String.IsNullOrWhiteSpace(userTitle) Then Exit Sub
        userTitle = userTitle.Trim()

        ' Ask user for an optional Home link
        Dim homeLink As String = ShowCustomInputBox(
            "Enter a Home URL (leave empty to skip):" & vbCrLf &
            "A Home button will appear so users can navigate back to this URL at any time.",
            AN & " Draw.io Converter", True, "")
        ' ShowCustomInputBox returns "" on cancel for SimpleInput mode — treat both
        ' empty and cancelled as "no home link"
        If homeLink IsNot Nothing Then homeLink = homeLink.Trim()
        If String.IsNullOrEmpty(homeLink) Then homeLink = Nothing

        ' Output path: same directory and name, .html extension
        Dim dir As String = Path.GetDirectoryName(selectedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputPath As String = Path.Combine(dir, $"{nameWithoutExt}.html")

        Try
            Dim success As Boolean = ConvertSingleDrawioFile(selectedPath, outputPath, userTitle, homeLink)

            If success Then
                ShowCustomMessageBox($"Converted successfully:{vbCrLf}{vbCrLf}{Path.GetFileName(outputPath)}", AN & " Draw.io Converter")
            Else
                ShowCustomMessageBox($"Conversion failed. The file may not contain a valid flowchart.", AN & " Draw.io Converter")
            End If
        Catch ex As Exception
            ShowCustomMessageBox($"Error during conversion:{vbCrLf}{vbCrLf}{ex.Message}", AN & " Draw.io Converter")
        End Try
    End Sub

    ' =========================================================================
    ' Core Conversion
    ' =========================================================================

    Private Function ConvertSingleDrawioFile(inputPath As String, outputPath As String,
                                             title As String, homeLink As String) As Boolean
        Try
            Dim content As String = File.ReadAllText(inputPath, Encoding.UTF8)

            Dim vertices As Dictionary(Of String, DrawioVertex) = Nothing
            Dim edges As List(Of DrawioEdge) = Nothing

            ParseDrawio(content, vertices, edges)

            If vertices.Count = 0 Then
                Debug.WriteLine($"ConvertSingleDrawioFile: No vertices found in {Path.GetFileName(inputPath)}")
                Return False
            End If

            Dim elements As List(Of DrawioElement) = ClassifyDrawioElements(vertices, edges)

            If elements.Count = 0 Then
                Debug.WriteLine($"ConvertSingleDrawioFile: No elements after classification in {Path.GetFileName(inputPath)}")
                Return False
            End If

            Dim flowchartData As New JObject()
            flowchartData("elements") = SerializeDrawioElements(elements)

            Dim htmlOutput As String = GenerateDrawioHtml(flowchartData, title, homeLink)
            File.WriteAllText(outputPath, htmlOutput, Encoding.UTF8)

            Return True

        Catch ex As XmlException
            Debug.WriteLine($"ConvertSingleDrawioFile XML error: {ex.Message}")
            Return False
        Catch ex As Exception
            Debug.WriteLine($"ConvertSingleDrawioFile error: {ex.Message}")
            Return False
        End Try
    End Function

    ' =========================================================================
    ' XML Parsing
    ' =========================================================================

    Private Shared Sub ParseDrawio(content As String,
                                   ByRef vertices As Dictionary(Of String, DrawioVertex),
                                   ByRef edges As List(Of DrawioEdge))

        vertices = New Dictionary(Of String, DrawioVertex)(StringComparer.Ordinal)
        edges = New List(Of DrawioEdge)()

        Dim rootDoc As New XmlDocument()
        rootDoc.LoadXml(content)

        Dim xmlDocsToParse As New List(Of XmlDocument)()

        Dim diagramNodes As XmlNodeList = rootDoc.SelectNodes("//diagram")
        If diagramNodes IsNot Nothing AndAlso diagramNodes.Count > 0 Then
            For Each diagramNode As XmlNode In diagramNodes
                Dim inlineModel As XmlNode = diagramNode.SelectSingleNode("mxGraphModel")
                If inlineModel IsNot Nothing Then
                    Dim inlineDoc As New XmlDocument()
                    inlineDoc.LoadXml(inlineModel.OuterXml)
                    xmlDocsToParse.Add(inlineDoc)
                Else
                    Dim compressedContent As String = diagramNode.InnerText?.Trim()
                    If Not String.IsNullOrEmpty(compressedContent) Then
                        Dim decompressed As String = TryDecompressDrawioContent(compressedContent)
                        If Not String.IsNullOrEmpty(decompressed) Then
                            Try
                                Dim decompDoc As New XmlDocument()
                                decompDoc.LoadXml(decompressed)
                                xmlDocsToParse.Add(decompDoc)
                            Catch ex As XmlException
                                Debug.WriteLine($"ParseDrawio: Failed to parse decompressed diagram: {ex.Message}")
                            End Try
                        End If
                    End If
                End If
            Next
        End If

        If xmlDocsToParse.Count = 0 Then
            xmlDocsToParse.Add(rootDoc)
        End If

        For Each xmlDocToParse As XmlDocument In xmlDocsToParse
            ExtractCellsFromXml(xmlDocToParse, vertices, edges)
        Next
    End Sub

    Private Shared Function TryDecompressDrawioContent(compressed As String) As String
        Try
            Dim bytes As Byte() = System.Convert.FromBase64String(compressed)
            Dim decompressed As String
            Using inputStream As New MemoryStream(bytes)
                Using deflateStream As New DeflateStream(inputStream, CompressionMode.Decompress)
                    Using reader As New StreamReader(deflateStream, Encoding.UTF8)
                        decompressed = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            decompressed = System.Net.WebUtility.UrlDecode(decompressed)
            Return decompressed
        Catch ex As Exception
            Debug.WriteLine($"TryDecompressDrawioContent failed: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Shared Sub ExtractCellsFromXml(xmlDoc As XmlDocument,
                                           vertices As Dictionary(Of String, DrawioVertex),
                                           edges As List(Of DrawioEdge))

        ' --- 1. Direct mxCell elements (skip those wrapped by object/UserObject) ---
        For Each cell As XmlNode In xmlDoc.SelectNodes("//mxCell")
            Dim parentName As String = cell.ParentNode?.Name
            If parentName = "object" OrElse parentName = "UserObject" Then Continue For

            Dim cellId As String = cell.Attributes("id")?.Value
            If String.IsNullOrEmpty(cellId) Then Continue For
            If cellId = "0" OrElse cellId = "1" Then Continue For

            If cell.Attributes("vertex")?.Value = "1" Then
                vertices(cellId) = New DrawioVertex() With {
                    .Id = cellId,
                    .Text = If(cell.Attributes("value")?.Value, "")
                }
            ElseIf cell.Attributes("edge")?.Value = "1" Then
                Dim source As String = cell.Attributes("source")?.Value
                Dim target As String = cell.Attributes("target")?.Value
                Dim label As String = If(cell.Attributes("value")?.Value, "")
                If Not String.IsNullOrEmpty(source) AndAlso Not String.IsNullOrEmpty(target) Then
                    edges.Add(New DrawioEdge() With {.Source = source, .Target = target, .Label = label})
                End If
            End If
        Next

        ' --- 2. <object> and <UserObject> wrappers (id/label on wrapper) ---
        For Each wrapperTag As String In {"object", "UserObject"}
            For Each wrapper As XmlNode In xmlDoc.SelectNodes("//" & wrapperTag)
                Dim wrapperId As String = wrapper.Attributes("id")?.Value
                If String.IsNullOrEmpty(wrapperId) Then Continue For
                If wrapperId = "0" OrElse wrapperId = "1" Then Continue For

                Dim wrapperLabel As String = If(wrapper.Attributes("label")?.Value, "")
                Dim innerCell As XmlNode = wrapper.SelectSingleNode("mxCell")
                If innerCell Is Nothing Then Continue For

                If innerCell.Attributes("vertex")?.Value = "1" Then
                    vertices(wrapperId) = New DrawioVertex() With {.Id = wrapperId, .Text = wrapperLabel}
                ElseIf innerCell.Attributes("edge")?.Value = "1" Then
                    Dim source As String = innerCell.Attributes("source")?.Value
                    Dim target As String = innerCell.Attributes("target")?.Value
                    If Not String.IsNullOrEmpty(source) AndAlso Not String.IsNullOrEmpty(target) Then
                        edges.Add(New DrawioEdge() With {.Source = source, .Target = target, .Label = wrapperLabel})
                    End If
                End If
            Next
        Next
    End Sub

    ' =========================================================================
    ' Element Classification — faithful to the Python reference implementation
    ' =========================================================================

    Private Shared Function ClassifyDrawioElements(vertices As Dictionary(Of String, DrawioVertex),
                                                   edges As List(Of DrawioEdge)) As List(Of DrawioElement)

        ' Filter out dangling edges
        Dim validEdges As List(Of DrawioEdge) = edges.Where(
            Function(e) vertices.ContainsKey(e.Source) AndAlso vertices.ContainsKey(e.Target)
        ).ToList()

        ' Group outgoing arrows by source
        Dim outgoingBySource As New Dictionary(Of String, List(Of DrawioEdge))(StringComparer.Ordinal)
        For Each vid As String In vertices.Keys
            outgoingBySource(vid) = New List(Of DrawioEdge)()
        Next
        For Each validEdge As DrawioEdge In validEdges
            If outgoingBySource.ContainsKey(validEdge.Source) Then
                outgoingBySource(validEdge.Source).Add(validEdge)
            End If
        Next

        ' Build filtered outgoing/incoming — skip unlabeled if source has labeled ones
        Dim outgoing As New Dictionary(Of String, List(Of DrawioOutgoingArrow))(StringComparer.Ordinal)
        Dim incoming As New Dictionary(Of String, List(Of DrawioIncomingArrow))(StringComparer.Ordinal)

        For Each vid As String In vertices.Keys
            outgoing(vid) = New List(Of DrawioOutgoingArrow)()
            incoming(vid) = New List(Of DrawioIncomingArrow)()
        Next

        For Each kvp As KeyValuePair(Of String, List(Of DrawioEdge)) In outgoingBySource
            Dim source As String = kvp.Key
            Dim arrows As List(Of DrawioEdge) = kvp.Value
            Dim hasLabeled As Boolean = arrows.Any(Function(a) Not String.IsNullOrEmpty(a.Label))

            For Each arrow As DrawioEdge In arrows
                If Not String.IsNullOrEmpty(arrow.Label) OrElse Not hasLabeled Then
                    outgoing(source).Add(New DrawioOutgoingArrow() With {
                        .TargetId = arrow.Target, .Label = arrow.Label
                    })
                    incoming(arrow.Target).Add(New DrawioIncomingArrow() With {
                        .SourceId = source, .Label = arrow.Label
                    })
                End If
            Next
        Next

        ' Classify — matching Python: has_incoming / has_outgoing booleans
        Dim elements As New List(Of DrawioElement)()

        For Each kvp As KeyValuePair(Of String, DrawioVertex) In vertices
            Dim vid As String = kvp.Key
            Dim vertex As DrawioVertex = kvp.Value

            Dim hasIncoming As Boolean = incoming(vid).Count > 0
            Dim hasOutgoing As Boolean = outgoing(vid).Count > 0

            Dim elementType As String
            If hasOutgoing AndAlso Not hasIncoming Then
                elementType = "start"
            ElseIf hasIncoming AndAlso hasOutgoing Then
                elementType = "decision"
            ElseIf hasIncoming AndAlso Not hasOutgoing Then
                elementType = "end"
            Else
                elementType = "isolated"
            End If

            elements.Add(New DrawioElement() With {
                .Id = vid,
                .Text = vertex.Text,
                .ElementType = elementType,
                .Incoming = If(hasIncoming, incoming(vid), Nothing),
                .Outgoing = If(hasOutgoing, outgoing(vid), Nothing)
            })
        Next

        ' Cyclic graph fallback: promote best candidate to start
        Dim hasStart As Boolean = elements.Any(Function(e) e.ElementType = "start")
        If Not hasStart Then
            Dim candidates = elements.Where(Function(e) e.ElementType = "decision").ToList()
            If candidates.Count = 0 Then
                candidates = elements.Where(Function(e) e.ElementType <> "isolated").ToList()
            End If
            If candidates.Count > 0 Then
                Dim best As DrawioElement = candidates.OrderBy(
                    Function(e) If(e.Incoming IsNot Nothing, e.Incoming.Count, 0)
                ).ThenByDescending(
                    Function(e) If(e.Outgoing IsNot Nothing, e.Outgoing.Count, 0)
                ).First()
                best.ElementType = "start"
            End If
        End If

        Return elements
    End Function

    ' =========================================================================
    ' JSON Serialization
    ' =========================================================================

    Private Shared Function SerializeDrawioElements(elements As List(Of DrawioElement)) As JArray
        Dim arr As New JArray()

        For Each el As DrawioElement In elements
            Dim obj As New JObject()
            obj("id") = el.Id
            obj("text") = el.Text
            obj("type") = el.ElementType

            If el.Incoming IsNot Nothing AndAlso el.Incoming.Count > 0 Then
                Dim incArr As New JArray()
                For Each inc As DrawioIncomingArrow In el.Incoming
                    Dim incObj As New JObject()
                    incObj("source_id") = inc.SourceId
                    incObj("label") = inc.Label
                    incArr.Add(incObj)
                Next
                obj("incoming") = incArr
            End If

            If el.Outgoing IsNot Nothing AndAlso el.Outgoing.Count > 0 Then
                Dim outArr As New JArray()
                For Each outArrow As DrawioOutgoingArrow In el.Outgoing
                    Dim outObj As New JObject()
                    outObj("target_id") = outArrow.TargetId
                    outObj("label") = outArrow.Label
                    outArr.Add(outObj)
                Next
                obj("outgoing") = outArr
            End If

            arr.Add(obj)
        Next

        Return arr
    End Function

    ' =========================================================================
    ' HTML Generation
    ' =========================================================================

    Private Shared Function GenerateDrawioHtml(flowchartData As JObject, title As String,
                                               homeLink As String) As String
        Dim jsonData As String = flowchartData.ToString(Newtonsoft.Json.Formatting.Indented)
        Dim safeTitle As String = System.Net.WebUtility.HtmlEncode(If(title, "Flowchart Navigator"))
        Dim jsSafeTitle As String = If(title, "Flowchart Navigator").
            Replace("\", "\\").Replace("'", "\'").Replace("""", "\""")

        ' Prepare safe home link for JS — either a quoted string or null
        Dim jsHomeLink As String = "null"
        If Not String.IsNullOrEmpty(homeLink) Then
            jsHomeLink = "'" & homeLink.Replace("\", "\\").Replace("'", "\'") & "'"
        End If

        Dim html As String = DrawioHtmlTemplate
        html = html.Replace("{{FLOWCHART_DATA}}", jsonData)
        html = html.Replace("{{TITLE}}", safeTitle)
        html = html.Replace("{{JS_TITLE}}", jsSafeTitle)
        html = html.Replace("{{HOME_LINK}}", jsHomeLink)
        Return html
    End Function

    Private Const DrawioHtmlTemplate As String =
"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>{{TITLE}}</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            background: #0a0a0a;
            color: #fafafa;
        }
        .container {
            width: 100%; max-width: 480px; padding: 2rem;
            flex: 1; display: flex; flex-direction: column; justify-content: center;
        }
        .header { text-align: center; margin-bottom: 2rem; }
        .header h1 {
            font-size: 1.1rem; font-weight: 400;
            letter-spacing: 0.05em; color: #ccc;
        }
        .node-text {
            font-size: 1.5rem; font-weight: 300; text-align: center;
            padding: 3rem 1rem; line-height: 1.5;
            border-bottom: 1px solid #222; margin-bottom: 2rem;
        }
        #end-badge { display: none; margin-bottom: 2rem; text-align: center; }
        .end-marker {
            display: inline-block;
            font-size: 0.65rem; font-weight: 500;
            letter-spacing: 0.2em; text-transform: uppercase;
            color: #B91312; padding: 0.5rem 1.5rem;
            border: 1px solid #B91312;
        }
        .choice-buttons { display: flex; flex-direction: column; gap: 0.75rem; }
        .choice-btn {
            background: transparent; border: 1px solid #333;
            color: #fafafa; padding: 1rem 1.5rem;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            font-size: 0.9rem; font-weight: 400;
            cursor: pointer; transition: all 0.15s ease; text-align: left;
        }
        .choice-btn:hover {
            background: #fafafa; color: #0a0a0a; border-color: #fafafa;
        }
        .choice-btn.accent {
            border-color: #B91312; color: #B91312; text-align: center;
        }
        .choice-btn.accent:hover {
            background: #B91312; color: #fafafa; border-color: #B91312;
        }
        .nav-buttons { display: flex; gap: 0.5rem; margin-top: 1rem; }
        .nav-btn {
            display: none; flex: 1;
            background: transparent; border: 1px solid #333;
            color: #666; padding: 0.75rem 1.5rem;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            font-size: 0.8rem; cursor: pointer;
            transition: all 0.15s ease;
        }
        .nav-btn:hover { color: #fafafa; border-color: #666; }
        .footer {
            text-align: center; padding: 1.5rem 1rem 1rem;
            width: 100%; max-width: 480px;
        }
        .footer-logo { margin-bottom: 0.75rem; }
        .footer-logo img { height: 32px; width: auto; opacity: 0.7; }
        .footer-text { font-size: 0.65rem; color: #444; }
        .footer-text a {
            color: #666; text-decoration: none; transition: color 0.15s ease;
        }
        .footer-text a:hover { color: #fafafa; }
        #isolated-section {
            display: none; margin-top: 2rem;
            padding-top: 1rem; border-top: 1px solid #222;
        }
        #isolated-section h2 {
            font-size: 0.7rem; font-weight: 500;
            letter-spacing: 0.15em; text-transform: uppercase;
            color: #666; margin-bottom: 1rem;
        }
        .isolated-node {
            padding: 0.75rem 1rem; border: 1px solid #222;
            margin-bottom: 0.5rem; font-size: 0.85rem; color: #999;
        }
    </style>
</head>
<body>
    <div class=""container"">
        <div class=""header""><h1>{{TITLE}}</h1></div>
        <div id=""node-display"" class=""node-text""></div>
        <div id=""end-badge""><span class=""end-marker"">Complete</span></div>
        <div id=""choice-buttons"" class=""choice-buttons""></div>
        <div class=""nav-buttons"">
            <button id=""back-btn"" class=""nav-btn"">Back</button>
            <button id=""restart-btn"" class=""nav-btn"">Start over</button>
            <button id=""home-btn"" class=""nav-btn"">Home</button>
        </div>
        <div id=""isolated-section"">
            <h2>Additional Notes</h2>
            <div id=""isolated-list""></div>
        </div>
    </div>
    <div class=""footer"">
        <div class=""footer-logo"">
            <a href=""https://redink.ai"" target=""_blank"" rel=""noopener"">
                <img src=""https://redink.ai/content/Red_Ink_Logo_Large.png""
                     alt=""Red Ink"" onerror=""this.parentElement.style.display='none'"">
            </a>
        </div>
        <div class=""footer-text"">
            Created with <a href=""https://redink.ai"" target=""_blank"" rel=""noopener"">Red Ink</a>
        </div>
    </div>
    <script>
        document.title = '{{JS_TITLE}}';
        var homeLink = {{HOME_LINK}};
        var flowchartData = {{FLOWCHART_DATA}};
        var nodeMap = {};
        flowchartData.elements.forEach(function(el) { nodeMap[el.id] = el; });

        var startNode = flowchartData.elements.find(function(el) { return el.type === 'start'; });
        if (!startNode) {
            startNode = flowchartData.elements.find(function(el) { return el.type !== 'isolated'; });
        }
        if (!startNode && flowchartData.elements.length > 0) {
            startNode = flowchartData.elements[0];
        }

        var navHistory = [];
        var nodeDisplay = document.getElementById('node-display');
        var endBadge = document.getElementById('end-badge');
        var choiceButtons = document.getElementById('choice-buttons');
        var backBtn = document.getElementById('back-btn');
        var restartBtn = document.getElementById('restart-btn');
        var homeBtn = document.getElementById('home-btn');

        /* Home button: always visible if homeLink is set */
        if (homeLink) {
            homeBtn.style.display = 'inline-block';
            homeBtn.addEventListener('click', function() {
                window.location.href = homeLink;
            });
        }

        function sanitizeText(text) {
            if (!text) return '';
            return text
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<\/?div>/gi, '\n')
                .replace(/<\/?p>/gi, '\n')
                .replace(/<\/?span[^>]*>/gi, '')
                .replace(/<\/?b>/gi, '')
                .replace(/<\/?i>/gi, '')
                .replace(/<\/?u>/gi, '')
                .replace(/<\/?font[^>]*>/gi, '')
                .replace(/&amp;/gi, '&')
                .replace(/&lt;/gi, '<')
                .replace(/&gt;/gi, '>')
                .replace(/&quot;/gi, '""')
                .replace(/&#39;/gi, ""'"")
                .replace(/&nbsp;/gi, ' ')
                .replace(/<[^>]*>/g, '')
                .replace(/\n{3,}/g, '\n\n')
                .trim();
        }

        function renderNode(nodeId, addToHistory) {
            if (addToHistory === undefined) addToHistory = true;
            var node = nodeMap[nodeId];
            if (!node) { nodeDisplay.textContent = 'Error: Node not found'; return; }
            if (addToHistory && navHistory[navHistory.length - 1] !== nodeId) {
                navHistory.push(nodeId);
            }

            var displayText = sanitizeText(node.text);
            nodeDisplay.textContent = displayText || '...';
            choiceButtons.innerHTML = '';
            backBtn.style.display = navHistory.length > 1 ? 'inline-block' : 'none';

            var isEnd = (node.type === 'end' || !node.outgoing || node.outgoing.length === 0);

            if (isEnd) {
                endBadge.style.display = 'block';
                restartBtn.style.display = 'inline-block';
            } else {
                endBadge.style.display = 'none';
                restartBtn.style.display = 'none';

                /* De-duplicate outgoing arrows */
                var seen = {};
                var uniqueOutgoing = node.outgoing.filter(function(arrow) {
                    var key = arrow.target_id + '|' + (arrow.label || '');
                    if (seen[key]) return false;
                    seen[key] = true;
                    return true;
                });

                var isSinglePath = (uniqueOutgoing.length === 1);
                var isFirstNode = (navHistory.length <= 1);

                uniqueOutgoing.forEach(function(arrow) {
                    var btn = document.createElement('button');
                    btn.className = 'choice-btn';

                    if (arrow.label) {
                        /* Labeled arrow: always use the label */
                        btn.textContent = arrow.label;
                    } else if (isSinglePath) {
                        /* Single unlabeled arrow: use Start/Next instead of
                           repeating the upcoming node text */
                        btn.textContent = isFirstNode ? 'Start' : 'Next';
                        btn.className = 'choice-btn accent';
                    } else {
                        /* Multiple unlabeled arrows: show target node text */
                        var targetNode = nodeMap[arrow.target_id];
                        btn.textContent = (targetNode && sanitizeText(targetNode.text))
                            ? sanitizeText(targetNode.text) : 'Next';
                    }

                    btn.addEventListener('click', function() { renderNode(arrow.target_id); });
                    choiceButtons.appendChild(btn);
                });
            }
        }

        backBtn.addEventListener('click', function() {
            if (navHistory.length > 1) {
                navHistory.pop();
                renderNode(navHistory[navHistory.length - 1], false);
            }
        });
        restartBtn.addEventListener('click', function() {
            navHistory.length = 0;
            if (startNode) renderNode(startNode.id);
        });

        /* Isolated nodes as static list */
        var isolatedNodes = flowchartData.elements.filter(function(el) {
            return el.type === 'isolated';
        });
        if (isolatedNodes.length > 0) {
            var isolatedSection = document.getElementById('isolated-section');
            var isolatedList = document.getElementById('isolated-list');
            isolatedNodes.forEach(function(el) {
                var nodeText = sanitizeText(el.text);
                if (nodeText) {
                    var div = document.createElement('div');
                    div.className = 'isolated-node';
                    div.textContent = nodeText;
                    isolatedList.appendChild(div);
                }
            });
            if (isolatedList.children.length > 0) {
                isolatedSection.style.display = 'block';
            }
        }

        if (startNode) { renderNode(startNode.id); }
        else { nodeDisplay.textContent = 'No flowchart nodes found.'; }
    </script>
</body>
</html>"

End Class