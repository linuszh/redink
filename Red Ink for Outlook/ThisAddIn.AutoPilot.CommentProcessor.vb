' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.CommentProcessor.vb
' Purpose: Adds Word comment bubbles to a DOCX file by having the LLM identify
'          relevant text portions and provide comments for each.
'          Operates directly on DOCX OpenXML — no Word interop dependency.
'
' Architecture:
'  - Extracts all text from the document, sends it to the LLM with the
'    SP_Add_Bubbles prompt so the LLM can see the entire text and determine
'    which portions need comments.
'  - Parses the LLM response (text1@@comment1§§§text2@@comment2§§§...).
'  - Locates each quoted text in the OpenXML and inserts w:commentRangeStart,
'    w:commentRangeEnd, and w:comment elements.
'  - Ensures comments.xml is registered in [Content_Types].xml and
'    word/_rels/document.xml.rels.
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
    '  CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Separator between comment entries in the LLM response.
    ''' </summary>
    Private Const AP_CommentSeparator As String = "§§§"

    ''' <summary>
    ''' Delimiter between quoted text and comment text in a single entry.
    ''' </summary>
    Private Const AP_CommentDelimiter As String = "@@"

    ''' <summary>
    ''' Maximum characters intended for a single LLM batch.
    ''' </summary>
    Private Const AP_CommentMaxCharsPerBatch As Integer = 30000

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Represents one quoted text segment and its comment text.
    ''' </summary>
    Private Class APCommentEntry
        Public Property QuotedText As String
        Public Property CommentText As String
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MAIN ENTRY POINT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Processes a DOCX file by having the LLM add comment bubbles to relevant sections
    ''' based on the given instruction.
    ''' </summary>
    ''' <param name="author">Optional author name for the comments. If Nothing or empty, defaults to AN6 ("Inky").
    ''' When a custom author is used, each comment is prefixed with "RI: " (AN5).</param>
    Private Async Function CommentDocxForAutoPilot(inputPath As String, outputPath As String,
                                                    instruction As String, ct As CancellationToken,
                                                    Optional author As String = Nothing) As Task(Of Boolean)
        Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & "cmt_" & Guid.NewGuid().ToString("N"))

        ' Resolve the effective author and determine whether to prefix comments
        Dim effectiveAuthor As String = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
        Dim usePrefix As Boolean = Not effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase)

        Try
            File.Copy(inputPath, outputPath, overwrite:=True)
            ZipFile.ExtractToDirectory(outputPath, tempDir)

            Dim documentXmlPath = Path.Combine(tempDir, "word", "document.xml")
            If Not File.Exists(documentXmlPath) Then Return False

            Dim xmlDoc As New System.Xml.XmlDocument()
            xmlDoc.PreserveWhitespace = True
            xmlDoc.Load(documentXmlPath)

            Dim nsMgr As New System.Xml.XmlNamespaceManager(xmlDoc.NameTable)
            nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
            nsMgr.AddNamespace("w14", "http://schemas.microsoft.com/office/word/2010/wordml")
            nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            ' Extract full document text for the LLM
            Dim fullText = APExtractFullText(xmlDoc, nsMgr)
            If String.IsNullOrWhiteSpace(fullText) Then Return False

            ' Call LLM with the bubbles prompt
            Dim comments = Await APGetCommentsFromLLM(fullText, instruction, ct)
            If comments Is Nothing OrElse comments.Count = 0 Then Return False

            ' Insert comments into the OpenXML
            Dim commentsXmlPath = Path.Combine(tempDir, "word", "comments.xml")
            APInsertCommentsIntoDocx(xmlDoc, nsMgr, comments, documentXmlPath, commentsXmlPath, tempDir,
                                     effectiveAuthor, usePrefix)

            ' Repack DOCX
            File.Delete(outputPath)
            ZipFile.CreateFromDirectory(tempDir, outputPath, CompressionLevel.Optimal, False)

            Return True

        Catch ex As System.Exception
            Debug.WriteLine("CommentDocxForAutoPilot error: " & ex.Message)
            Return False
        Finally
            If Directory.Exists(tempDir) Then
                Try : Directory.Delete(tempDir, recursive:=True) : Catch : End Try
            End If
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT EXTRACTION (full document as plain text for LLM)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts full document text from all paragraph text nodes.
    ''' </summary>
    Private Function APExtractFullText(xmlDoc As System.Xml.XmlDocument,
                                        nsMgr As System.Xml.XmlNamespaceManager) As String
        Dim sb As New StringBuilder()
        Dim paraNodes = xmlDoc.SelectNodes("//w:p", nsMgr)

        For Each paraNode As System.Xml.XmlNode In paraNodes
            Dim textNodes = paraNode.SelectNodes(".//w:t", nsMgr)
            Dim paraText As New StringBuilder()
            For Each textNode As System.Xml.XmlNode In textNodes
                paraText.Append(textNode.InnerText)
            Next
            If paraText.Length > 0 Then
                sb.AppendLine(paraText.ToString())
            End If
        Next

        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  LLM CALL
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Sends the full text to the LLM with the bubbles prompt and instruction,
    ''' then parses the text@@comment§§§ response format.
    ''' </summary>
    Private Async Function APGetCommentsFromLLM(fullText As String, instruction As String,
                                                 ct As CancellationToken) As Task(Of List(Of APCommentEntry))
        ct.ThrowIfCancellationRequested()

        ' Build the bubbles system prompt
        Dim bubblesPrompt As String = InterpolateAtRuntime(SP_Add_Bubbles)

        ' Replace {FormatInstruction} placeholder — for offline processing we want plain text comments
        bubblesPrompt = bubblesPrompt.Replace("{FormatInstruction}",
            "Provide each comment as plain text without any markdown formatting.")

        Dim systemPrompt As String =
            "You are a professional document reviewer. " &
            "Apply the following instruction to the document text provided." & vbCrLf & vbCrLf &
            "INSTRUCTION: " & instruction & vbCrLf & vbCrLf &
            bubblesPrompt

        ' Build user prompt — batch if text is very large
        Dim userPrompt As String = "[TEXTTOPROCESS]" & vbCrLf & fullText & vbCrLf & "[/TEXTTOPROCESS]"

        Dim llmResponse = Await LLM(systemPrompt, userPrompt,
                                     UseSecondAPI:=False,
                                     HideSplash:=True, EnsureUI:=False,
                                     cancellationToken:=ct)

        If String.IsNullOrWhiteSpace(llmResponse) Then Return Nothing

        Return APParseBubblesResponse(llmResponse)
    End Function

    ''' <summary>
    ''' Parses the LLM response in the format: text1@@comment1§§§text2@@comment2§§§...
    ''' </summary>
    Private Function APParseBubblesResponse(response As String) As List(Of APCommentEntry)
        Dim results As New List(Of APCommentEntry)()

        ' Clean up the response — remove leading/trailing whitespace and any wrapping quotes
        response = response.Trim()
        If response.StartsWith("""") AndAlso response.EndsWith("""") Then
            response = response.Substring(1, response.Length - 2)
        End If

        ' Split on §§§
        Dim pairs = response.Split({AP_CommentSeparator}, StringSplitOptions.RemoveEmptyEntries)

        For Each pair In pairs
            Dim delimIdx = pair.IndexOf(AP_CommentDelimiter, StringComparison.Ordinal)
            If delimIdx > 0 Then
                Dim quotedText = pair.Substring(0, delimIdx).Trim()
                Dim commentText = pair.Substring(delimIdx + AP_CommentDelimiter.Length).Trim()

                If Not String.IsNullOrWhiteSpace(quotedText) AndAlso Not String.IsNullOrWhiteSpace(commentText) Then
                    results.Add(New APCommentEntry() With {
                        .QuotedText = quotedText,
                        .CommentText = commentText
                    })
                End If
            End If
        Next

        Return results
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  COMMENT INSERTION INTO OPENXML
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Inserts Word comments into the DOCX by locating quoted text spans and
    ''' adding commentRangeStart/End markers plus a comments.xml part.
    ''' </summary>
    ''' <param name="commentAuthor">The author name to set on each comment.</param>
    ''' <param name="prefixComments">If True, each comment body is prefixed with "RI: " (AN5).</param>
    Private Sub APInsertCommentsIntoDocx(xmlDoc As System.Xml.XmlDocument,
                                          nsMgr As System.Xml.XmlNamespaceManager,
                                          comments As List(Of APCommentEntry),
                                          documentXmlPath As String,
                                          commentsXmlPath As String,
                                          tempDir As String,
                                          commentAuthor As String,
                                          prefixComments As Boolean)

        Const wNs As String = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        ' Determine the starting comment ID (if comments.xml already exists, continue from max)
        Dim nextCommentId As Integer = 0
        Dim commentsXmlDoc As System.Xml.XmlDocument = Nothing
        Dim commentsRoot As System.Xml.XmlNode = Nothing

        If File.Exists(commentsXmlPath) Then
            commentsXmlDoc = New System.Xml.XmlDocument()
            commentsXmlDoc.PreserveWhitespace = True
            commentsXmlDoc.Load(commentsXmlPath)

            Dim cmNsMgr As New System.Xml.XmlNamespaceManager(commentsXmlDoc.NameTable)
            cmNsMgr.AddNamespace("w", wNs)
            commentsRoot = commentsXmlDoc.SelectSingleNode("//w:comments", cmNsMgr)

            ' Find max existing comment ID
            Dim existingComments = commentsXmlDoc.SelectNodes("//w:comment", cmNsMgr)
            For Each c As System.Xml.XmlNode In existingComments
                Dim idAttr = c.Attributes("w:id")
                If idAttr IsNot Nothing Then
                    Dim existingId As Integer
                    If Integer.TryParse(idAttr.Value, existingId) AndAlso existingId >= nextCommentId Then
                        nextCommentId = existingId + 1
                    End If
                End If
            Next
        Else
            ' Create a new comments.xml
            commentsXmlDoc = New System.Xml.XmlDocument()
            commentsXmlDoc.PreserveWhitespace = True

            Dim declaration = commentsXmlDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
            commentsXmlDoc.AppendChild(declaration)

            commentsRoot = commentsXmlDoc.CreateElement("w", "comments", wNs)
            Dim mcAttr = commentsXmlDoc.CreateAttribute("xmlns", "r", "http://www.w3.org/2000/xmlns/")
            mcAttr.Value = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            commentsRoot.Attributes.Append(mcAttr)
            commentsXmlDoc.AppendChild(commentsRoot)
        End If

        ' Build a flat list of (textNode, runningStartPos) for text searching
        Dim textNodeMap = APBuildTextNodeMap(xmlDoc, nsMgr)

        ' Derive initials from the author name
        Dim initials As String = APDeriveInitials(commentAuthor)

        ' Build the comment text prefix (e.g. "RI: ") when using a custom author
        Dim commentPrefix As String = If(prefixComments, AN5 & ": ", "")

        Dim insertedCount As Integer = 0

        For Each entry In comments
            Dim commentId = nextCommentId
            nextCommentId += 1

            ' Find the quoted text in the document
            Dim span = APFindTextSpan(textNodeMap, entry.QuotedText)
            If span Is Nothing Then
                Debug.WriteLine("CommentProcessor: Could not find quoted text: " & entry.QuotedText.Substring(0, Math.Min(60, entry.QuotedText.Length)))
                Continue For
            End If

            ' Insert commentRangeStart before the first run containing the match
            Dim startRun = span.StartTextNode.ParentNode  ' w:r
            Dim startPara = APFindAncestorParagraph(startRun, nsMgr)
            If startPara Is Nothing Then Continue For

            Dim rangeStart = xmlDoc.CreateElement("w", "commentRangeStart", wNs)
            Dim startIdAttr = xmlDoc.CreateAttribute("w", "id", wNs)
            startIdAttr.Value = commentId.ToString()
            rangeStart.Attributes.Append(startIdAttr)
            startRun.ParentNode.InsertBefore(rangeStart, startRun)

            ' Insert commentRangeEnd after the last run containing the match
            Dim endRun = span.EndTextNode.ParentNode  ' w:r
            Dim rangeEnd = xmlDoc.CreateElement("w", "commentRangeEnd", wNs)
            Dim endIdAttr = xmlDoc.CreateAttribute("w", "id", wNs)
            endIdAttr.Value = commentId.ToString()
            rangeEnd.Attributes.Append(endIdAttr)
            endRun.ParentNode.InsertAfter(rangeEnd, endRun)

            ' Insert w:r > w:commentReference after the commentRangeEnd
            Dim refRun = xmlDoc.CreateElement("w", "r", wNs)
            Dim rPr = xmlDoc.CreateElement("w", "rPr", wNs)
            Dim rStyle = xmlDoc.CreateElement("w", "rStyle", wNs)
            Dim styleAttr = xmlDoc.CreateAttribute("w", "val", wNs)
            styleAttr.Value = "CommentReference"
            rStyle.Attributes.Append(styleAttr)
            rPr.AppendChild(rStyle)
            refRun.AppendChild(rPr)
            Dim commentRef = xmlDoc.CreateElement("w", "commentReference", wNs)
            Dim refIdAttr = xmlDoc.CreateAttribute("w", "id", wNs)
            refIdAttr.Value = commentId.ToString()
            commentRef.Attributes.Append(refIdAttr)
            refRun.AppendChild(commentRef)
            endRun.ParentNode.InsertAfter(refRun, rangeEnd)

            ' Add the comment element to comments.xml
            Dim commentEl = commentsXmlDoc.CreateElement("w", "comment", wNs)
            Dim cmtIdAttr = commentsXmlDoc.CreateAttribute("w", "id", wNs)
            cmtIdAttr.Value = commentId.ToString()
            commentEl.Attributes.Append(cmtIdAttr)

            Dim authorAttr = commentsXmlDoc.CreateAttribute("w", "author", wNs)
            authorAttr.Value = commentAuthor
            commentEl.Attributes.Append(authorAttr)

            Dim dateAttr = commentsXmlDoc.CreateAttribute("w", "date", wNs)
            dateAttr.Value = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
            commentEl.Attributes.Append(dateAttr)

            Dim initialsAttr = commentsXmlDoc.CreateAttribute("w", "initials", wNs)
            initialsAttr.Value = initials
            commentEl.Attributes.Append(initialsAttr)

            ' Build comment body: w:p > w:r > w:t
            Dim cmtPara = commentsXmlDoc.CreateElement("w", "p", wNs)
            Dim cmtRun = commentsXmlDoc.CreateElement("w", "r", wNs)
            Dim cmtText = commentsXmlDoc.CreateElement("w", "t", wNs)

            Dim spaceAttr = commentsXmlDoc.CreateAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace")
            spaceAttr.Value = "preserve"
            cmtText.Attributes.Append(spaceAttr)
            cmtText.InnerText = commentPrefix & entry.CommentText

            cmtRun.AppendChild(cmtText)
            cmtPara.AppendChild(cmtRun)
            commentEl.AppendChild(cmtPara)
            commentsRoot.AppendChild(commentEl)

            insertedCount += 1
        Next

        If insertedCount = 0 Then Return

        ' Save document.xml and comments.xml
        xmlDoc.Save(documentXmlPath)
        commentsXmlDoc.Save(commentsXmlPath)

        ' Ensure the comments.xml part is referenced in [Content_Types].xml and document.xml.rels
        APEnsureCommentsRelationship(tempDir)
    End Sub

    ''' <summary>
    ''' Derives initials from an author name. For multi-word names, takes the first letter
    ''' of each word (e.g. "John Smith" → "JS"). For single-word names, takes up to the
    ''' first two characters (e.g. "Inky" → "IN").
    ''' </summary>
    Private Shared Function APDeriveInitials(authorName As String) As String
        If String.IsNullOrWhiteSpace(authorName) Then Return ""
        Dim parts = authorName.Trim().Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
        If parts.Length > 1 Then
            Return String.Concat(parts.Select(Function(p) p(0).ToString().ToUpper()))
        End If
        Return authorName.Substring(0, Math.Min(2, authorName.Length)).ToUpper()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TEXT SPAN SEARCH HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Maps a text node to its position in the concatenated document text.
    ''' </summary>
    Private Class APTextNodeMapEntry
        Public Property TextNode As System.Xml.XmlNode
        Public Property StartPos As Integer
        Public Property EndPos As Integer
    End Class

    ''' <summary>
    ''' Represents the start and end text nodes for a matched span.
    ''' </summary>
    Private Class APTextSpanResult
        Public Property StartTextNode As System.Xml.XmlNode
        Public Property EndTextNode As System.Xml.XmlNode
    End Class

    ''' <summary>
    ''' Builds a flat map of all w:t nodes with their character positions in the document stream.
    ''' </summary>
    Private Function APBuildTextNodeMap(xmlDoc As System.Xml.XmlDocument,
                                        nsMgr As System.Xml.XmlNamespaceManager) As List(Of APTextNodeMapEntry)
        Dim map As New List(Of APTextNodeMapEntry)()
        Dim allTextNodes = xmlDoc.SelectNodes("//w:t", nsMgr)
        Dim pos As Integer = 0

        For Each textNode As System.Xml.XmlNode In allTextNodes
            Dim text = textNode.InnerText
            If text.Length > 0 Then
                map.Add(New APTextNodeMapEntry() With {
                    .TextNode = textNode,
                    .StartPos = pos,
                    .EndPos = pos + text.Length - 1
                })
                pos += text.Length
            End If
        Next

        Return map
    End Function

    ''' <summary>
    ''' Finds the text nodes spanning the given search text in the concatenated document text.
    ''' Uses a normalized search to be tolerant of minor whitespace differences.
    ''' </summary>
    Private Function APFindTextSpan(textNodeMap As List(Of APTextNodeMapEntry),
                                     searchText As String) As APTextSpanResult
        If textNodeMap.Count = 0 OrElse String.IsNullOrWhiteSpace(searchText) Then Return Nothing

        ' Build concatenated text from the map
        Dim fullTextBuilder As New StringBuilder()
        For Each entry In textNodeMap
            fullTextBuilder.Append(entry.TextNode.InnerText)
        Next
        Dim fullText = fullTextBuilder.ToString()

        ' Try exact match first
        Dim matchIdx = fullText.IndexOf(searchText, StringComparison.Ordinal)

        ' Try case-insensitive if exact fails
        If matchIdx < 0 Then
            matchIdx = fullText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase)
        End If

        ' Try normalized whitespace match
        If matchIdx < 0 Then
            Dim normalizedFull = Regex.Replace(fullText, "\s+", " ")
            Dim normalizedSearch = Regex.Replace(searchText, "\s+", " ")
            Dim normIdx = normalizedFull.IndexOf(normalizedSearch, StringComparison.OrdinalIgnoreCase)
            If normIdx >= 0 Then
                ' Map normalized index back to original — approximate by scanning
                matchIdx = APMapNormalizedIndexToOriginal(fullText, normIdx, normalizedSearch.Length)
            End If
        End If

        If matchIdx < 0 Then Return Nothing

        Dim matchEnd = matchIdx + searchText.Length - 1

        ' Find start and end text nodes
        Dim startNode As System.Xml.XmlNode = Nothing
        Dim endNode As System.Xml.XmlNode = Nothing

        For Each entry In textNodeMap
            If startNode Is Nothing AndAlso entry.EndPos >= matchIdx Then
                startNode = entry.TextNode
            End If
            If entry.StartPos <= matchEnd Then
                endNode = entry.TextNode
            End If
            If entry.StartPos > matchEnd Then Exit For
        Next

        If startNode Is Nothing OrElse endNode Is Nothing Then Return Nothing

        Return New APTextSpanResult() With {
            .StartTextNode = startNode,
            .EndTextNode = endNode
        }
    End Function

    ''' <summary>
    ''' Approximately maps a position in normalized (collapsed-whitespace) text back to the original.
    ''' </summary>
    Private Shared Function APMapNormalizedIndexToOriginal(original As String, normalizedIdx As Integer, normalizedLen As Integer) As Integer
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
    ''' Walks up from a node to find the w:p ancestor.
    ''' </summary>
    Private Shared Function APFindAncestorParagraph(node As System.Xml.XmlNode,
                                                     nsMgr As System.Xml.XmlNamespaceManager) As System.Xml.XmlNode
        Dim current = node
        While current IsNot Nothing
            If current.LocalName = "p" AndAlso current.NamespaceURI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main" Then
                Return current
            End If
            current = current.ParentNode
        End While
        Return Nothing
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  RELATIONSHIP & CONTENT TYPE MANAGEMENT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Ensures that comments.xml is registered in [Content_Types].xml and
    ''' word/_rels/document.xml.rels so Word will load the comments part.
    ''' </summary>
    Private Sub APEnsureCommentsRelationship(tempDir As String)
        Const commentsRelType As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        Const commentsContentType As String = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"

        ' --- [Content_Types].xml ---
        Dim contentTypesPath = Path.Combine(tempDir, "[Content_Types].xml")
        If File.Exists(contentTypesPath) Then
            Dim ctDoc As New System.Xml.XmlDocument()
            ctDoc.PreserveWhitespace = True
            ctDoc.Load(contentTypesPath)

            Dim ctNs As String = "http://schemas.openxmlformats.org/package/2006/content-types"
            Dim ctNsMgr As New System.Xml.XmlNamespaceManager(ctDoc.NameTable)
            ctNsMgr.AddNamespace("ct", ctNs)

            ' Check if Override for /word/comments.xml already exists
            Dim existing = ctDoc.SelectSingleNode("//ct:Override[@PartName='/word/comments.xml']", ctNsMgr)
            If existing Is Nothing Then
                Dim typesNode = ctDoc.DocumentElement
                Dim overrideEl = ctDoc.CreateElement("Override", ctNs)

                Dim partNameAttr = ctDoc.CreateAttribute("PartName")
                partNameAttr.Value = "/word/comments.xml"
                overrideEl.Attributes.Append(partNameAttr)

                Dim contentTypeAttr = ctDoc.CreateAttribute("ContentType")
                contentTypeAttr.Value = commentsContentType
                overrideEl.Attributes.Append(contentTypeAttr)

                typesNode.AppendChild(overrideEl)
                ctDoc.Save(contentTypesPath)
            End If
        End If

        ' --- word/_rels/document.xml.rels ---
        Dim relsDir = Path.Combine(tempDir, "word", "_rels")
        Dim relsPath = Path.Combine(relsDir, "document.xml.rels")

        If Not Directory.Exists(relsDir) Then
            Directory.CreateDirectory(relsDir)
        End If

        Dim relsNs As String = "http://schemas.openxmlformats.org/package/2006/relationships"

        If File.Exists(relsPath) Then
            Dim relsDoc As New System.Xml.XmlDocument()
            relsDoc.PreserveWhitespace = True
            relsDoc.Load(relsPath)

            Dim relsNsMgr As New System.Xml.XmlNamespaceManager(relsDoc.NameTable)
            relsNsMgr.AddNamespace("r", relsNs)

            ' Check if relationship already exists
            Dim existingRel = relsDoc.SelectSingleNode(
                "//r:Relationship[@Type='" & commentsRelType & "']", relsNsMgr)
            If existingRel Is Nothing Then
                ' Determine next rId
                Dim maxId As Integer = 0
                Dim allRels = relsDoc.SelectNodes("//r:Relationship", relsNsMgr)
                For Each rel As System.Xml.XmlNode In allRels
                    Dim idVal = rel.Attributes("Id")?.Value
                    If idVal IsNot Nothing AndAlso idVal.StartsWith("rId") Then
                        Dim num As Integer
                        If Integer.TryParse(idVal.Substring(3), num) AndAlso num > maxId Then
                            maxId = num
                        End If
                    End If
                Next

                Dim newRel = relsDoc.CreateElement("Relationship", relsNs)

                Dim idAttr = relsDoc.CreateAttribute("Id")
                idAttr.Value = "rId" & (maxId + 1).ToString()
                newRel.Attributes.Append(idAttr)

                Dim typeAttr = relsDoc.CreateAttribute("Type")
                typeAttr.Value = commentsRelType
                newRel.Attributes.Append(typeAttr)

                Dim targetAttr = relsDoc.CreateAttribute("Target")
                targetAttr.Value = "comments.xml"
                newRel.Attributes.Append(targetAttr)

                relsDoc.DocumentElement.AppendChild(newRel)
                relsDoc.Save(relsPath)
            End If
        Else
            ' Create a minimal .rels file
            Dim relsDoc As New System.Xml.XmlDocument()
            Dim decl = relsDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes")
            relsDoc.AppendChild(decl)

            Dim root = relsDoc.CreateElement("Relationships", relsNs)
            relsDoc.AppendChild(root)

            Dim newRel = relsDoc.CreateElement("Relationship", relsNs)

            Dim idAttr = relsDoc.CreateAttribute("Id")
            idAttr.Value = "rId1"
            newRel.Attributes.Append(idAttr)

            Dim typeAttr = relsDoc.CreateAttribute("Type")
            typeAttr.Value = commentsRelType
            newRel.Attributes.Append(typeAttr)

            Dim targetAttr = relsDoc.CreateAttribute("Target")
            targetAttr.Value = "comments.xml"
            newRel.Attributes.Append(targetAttr)

            root.AppendChild(newRel)
            relsDoc.Save(relsPath)
        End If
    End Sub

End Class