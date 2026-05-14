' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WordTools.vb
' Purpose: File-level (.docx) tools using DocumentFormat.OpenXml. Operate ONLY
'          on file paths sandboxed by PathPolicy. Never touches the currently
'          open Word document (that surface is worddoc_* in step 10C).
'
' Tools:
'   word_extract_text   — return plain text of the document.
'   word_search         — substring/regex search; returns W.Paragraph index + match.
'   word_write          — replace/insert/append W.Paragraph text WITHOUT markup.
'   word_markup         — same as word_write but produces tracked-change w:ins / w:del.
'   word_comment_add    — add a Word comment to a matched substring.
'   word_comment_list   — list all comments.
'   word_comment_remove — remove a comment by id.
'   word_format         — set W.Paragraph style and/or W.Run formatting on a match.
'   word_apply_template — clone a template from a skill's references/ dir and
'                         substitute {{placeholders}} from a JSON map.
'   word_save_as        — copy a .docx to a new path (workspace or Desktop).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports W = DocumentFormat.OpenXml.Wordprocessing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public NotInheritable Class WordTools

        Private Sub New()
        End Sub

        Public Const ToolExtract As String = "word_extract_text"
        Public Const ToolSearch As String = "word_search"
        Public Const ToolWrite As String = "word_write"
        Public Const ToolMarkup As String = "word_markup"
        Public Const ToolCommentAdd As String = "word_comment_add"
        Public Const ToolCommentList As String = "word_comment_list"
        Public Const ToolCommentRemove As String = "word_comment_remove"
        Public Const ToolFormat As String = "word_format"
        Public Const ToolApplyTemplate As String = "word_apply_template"
        Public Const ToolSaveAs As String = "word_save_as"

        Public Shared Function IsWordTool(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False

            Select Case name
                Case ToolExtract, ToolSearch, ToolWrite, ToolMarkup,
                     ToolCommentAdd, ToolCommentList, ToolCommentRemove,
                     ToolFormat, ToolApplyTemplate, ToolSaveAs
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Public Shared Function BuildAll() As List(Of ModelConfig)
            Return New List(Of ModelConfig) From {
                BuildExtract(), BuildSearch(), BuildWrite(), BuildMarkup(),
                BuildCommentAdd(), BuildCommentList(), BuildCommentRemove(),
                BuildFormat(), BuildApplyTemplate(), BuildSaveAs()
            }
        End Function

        ' --------------------------------------------------------------- dispatch

        Public Shared Function Execute(toolName As String, arguments As IDictionary(Of String, Object)) As String
            Try
                Select Case toolName
                    Case ToolExtract
                        Return ExecuteExtract(arguments)
                    Case ToolSearch
                        Return ExecuteSearch(arguments)
                    Case ToolWrite
                        Return ExecuteWriteOrMarkup(arguments, asMarkup:=False)
                    Case ToolMarkup
                        Return ExecuteWriteOrMarkup(arguments, asMarkup:=True)
                    Case ToolCommentAdd
                        Return ExecuteCommentAdd(arguments)
                    Case ToolCommentList
                        Return ExecuteCommentList(arguments)
                    Case ToolCommentRemove
                        Return ExecuteCommentRemove(arguments)
                    Case ToolFormat
                        Return ExecuteFormat(arguments)
                    Case ToolApplyTemplate
                        Return ExecuteApplyTemplate(arguments)
                    Case ToolSaveAs
                        Return ExecuteSaveAs(arguments)
                    Case Else
                        Return Err_("unknown_word_tool", "Unknown tool '" & toolName & "'.")
                End Select
            Catch uae As UnauthorizedAccessException
                Return Err_("access_denied", uae.Message)
            Catch ex As Exception
                Return Err_("word_tool_failed", ex.Message)
            End Try
        End Function

        ' --------------------------------------------------------------- extract / search

        Private Shared Function ExecuteExtract(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=False)
                Dim paragraphs As List(Of ParagraphRow) = ExtractParagraphs(doc)
                Dim joined As String = String.Join(Environment.NewLine, paragraphs.Select(Function(t) t.Text))
                Dim maxChars As Integer = GetInt(args, "max_chars", 0)
                Dim truncated As Boolean = False

                If maxChars > 0 AndAlso joined.Length > maxChars Then
                    joined = joined.Substring(0, maxChars)
                    truncated = True
                End If

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .paragraphs = paragraphs.Count,
                    Key .truncated = truncated,
                    Key .text = joined
                })
            End Using
        End Function

        Private Shared Function ExecuteSearch(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim query As String = GetStr(args, "query")
            If String.IsNullOrWhiteSpace(query) Then Return Err_("missing_query", "query is required.")

            Dim useRegex As Boolean = GetBool(args, "regex", False)
            Dim ignoreCase As Boolean = GetBool(args, "ignore_case", True)
            Dim maxHits As Integer = System.Math.Min(System.Math.Max(GetInt(args, "max_hits", 50), 1), 500)

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=False)
                Dim paragraphs As List(Of ParagraphRow) = ExtractParagraphs(doc)
                Dim hits As New List(Of Object)()
                Dim cmp As StringComparison = If(ignoreCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
                Dim rx As Regex = Nothing

                If useRegex Then
                    Dim opt As RegexOptions = RegexOptions.CultureInvariant
                    If ignoreCase Then opt = opt Or RegexOptions.IgnoreCase
                    rx = New Regex(query, opt, TimeSpan.FromSeconds(2))
                End If

                For idx As Integer = 0 To paragraphs.Count - 1
                    Dim text As String = paragraphs(idx).Text

                    If useRegex Then
                        For Each m As Match In rx.Matches(text)
                            If hits.Count >= maxHits Then Exit For
                            hits.Add(BuildHit(idx, text, m.Index, m.Length, m.Value))
                        Next
                    Else
                        Dim pos As Integer = 0
                        While pos < text.Length
                            Dim f As Integer = text.IndexOf(query, pos, cmp)
                            If f < 0 Then Exit While

                            hits.Add(BuildHit(idx, text, f, query.Length, text.Substring(f, query.Length)))
                            If hits.Count >= maxHits Then Exit While

                            pos = f + System.Math.Max(1, query.Length)
                        End While
                    End If

                    If hits.Count >= maxHits Then Exit For
                Next

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .total = hits.Count,
                    Key .hits = hits
                })
            End Using
        End Function

        ' --------------------------------------------------------------- write / markup

        Private Shared Function ExecuteWriteOrMarkup(args As IDictionary(Of String, Object), asMarkup As Boolean) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim op As String = GetStr(args, "op").ToLowerInvariant() ' replace | insert_before | insert_after | append
            Dim find As String = GetStr(args, "find")
            Dim text As String = GetStr(args, "text")
            Dim author As String = If(GetStr(args, "author"), "Red Ink")
            Dim onlyFirst As Boolean = GetBool(args, "only_first", True)

            If String.IsNullOrWhiteSpace(op) Then op = "replace"
            If op <> "append" AndAlso String.IsNullOrWhiteSpace(find) Then
                Return Err_("missing_find", "find is required for op '" & op & "'.")
            End If

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim body As W.Body = doc.MainDocumentPart.Document.Body
                Dim mutated As Integer = 0

                If op = "append" Then
                    Dim para As New W.Paragraph()
                    Dim run As W.Run = MakeRun(text)

                    If asMarkup Then
                        Dim ins As New W.InsertedRun() With {
                            .Id = NextChangeId(body).ToString(),
                            .Author = author,
                            .Date = DateTime.UtcNow
                        }
                        ins.AppendChild(run)
                        para.AppendChild(ins)
                    Else
                        para.AppendChild(run)
                    End If

                    body.AppendChild(para)
                    mutated = 1
                Else
                    For Each para As W.Paragraph In body.Descendants(Of W.Paragraph)().ToList()
                        Dim pt As String = GetParagraphText(para)
                        Dim idx As Integer = pt.IndexOf(find, StringComparison.Ordinal)
                        If idx < 0 Then Continue For

                        ApplyTextOpToParagraph(para, find, text, op, asMarkup, author, body)
                        mutated += 1

                        If onlyFirst Then Exit For
                    Next
                End If

                If mutated = 0 Then
                    Return Err_("no_match", "No W.Paragraph contained the 'find' text.")
                End If

                doc.MainDocumentPart.Document.Save()
            End Using

            Return JsonConvert.SerializeObject(New With {
                Key .path = p,
                Key .mutated = True,
                Key .op = op,
                Key .markup = asMarkup
            })
        End Function

        Private Shared Sub ApplyTextOpToParagraph(para As W.Paragraph,
                                                   find As String,
                                                   newText As String,
                                                   op As String,
                                                   asMarkup As Boolean,
                                                   author As String,
                                                   body As W.Body)
            ' Naive strategy: rebuild W.Paragraph from its plain text. Preserves the W.Paragraph's pPr;
            ' loses inline W.Run formatting that splits across the matched substring. Acceptable for
            ' agent-level operations and matches Claude Cowork's "Edit W.Paragraph" behavior.
            Dim pt As String = GetParagraphText(para)
            Dim idx As Integer = pt.IndexOf(find, StringComparison.Ordinal)
            If idx < 0 Then Return

            Dim before As String = pt.Substring(0, idx)
            Dim mid As String = pt.Substring(idx, find.Length)
            Dim after As String = pt.Substring(idx + find.Length)

            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
            para.RemoveAllChildren()

            If pPr IsNot Nothing Then
                para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
            End If

            If before.Length > 0 Then
                para.AppendChild(MakeRun(before))
            End If

            Select Case op
                Case "replace"
                    If asMarkup Then
                        If mid.Length > 0 Then
                            Dim del As New W.DeletedRun() With {
                                .Id = NextChangeId(body).ToString(),
                                .Author = author,
                                .Date = DateTime.UtcNow
                            }
                            del.AppendChild(MakeRun(mid, deletedText:=True))
                            para.AppendChild(del)
                        End If

                        If newText.Length > 0 Then
                            Dim ins As New W.InsertedRun() With {
                                .Id = NextChangeId(body).ToString(),
                                .Author = author,
                                .Date = DateTime.UtcNow
                            }
                            ins.AppendChild(MakeRun(newText))
                            para.AppendChild(ins)
                        End If
                    Else
                        If newText.Length > 0 Then
                            para.AppendChild(MakeRun(newText))
                        End If
                    End If

                Case "insert_before"
                    If asMarkup AndAlso newText.Length > 0 Then
                        Dim ins As New W.InsertedRun() With {
                            .Id = NextChangeId(body).ToString(),
                            .Author = author,
                            .Date = DateTime.UtcNow
                        }
                        ins.AppendChild(MakeRun(newText))
                        para.AppendChild(ins)
                    ElseIf newText.Length > 0 Then
                        para.AppendChild(MakeRun(newText))
                    End If

                    If mid.Length > 0 Then
                        para.AppendChild(MakeRun(mid))
                    End If

                Case "insert_after"
                    If mid.Length > 0 Then
                        para.AppendChild(MakeRun(mid))
                    End If

                    If asMarkup AndAlso newText.Length > 0 Then
                        Dim ins As New W.InsertedRun() With {
                            .Id = NextChangeId(body).ToString(),
                            .Author = author,
                            .Date = DateTime.UtcNow
                        }
                        ins.AppendChild(MakeRun(newText))
                        para.AppendChild(ins)
                    ElseIf newText.Length > 0 Then
                        para.AppendChild(MakeRun(newText))
                    End If

                Case Else
                    If mid.Length > 0 Then
                        para.AppendChild(MakeRun(mid))
                    End If
            End Select

            If after.Length > 0 Then
                para.AppendChild(MakeRun(after))
            End If
        End Sub

        ' --------------------------------------------------------------- comments

        Private Shared Function ExecuteCommentAdd(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim find As String = GetStr(args, "find")
            Dim text As String = GetStr(args, "text")
            Dim author As String = If(GetStr(args, "author"), "Red Ink")
            Dim initials As String = If(GetStr(args, "initials"), "RI")

            If String.IsNullOrWhiteSpace(find) Then Return Err_("missing_find", "find is required.")
            If String.IsNullOrWhiteSpace(text) Then Return Err_("missing_text", "text is required.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim main As MainDocumentPart = doc.MainDocumentPart
                Dim commentsPart As WordprocessingCommentsPart = main.WordprocessingCommentsPart

                If commentsPart Is Nothing Then
                    commentsPart = main.AddNewPart(Of WordprocessingCommentsPart)()
                    commentsPart.Comments = New W.Comments()
                End If

                Dim newId As Integer = NextCommentId(commentsPart.Comments)
                Dim cmt As New W.Comment() With {
                    .Id = newId.ToString(),
                    .Author = author,
                    .Initials = initials,
                    .Date = DateTime.UtcNow
                }

                cmt.AppendChild(
                    New W.Paragraph(
                        New W.Run(
                            New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve}
                        )
                    )
                )

                commentsPart.Comments.AppendChild(cmt)

                Dim attached As Boolean = False

                For Each para As W.Paragraph In main.Document.Body.Descendants(Of W.Paragraph)().ToList()
                    Dim pt As String = GetParagraphText(para)
                    Dim idx As Integer = pt.IndexOf(find, StringComparison.Ordinal)
                    If idx < 0 Then Continue For

                    AttachCommentRangeToParagraph(para, find, newId)
                    attached = True
                    Exit For
                Next

                If Not attached Then
                    Return Err_("no_match", "No W.Paragraph contained the 'find' text.")
                End If

                commentsPart.Comments.Save()
                main.Document.Save()

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .comment_id = newId
                })
            End Using
        End Function

        Private Shared Function ExecuteCommentList(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Read)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=False)
                Dim part As WordprocessingCommentsPart = doc.MainDocumentPart.WordprocessingCommentsPart
                Dim list As New List(Of Object)()

                If part IsNot Nothing AndAlso part.Comments IsNot Nothing Then
                    For Each c As W.Comment In part.Comments.Elements(Of W.Comment)()
                        list.Add(New With {
                            Key .id = If(c.Id Is Nothing, Nothing, c.Id.Value),
                            Key .author = If(c.Author Is Nothing, Nothing, c.Author.Value),
                            Key .initials = If(c.Initials Is Nothing, Nothing, c.Initials.Value),
                            Key .date = If(c.Date Is Nothing, CType(Nothing, Nullable(Of DateTime)), c.Date.Value),
                            Key .text = c.InnerText
                        })
                    Next
                End If

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .comments = list
                })
            End Using
        End Function

        Private Shared Function ExecuteCommentRemove(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim id As String = GetStr(args, "id")
            If String.IsNullOrWhiteSpace(id) Then Return Err_("missing_id", "id is required.")

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim main As MainDocumentPart = doc.MainDocumentPart
                Dim part As WordprocessingCommentsPart = main.WordprocessingCommentsPart

                If part Is Nothing OrElse part.Comments Is Nothing Then
                    Return Err_("not_found", "No comments part.")
                End If

                Dim cmt As W.Comment = part.Comments.Elements(Of W.Comment)().
                    FirstOrDefault(Function(c) c.Id IsNot Nothing AndAlso c.Id.Value = id)

                If cmt Is Nothing Then
                    Return Err_("not_found", "No comment with id '" & id & "'.")
                End If

                cmt.Remove()

                For Each n As W.CommentRangeStart In main.Document.Body.Descendants(Of W.CommentRangeStart)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    n.Remove()
                Next

                For Each n As W.CommentRangeEnd In main.Document.Body.Descendants(Of W.CommentRangeEnd)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    n.Remove()
                Next

                For Each n As W.CommentReference In main.Document.Body.Descendants(Of W.CommentReference)().
                    Where(Function(x) x.Id IsNot Nothing AndAlso x.Id.Value = id).ToList()
                    If n.Parent IsNot Nothing Then
                        n.Parent.Remove()
                    End If
                Next

                part.Comments.Save()
                main.Document.Save()

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .removed_id = id
                })
            End Using
        End Function

        ' --------------------------------------------------------------- format

        Private Shared Function ExecuteFormat(args As IDictionary(Of String, Object)) As String
            Dim p As String = PathPolicy.Resolve(GetStr(args, "path"), PathAccess.Write)
            If Not File.Exists(p) Then Return Err_("not_found", "File not found.")

            Dim find As String = GetStr(args, "find")
            Dim styleId As String = GetStr(args, "style") ' W.Paragraph style id (e.g. "Heading1")
            Dim bold As Boolean? = GetNullableBool(args, "bold")
            Dim italic As Boolean? = GetNullableBool(args, "italic")
            Dim underline As Boolean? = GetNullableBool(args, "underline")
            Dim sizePt As Integer = GetInt(args, "size", 0)
            Dim color As String = GetStr(args, "color") ' "RRGGBB" hex
            Dim align As String = GetStr(args, "align").ToLowerInvariant() ' left|center|right|justify

            If String.IsNullOrWhiteSpace(find) Then
                Return Err_("missing_find", "find is required (use empty W.Paragraph match by passing the W.Paragraph's text).")
            End If

            Using doc As WordprocessingDocument = WordprocessingDocument.Open(p, isEditable:=True)
                Dim mutated As Integer = 0

                For Each para As W.Paragraph In doc.MainDocumentPart.Document.Body.Descendants(Of W.Paragraph)().ToList()
                    Dim pt As String = GetParagraphText(para)
                    If pt.IndexOf(find, StringComparison.Ordinal) < 0 Then Continue For

                    Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
                    If pPr Is Nothing Then
                        pPr = New W.ParagraphProperties()
                        para.InsertAt(Of W.ParagraphProperties)(pPr, 0)
                    End If

                    If Not String.IsNullOrWhiteSpace(styleId) Then
                        Dim psid As W.ParagraphStyleId = pPr.Elements(Of W.ParagraphStyleId)().FirstOrDefault()
                        If psid Is Nothing Then
                            pPr.AppendChild(New W.ParagraphStyleId() With {.Val = styleId})
                        Else
                            psid.Val = styleId
                        End If
                    End If

                    Select Case align
                        Case "left", "center", "right", "justify"
                            Dim ja As W.Justification = pPr.Elements(Of W.Justification)().FirstOrDefault()
                            If ja Is Nothing Then
                                ja = New W.Justification()
                                pPr.AppendChild(ja)
                            End If
                            ja.Val = AlignmentFromString(align)
                    End Select

                    For Each run As W.Run In para.Elements(Of W.Run)()
                        Dim rPr As W.RunProperties = run.RunProperties
                        If rPr Is Nothing Then
                            rPr = New W.RunProperties()
                            run.InsertAt(Of W.RunProperties)(rPr, 0)
                        End If

                        If bold.HasValue Then
                            SetBool(Of W.Bold)(rPr, bold.Value)
                        End If

                        If italic.HasValue Then
                            SetBool(Of W.Italic)(rPr, italic.Value)
                        End If

                        If underline.HasValue Then
                            Dim u As W.Underline = rPr.Elements(Of W.Underline)().FirstOrDefault()
                            If underline.Value Then
                                If u Is Nothing Then
                                    rPr.AppendChild(New W.Underline() With {.Val = W.UnderlineValues.Single})
                                End If
                            ElseIf u IsNot Nothing Then
                                u.Remove()
                            End If
                        End If

                        If sizePt > 0 Then
                            Dim sz As W.FontSize = rPr.Elements(Of W.FontSize)().FirstOrDefault()
                            If sz Is Nothing Then
                                sz = New W.FontSize()
                                rPr.AppendChild(sz)
                            End If
                            sz.Val = (sizePt * 2).ToString() ' half-points
                        End If

                        If Not String.IsNullOrWhiteSpace(color) Then
                            Dim cc As W.Color = rPr.Elements(Of W.Color)().FirstOrDefault()
                            If cc Is Nothing Then
                                cc = New W.Color()
                                rPr.AppendChild(cc)
                            End If
                            cc.Val = color.TrimStart("#"c)
                        End If
                    Next

                    mutated += 1
                Next

                If mutated = 0 Then
                    Return Err_("no_match", "No W.Paragraph contained the 'find' text.")
                End If

                doc.MainDocumentPart.Document.Save()

                Return JsonConvert.SerializeObject(New With {
                    Key .path = p,
                    Key .paragraphs_changed = mutated
                })
            End Using
        End Function

        ' --------------------------------------------------------------- template / save_as

        Private Shared Function ExecuteApplyTemplate(args As IDictionary(Of String, Object)) As String
            Dim skillName As String = GetStr(args, "skill")
            Dim relTemplate As String = GetStr(args, "template") ' relative to the skill's references/
            Dim outName As String = If(GetStr(args, "output_name"), "from_template.docx")
            Dim subsToken As JToken = Nothing

            If args IsNot Nothing AndAlso args.ContainsKey("substitutions") Then
                Try
                    subsToken = JToken.FromObject(args("substitutions"))
                Catch
                End Try
            End If

            If String.IsNullOrWhiteSpace(skillName) OrElse String.IsNullOrWhiteSpace(relTemplate) Then
                Return Err_("missing_args", "Both 'skill' and 'template' are required.")
            End If

            Dim sk = AgentResources.FindSkill(skillName)
            If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.ReferencesDir) OrElse Not Directory.Exists(sk.ReferencesDir) Then
                Return Err_("skill_not_found", "Skill or references/ directory not found.")
            End If

            Dim src As String = Path.GetFullPath(Path.Combine(sk.ReferencesDir, relTemplate))
            If Not src.StartsWith(Path.GetFullPath(sk.ReferencesDir), StringComparison.OrdinalIgnoreCase) Then
                Return Err_("path_escape", "Template path escapes references/.")
            End If

            If Not File.Exists(src) Then Return Err_("not_found", "Template not found: " & src)

            Dim dst As String = PathPolicy.NewWritablePath(outName)
            File.Copy(src, dst, overwrite:=False)

            Dim placeholders As New Dictionary(Of String, String)(StringComparer.Ordinal)
            If subsToken IsNot Nothing AndAlso subsToken.Type = JTokenType.Object Then
                For Each kv As JProperty In CType(subsToken, JObject).Properties()
                    placeholders("{{" & kv.Name & "}}") = If(kv.Value Is Nothing, "", kv.Value.ToString())
                Next
            End If

            If placeholders.Count > 0 Then
                Using doc As WordprocessingDocument = WordprocessingDocument.Open(dst, isEditable:=True)
                    For Each para As W.Paragraph In doc.MainDocumentPart.Document.Body.Descendants(Of W.Paragraph)().ToList()
                        Dim pt As String = GetParagraphText(para)
                        Dim changed As Boolean = False
                        Dim newText As String = pt

                        For Each kv As KeyValuePair(Of String, String) In placeholders
                            If newText.IndexOf(kv.Key, StringComparison.Ordinal) >= 0 Then
                                newText = newText.Replace(kv.Key, kv.Value)
                                changed = True
                            End If
                        Next

                        If changed Then
                            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()
                            para.RemoveAllChildren()

                            If pPr IsNot Nothing Then
                                para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
                            End If

                            para.AppendChild(MakeRun(newText))
                        End If
                    Next

                    doc.MainDocumentPart.Document.Save()
                End Using
            End If

            Return JsonConvert.SerializeObject(New With {
                Key .path = dst,
                Key .template = src,
                Key .substitutions = placeholders.Count
            })
        End Function

        Private Shared Function ExecuteSaveAs(args As IDictionary(Of String, Object)) As String
            Dim src As String = PathPolicy.Resolve(GetStr(args, "source"), PathAccess.Read)
            If Not File.Exists(src) Then Return Err_("not_found", "Source not found.")

            Dim outName As String = If(GetStr(args, "output_name"), Path.GetFileName(src))
            Dim dst As String = PathPolicy.NewWritablePath(outName)
            File.Copy(src, dst, overwrite:=False)

            Return JsonConvert.SerializeObject(New With {
                Key .source = src,
                Key .path = dst
            })
        End Function

        ' --------------------------------------------------------------- OOXML helpers

        Private Class ParagraphRow
            Public Index As Integer
            Public Text As String
        End Class

        Private Shared Function ExtractParagraphs(doc As WordprocessingDocument) As List(Of ParagraphRow)
            Dim output As New List(Of ParagraphRow)()
            Dim i As Integer = 0

            For Each p As W.Paragraph In doc.MainDocumentPart.Document.Body.Descendants(Of W.Paragraph)()
                output.Add(New ParagraphRow With {
                    .Index = i,
                    .Text = GetParagraphText(p)
                })
                i += 1
            Next

            Return output
        End Function

        Private Shared Function GetParagraphText(p As W.Paragraph) As String
            Dim sb As New StringBuilder()

            For Each t As W.Text In p.Descendants(Of W.Text)()
                sb.Append(t.Text)
            Next

            Return sb.ToString()
        End Function

        Private Shared Function MakeRun(text As String, Optional deletedText As Boolean = False) As W.Run
            Dim r As New W.Run()

            If deletedText Then
                Dim dt As New W.DeletedText(text) With {.Space = SpaceProcessingModeValues.Preserve}
                r.AppendChild(dt)
            Else
                r.AppendChild(New W.Text(text) With {.Space = SpaceProcessingModeValues.Preserve})
            End If

            Return r
        End Function

        Private Shared Function NextChangeId(body As W.Body) As Integer
            Dim maxId As Integer = 0

            For Each n As W.InsertedRun In body.Descendants(Of W.InsertedRun)()
                Dim v As Integer
                If n.Id IsNot Nothing AndAlso Integer.TryParse(n.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            For Each n As W.DeletedRun In body.Descendants(Of W.DeletedRun)()
                Dim v As Integer
                If n.Id IsNot Nothing AndAlso Integer.TryParse(n.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            Return maxId + 1
        End Function

        Private Shared Function NextCommentId(comments As W.Comments) As Integer
            Dim maxId As Integer = 0

            For Each c As W.Comment In comments.Elements(Of W.Comment)()
                Dim v As Integer
                If c.Id IsNot Nothing AndAlso Integer.TryParse(c.Id.Value, v) AndAlso v > maxId Then
                    maxId = v
                End If
            Next

            Return maxId + 1
        End Function

        Private Shared Sub AttachCommentRangeToParagraph(para As W.Paragraph, find As String, commentId As Integer)
            Dim pt As String = GetParagraphText(para)
            Dim idx As Integer = pt.IndexOf(find, StringComparison.Ordinal)
            If idx < 0 Then Return

            Dim before As String = pt.Substring(0, idx)
            Dim mid As String = pt.Substring(idx, find.Length)
            Dim after As String = pt.Substring(idx + find.Length)
            Dim pPr As W.ParagraphProperties = para.Elements(Of W.ParagraphProperties)().FirstOrDefault()

            para.RemoveAllChildren()

            If pPr IsNot Nothing Then
                para.AppendChild(CType(pPr.CloneNode(True), W.ParagraphProperties))
            End If

            If before.Length > 0 Then
                para.AppendChild(MakeRun(before))
            End If

            para.AppendChild(New W.CommentRangeStart() With {.Id = commentId.ToString()})

            If mid.Length > 0 Then
                para.AppendChild(MakeRun(mid))
            End If

            para.AppendChild(New W.CommentRangeEnd() With {.Id = commentId.ToString()})

            Dim refRun As New W.Run()
            refRun.AppendChild(New W.CommentReference() With {.Id = commentId.ToString()})
            para.AppendChild(refRun)

            If after.Length > 0 Then
                para.AppendChild(MakeRun(after))
            End If
        End Sub

        Private Shared Function AlignmentFromString(s As String) As W.JustificationValues
            Select Case s
                Case "center"
                    Return W.JustificationValues.Center
                Case "right"
                    Return W.JustificationValues.Right
                Case "justify"
                    Return W.JustificationValues.Both
                Case Else
                    Return W.JustificationValues.Left
            End Select
        End Function

        Private Shared Sub SetBool(Of T As {OpenXmlElement, New})(rPr As W.RunProperties, value As Boolean)
            Dim existing As T = rPr.Elements(Of T)().FirstOrDefault()

            If value Then
                If existing Is Nothing Then
                    rPr.AppendChild(New T())
                End If
            ElseIf existing IsNot Nothing Then
                existing.Remove()
            End If
        End Sub

        Private Shared Function BuildHit(paragraphIndex As Integer, paraText As String, index As Integer, length As Integer, match As String) As Object
            Dim winStart As Integer = System.Math.Max(0, index - 40)
            Dim winEnd As Integer = System.Math.Min(paraText.Length, index + length + 40)
            Dim ctx As String = paraText.Substring(winStart, winEnd - winStart)

            Return New With {
                Key .paragraph_index = paragraphIndex,
                Key .index_in_paragraph = index,
                Key .length = length,
                Key .match = match,
                Key .context = ctx
            }
        End Function

        Private Shared Function Err_(code As String, message As String) As String
            Return JsonConvert.SerializeObject(New With {
                Key .error = code,
                Key .message = message
            })
        End Function

        ' --------------------------------------------------------------- argument helpers

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""

            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, defaultValue As Integer) As Integer
            If args Is Nothing Then Return defaultValue

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return defaultValue

            Try
                Return System.Convert.ToInt32(v)
            Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return defaultValue
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, defaultValue As Boolean) As Boolean
            Dim nb As Boolean? = GetNullableBool(args, name)
            If nb.HasValue Then Return nb.Value
            Return defaultValue
        End Function

        Private Shared Function GetNullableBool(args As IDictionary(Of String, Object), name As String) As Boolean?
            If args Is Nothing Then Return Nothing

            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return Nothing

            Try
                Return System.Convert.ToBoolean(v)
            Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes"
                        Return True
                    Case "false", "0", "no"
                        Return False
                    Case Else
                        Return Nothing
                End Select
            End Try
        End Function

        ' --------------------------------------------------------------- factories

        Private Shared Function BuildExtract() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolExtract,
                .Tool = True,
                .ToolPriority = 880,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (extract text)",
                .ToolDefinition = "{""name"":""" & ToolExtract & """,""description"":""Extract plain text from a .docx file. Returns paragraphs joined by newlines."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""max_chars"":{""type"":""integer"",""description"":""Optional cap on returned text length.""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolExtract & ": Extract plain text from a .docx file."
            }
        End Function

        Private Shared Function BuildSearch() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSearch,
                .Tool = True,
                .ToolPriority = 881,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (search)",
                .ToolDefinition = "{""name"":""" & ToolSearch & """,""description"":""Search a .docx for a substring or regex. Returns W.Paragraph index and a small context window per hit."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""query"":{""type"":""string""},""regex"":{""type"":""boolean""},""ignore_case"":{""type"":""boolean""},""max_hits"":{""type"":""integer""}},""required"":[""path"",""query""]}}",
                .ToolInstructionsPrompt = ToolSearch & ": Find text inside a .docx file."
            }
        End Function

        Private Shared Function BuildWrite() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolWrite,
                .Tool = True,
                .ToolPriority = 882,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (write, no markup)",
                .ToolDefinition = "{""name"":""" & ToolWrite & """,""description"":""Modify text in a .docx WITHOUT tracked changes. Ops: replace | insert_before | insert_after | append (no 'find' required for append)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append""]},""find"":{""type"":""string""},""text"":{""type"":""string""},""only_first"":{""type"":""boolean"",""description"":""Default true.""}},""required"":[""path"",""text""]}}",
                .ToolInstructionsPrompt = ToolWrite & ": Edit a .docx without revision marks."
            }
        End Function

        Private Shared Function BuildMarkup() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolMarkup,
                .Tool = True,
                .ToolPriority = 883,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (markup / tracked changes)",
                .ToolDefinition = "{""name"":""" & ToolMarkup & """,""description"":""Modify text in a .docx using tracked changes (Word revision marks). Same ops as word_write."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""op"":{""type"":""string"",""enum"":[""replace"",""insert_before"",""insert_after"",""append""]},""find"":{""type"":""string""},""text"":{""type"":""string""},""author"":{""type"":""string""},""only_first"":{""type"":""boolean""}},""required"":[""path"",""text""]}}",
                .ToolInstructionsPrompt = ToolMarkup & ": Edit a .docx with revision marks (tracked changes)."
            }
        End Function

        Private Shared Function BuildCommentAdd() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentAdd,
                .Tool = True,
                .ToolPriority = 884,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comment add)",
                .ToolDefinition = "{""name"":""" & ToolCommentAdd & """,""description"":""Attach a Word comment to the first W.Paragraph containing 'find'. Returns the new comment id."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""find"":{""type"":""string""},""text"":{""type"":""string""},""author"":{""type"":""string""},""initials"":{""type"":""string""}},""required"":[""path"",""find"",""text""]}}",
                .ToolInstructionsPrompt = ToolCommentAdd & ": Add a Word bubble comment to a matched span."
            }
        End Function

        Private Shared Function BuildCommentList() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentList,
                .Tool = True,
                .ToolPriority = 885,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comments list)",
                .ToolDefinition = "{""name"":""" & ToolCommentList & """,""description"":""List all comments in a .docx with id, author, initials, date and inner text."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""}},""required"":[""path""]}}",
                .ToolInstructionsPrompt = ToolCommentList & ": List Word comments in a .docx."
            }
        End Function

        Private Shared Function BuildCommentRemove() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolCommentRemove,
                .Tool = True,
                .ToolPriority = 886,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (comment remove)",
                .ToolDefinition = "{""name"":""" & ToolCommentRemove & """,""description"":""Remove a Word comment by id (also strips its range markers and reference)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""id"":{""type"":""string""}},""required"":[""path"",""id""]}}",
                .ToolInstructionsPrompt = ToolCommentRemove & ": Remove a Word comment by id."
            }
        End Function

        Private Shared Function BuildFormat() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolFormat,
                .Tool = True,
                .ToolPriority = 887,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (format)",
                .ToolDefinition = "{""name"":""" & ToolFormat & """,""description"":""Apply W.Paragraph style and/or W.Run formatting to every W.Paragraph containing 'find'. Available: style (Word style id, e.g. 'Heading1'), bold, italic, underline, size (pt), color (RRGGBB), align (left|center|right|justify)."",""parameters"":{""type"":""object"",""properties"":{""path"":{""type"":""string""},""find"":{""type"":""string""},""style"":{""type"":""string""},""bold"":{""type"":""boolean""},""italic"":{""type"":""boolean""},""underline"":{""type"":""boolean""},""size"":{""type"":""integer""},""color"":{""type"":""string""},""align"":{""type"":""string"",""enum"":[""left"",""center"",""right"",""justify""]}},""required"":[""path"",""find""]}}",
                .ToolInstructionsPrompt = ToolFormat & ": Apply W.Paragraph/W.Run formatting (style, bold/italic/underline, size, color, alignment)."
            }
        End Function

        Private Shared Function BuildApplyTemplate() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolApplyTemplate,
                .Tool = True,
                .ToolPriority = 888,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (apply template)",
                .ToolDefinition = "{""name"":""" & ToolApplyTemplate & """,""description"":""Clone a .docx template from a skill's references/ directory to a new file in the workspace (or Desktop) and substitute {{placeholders}} from the 'substitutions' object."",""parameters"":{""type"":""object"",""properties"":{""skill"":{""type"":""string"",""description"":""Skill name.""},""template"":{""type"":""string"",""description"":""Path relative to the skill's references/ directory.""},""output_name"":{""type"":""string"",""description"":""Suggested output filename (default 'from_template.docx').""},""substitutions"":{""type"":""object"",""description"":""Object of {placeholderName: value}; each key K becomes the literal '{{K}}' in the template.""}},""required"":[""skill"",""template""]}}",
                .ToolInstructionsPrompt = ToolApplyTemplate & ": Instantiate a Word template from a skill's references/ directory."
            }
        End Function

        Private Shared Function BuildSaveAs() As ModelConfig
            Return New ModelConfig() With {
                .ToolName = ToolSaveAs,
                .Tool = True,
                .ToolPriority = 889,
                .ToolErrorHandling = "skip",
                .ModelDescription = "Word (save as)",
                .ToolDefinition = "{""name"":""" & ToolSaveAs & """,""description"":""Copy a .docx to a new path inside the writable root (workspace or Desktop)."",""parameters"":{""type"":""object"",""properties"":{""source"":{""type"":""string""},""output_name"":{""type"":""string""}},""required"":[""source""]}}",
                .ToolInstructionsPrompt = ToolSaveAs & ": Copy a .docx to a new path in the writable root."
            }
        End Function

    End Class

End Namespace