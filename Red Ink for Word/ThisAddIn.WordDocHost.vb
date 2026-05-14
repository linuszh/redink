' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: WordDocHost.vb
' PURPOSE
'   Services for agents to interact with Word documents: list, read, search, insert, format, comment.
'
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.Text
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports SharedLibrary
Imports SharedLibrary.Agents
Imports Word = Microsoft.Office.Interop.Word

Partial Public Class ThisAddIn
    Implements SharedLibrary.Agents.IWordDocumentHost

    Public Function HasActiveDocument() As Boolean Implements IWordDocumentHost.HasActiveDocument
        Try
            Return Globals.ThisAddIn.Application IsNot Nothing AndAlso
                   Globals.ThisAddIn.Application.Documents IsNot Nothing AndAlso
                   Globals.ThisAddIn.Application.Documents.Count > 0
        Catch
            Return False
        End Try
    End Function

    Public Function ListOpenDocuments() As List(Of OpenDocInfo) Implements IWordDocumentHost.ListOpenDocuments
        Dim list As New List(Of OpenDocInfo)
        Try
            Dim app = Globals.ThisAddIn.Application
            Dim active = TryGetActive(app)
            For Each d As Word.Document In app.Documents
                list.Add(New OpenDocInfo With {
                    .Name = d.Name,
                    .Path = TryGetFullName(d),
                    .IsActive = active IsNot Nothing AndAlso d Is active,
                    .IsReadOnly = d.ReadOnly,
                    .IsSaved = d.Saved
                })
            Next
        Catch
        End Try
        Return list
    End Function

    Public Function GetActiveDocument() As OpenDocInfo Implements IWordDocumentHost.GetActiveDocument
        Dim app = Globals.ThisAddIn.Application
        Dim d = TryGetActive(app)
        If d Is Nothing Then Return Nothing
        Return New OpenDocInfo With {
            .Name = d.Name, .Path = TryGetFullName(d), .IsActive = True,
            .IsReadOnly = d.ReadOnly, .IsSaved = d.Saved
        }
    End Function

    Public Function ExtractText(target As String, maxChars As Integer) As String Implements IWordDocumentHost.ExtractText
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return ""
        Dim text = d.Content.Text
        If maxChars > 0 AndAlso text.Length > maxChars Then text = text.Substring(0, maxChars)
        Return text
    End Function

    Public Function SearchJson(target As String, query As String, useRegex As Boolean,
                                ignoreCase As Boolean, maxHits As Integer) As String _
                                Implements IWordDocumentHost.SearchJson
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return Err_("not_found", "Document not found.")
        If maxHits < 1 Then maxHits = 50
        If maxHits > 500 Then maxHits = 500
        Dim text = d.Content.Text
        Dim hits As New List(Of Object)
        If useRegex Then
            Dim opt As RegexOptions = RegexOptions.CultureInvariant
            If ignoreCase Then opt = opt Or RegexOptions.IgnoreCase
            Dim rx As New Regex(query, opt, TimeSpan.FromSeconds(2))
            For Each m As Match In rx.Matches(text)
                If hits.Count >= maxHits Then Exit For
                hits.Add(BuildHit(text, m.Index, m.Length, m.Value))
            Next
        Else
            Dim cmp = If(ignoreCase, StringComparison.OrdinalIgnoreCase, StringComparison.Ordinal)
            Dim idx = 0
            While idx < text.Length
                Dim f = text.IndexOf(query, idx, cmp)
                If f < 0 Then Exit While
                hits.Add(BuildHit(text, f, query.Length, text.Substring(f, query.Length)))
                If hits.Count >= maxHits Then Exit While
                idx = f + Math.Max(1, query.Length)
            End While
        End If
        Return JsonConvert.SerializeObject(New With {Key .target = d.Name, Key .total = hits.Count, Key .hits = hits})
    End Function

    Public Function ListComments(target As String) As List(Of WordDocComment) Implements IWordDocumentHost.ListComments
        Dim out As New List(Of WordDocComment)
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return out
        For Each c As Word.Comment In d.Comments
            Dim anchor As String = ""
            Try : anchor = c.Scope.Text : Catch : End Try
            out.Add(New WordDocComment With {
                .Id = c.Index.ToString(),
                .Author = c.Author,
                .Initials = c.Initial,
                .Date_ = SafeDate(c.Date),
                .Text = c.Range.Text,
                .AnchorText = anchor
            })
        Next
        Return out
    End Function

    Public Function InsertTextJson(target As String, text As String, location As String) As String _
                                    Implements IWordDocumentHost.InsertTextJson
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return Err_("not_found", "Document not found.")
        Dim rng As Word.Range
        Select Case If(location, "end").ToLowerInvariant()
            Case "start" : rng = d.Range(0, 0)
            Case "cursor"
                Try : rng = Globals.ThisAddIn.Application.Selection.Range : Catch : rng = d.Content : rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                End Try
            Case Else
                rng = d.Content : rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
        End Select
        rng.InsertAfter(text)
        Return JsonConvert.SerializeObject(New With {Key .target = d.Name, Key .inserted = text.Length})
    End Function

    Public Function ReplaceJson(target As String, find As String, replacement As String,
                                  onlyFirst As Boolean) As String Implements IWordDocumentHost.ReplaceJson
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return Err_("not_found", "Document not found.")
        If String.IsNullOrEmpty(find) Then Return Err_("missing_find", "find is required.")
        Dim rng = d.Content
        rng.Find.ClearFormatting()
        rng.Find.Replacement.ClearFormatting()
        Dim wrap As Word.WdFindWrap = Word.WdFindWrap.wdFindStop
        Dim replaceMode As Word.WdReplace = If(onlyFirst, Word.WdReplace.wdReplaceOne, Word.WdReplace.wdReplaceAll)
        rng.Find.Execute(FindText:=find, Replace:=replaceMode, ReplaceWith:=replacement,
                         MatchCase:=False, MatchWholeWord:=False, MatchWildcards:=False,
                         Forward:=True, Wrap:=wrap)
        Return JsonConvert.SerializeObject(New With {Key .target = d.Name, Key .replaced = True})
    End Function

    Public Function AddCommentJson(target As String, find As String, text As String,
                                     author As String, initials As String) As String _
                                     Implements IWordDocumentHost.AddCommentJson
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return Err_("not_found", "Document not found.")
        If String.IsNullOrEmpty(find) Then Return Err_("missing_find", "find is required.")
        Dim rng = d.Content
        rng.Find.ClearFormatting()
        If rng.Find.Execute(FindText:=find, Forward:=True, Wrap:=Word.WdFindWrap.wdFindStop) Then
            Dim cmt = d.Comments.Add(rng, text)
            If Not String.IsNullOrEmpty(author) Then cmt.Author = author
            If Not String.IsNullOrEmpty(initials) Then cmt.Initial = initials
            Return JsonConvert.SerializeObject(New With {Key .target = d.Name, Key .comment_id = cmt.Index})
        End If
        Return Err_("no_match", "No match for 'find'.")
    End Function

    Public Function FormatJson(target As String, find As String, styleId As String,
                                 bold As Boolean?, italic As Boolean?, underline As Boolean?,
                                 sizePt As Integer, color As String, align As String) As String _
                                 Implements IWordDocumentHost.FormatJson
        Dim d = ResolveDoc(target)
        If d Is Nothing Then Return Err_("not_found", "Document not found.")
        If String.IsNullOrEmpty(find) Then Return Err_("missing_find", "find is required.")
        Dim rng = d.Content
        rng.Find.ClearFormatting()
        If Not rng.Find.Execute(FindText:=find, Forward:=True, Wrap:=Word.WdFindWrap.wdFindStop) Then
            Return Err_("no_match", "No match for 'find'.")
        End If
        Try
            If Not String.IsNullOrWhiteSpace(styleId) Then
                Try
                    Dim styleObj As Object = d.Styles(styleId)
                    rng.Style = styleObj
                Catch
                End Try
            End If
        Catch
        End Try
        If bold.HasValue Then rng.Font.Bold = If(bold.Value, -1, 0)
        If italic.HasValue Then rng.Font.Italic = If(italic.Value, -1, 0)
        If underline.HasValue Then rng.Font.Underline = If(underline.Value, Word.WdUnderline.wdUnderlineSingle, Word.WdUnderline.wdUnderlineNone)
        If sizePt > 0 Then rng.Font.Size = sizePt
        If Not String.IsNullOrWhiteSpace(color) Then
            Try
                Dim hex = color.TrimStart("#"c)
                Dim n = Convert.ToInt32(hex, 16)
                ' Word uses BGR
                Dim r = (n >> 16) And &HFF
                Dim g = (n >> 8) And &HFF
                Dim b = n And &HFF
                rng.Font.Color = CType((b << 16) Or (g << 8) Or r, Word.WdColor)
            Catch
            End Try
        End If
        Select Case If(align, "").ToLowerInvariant()
            Case "left" : rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
            Case "center" : rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            Case "right" : rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
            Case "justify" : rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify
        End Select
        Return JsonConvert.SerializeObject(New With {Key .target = d.Name, Key .formatted = True})
    End Function

    ' --------------------------------------------------------------- helpers

    Private Function ResolveDoc(target As String) As Word.Document
        Try
            Dim app = Globals.ThisAddIn.Application
            If String.IsNullOrWhiteSpace(target) Then Return TryGetActive(app)
            For Each d As Word.Document In app.Documents
                If String.Equals(d.Name, target, StringComparison.OrdinalIgnoreCase) Then Return d
                Dim full = TryGetFullName(d)
                If Not String.IsNullOrWhiteSpace(full) AndAlso String.Equals(full, target, StringComparison.OrdinalIgnoreCase) Then Return d
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function TryGetActive(app As Word.Application) As Word.Document
        Try : Return app.ActiveDocument : Catch : Return Nothing : End Try
    End Function

    Private Shared Function TryGetFullName(d As Word.Document) As String
        Try : Return d.FullName : Catch : Return "" : End Try
    End Function

    Private Shared Function BuildHit(text As String, index As Integer, length As Integer, match As String) As Object
        Dim winStart = Math.Max(0, index - 40)
        Dim winEnd = Math.Min(text.Length, index + length + 40)
        Dim ctx = text.Substring(winStart, winEnd - winStart).Replace(vbCr, " ").Replace(vbLf, " ")
        Return New With {Key .index = index, Key .length = length, Key .match = match, Key .context = ctx}
    End Function

    Private Shared Function SafeDate(d As Object) As DateTime?
        Try : Return CType(d, DateTime) : Catch : Return Nothing : End Try
    End Function

    Private Shared Function Err_(code As String, message As String) As String
        Return JsonConvert.SerializeObject(New With {Key .error = code, Key .message = message})
    End Function

End Class