' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.M365.vb (Word host adapter)
' Purpose: Bridges Word's host-local ToolCall / ToolResponse / ToolExecutionContext
'          to the host-neutral SharedLibrary.M365ToolService.
'
'   Wire-up:
'      ExecuteToolCall, after the InternalKnowledge dispatch arm:
'          ElseIf SharedLibrary.M365ToolService.IsM365ToolName(toolCall.ToolName) Then
'              response = Await ExecuteInternalM365Tool(toolCall, context)
'              ToolingFileLogger.LogRawResponseStub(
'                  $"Internal tool ({toolCall.ToolName})", response.Response)
'
'      GetAvailableTools, after GetInternalKnowledgeTools:
'          tools.AddRange(SharedLibrary.M365ToolService.GetTools(_context, InternalToolSuffix))
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Adapter that hands a Word ToolCall over to the shared M365ToolService and
    ''' converts the neutral result back into a host-local ToolResponse.
    ''' </summary>
    Private Async Function ExecuteInternalM365Tool(toolCall As ToolCall,
                                                   context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim resp As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Dim r = Await M365ToolService.ExecuteAsync(
            _context,
            toolCall.ToolName,
            toolCall.Arguments,
            log:=Sub(s) context.Log(s),
            ct:=Nothing)

        resp.Success = r.Success
        resp.Response = r.Response
        resp.ErrorMessage = r.ErrorMessage
        Return resp
    End Function

    Private Class ToolSourceLink
        Public Property Url As String
        Public Property Title As String
        Public Property Source As String
    End Class

    Private Function AppendM365SourcesFooter(finalAnswer As String,
                                             toolResponses As List(Of ToolResponse)) As String
        Dim answer As String = If(finalAnswer, "").Trim()
        Dim links As List(Of ToolSourceLink) = ExtractM365SourceLinks(toolResponses, answer)

        If links.Count = 0 Then
            Return answer
        End If

        Dim sb As New StringBuilder()

        If answer.Length > 0 Then
            sb.AppendLine(answer)
            sb.AppendLine()
        End If

        sb.AppendLine("### Sources")

        For Each link In links
            Dim label As String = BuildSourceLinkLabel(link)
            sb.AppendLine($"- [{EscapeMarkdownLinkText(label)}]({link.Url})")
        Next

        Return sb.ToString().Trim()
    End Function

    Private Function ExtractM365SourceLinks(toolResponses As List(Of ToolResponse),
                                            existingAnswer As String) As List(Of ToolSourceLink)
        Dim results As New List(Of ToolSourceLink)()
        Dim seenUrls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim answerText As String = If(existingAnswer, "")

        If toolResponses Is Nothing OrElse toolResponses.Count = 0 Then
            Return results
        End If

        For Each response As ToolResponse In toolResponses
            If response Is Nothing OrElse Not response.Success OrElse String.IsNullOrWhiteSpace(response.Response) Then
                Continue For
            End If

            If String.Equals(response.ToolName, "m365_search", StringComparison.OrdinalIgnoreCase) Then
                ExtractM365SearchLinks(response.Response, answerText, seenUrls, results)
            ElseIf IsM365RetrievalToolName(response.ToolName) Then
                ExtractM365WrappedContentLink(response.ToolName, response.Response, answerText, seenUrls, results)
            End If

            If results.Count >= 12 Then
                Exit For
            End If
        Next

        Return results
    End Function

    Private Sub ExtractM365SearchLinks(responseText As String,
                                       existingAnswer As String,
                                       seenUrls As HashSet(Of String),
                                       results As List(Of ToolSourceLink))
        Dim root As JObject = Nothing

        Try
            root = JObject.Parse(responseText)
        Catch
            Exit Sub
        End Try

        Dim hits As JArray = TryCast(root("hits"), JArray)
        If hits Is Nothing OrElse hits.Count = 0 Then
            Exit Sub
        End If

        For Each hitToken As JToken In hits
            Dim hit As JObject = TryCast(hitToken, JObject)
            If hit Is Nothing Then Continue For

            Dim url As String = If(hit("web_url")?.ToString(), "").Trim()
            Dim title As String = If(hit("title")?.ToString(), "").Trim()
            Dim source As String = If(hit("source")?.ToString(), "").Trim()

            TryAddSourceLink(url, title, source, existingAnswer, seenUrls, results)

            If results.Count >= 12 Then
                Exit For
            End If
        Next
    End Sub

    Private Sub ExtractM365WrappedContentLink(toolName As String,
                                              responseText As String,
                                              existingAnswer As String,
                                              seenUrls As HashSet(Of String),
                                              results As List(Of ToolSourceLink))
        Dim urlMatch As Match = Regex.Match(
            responseText,
            "<WEB_URL>\s*(.*?)\s*</WEB_URL>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        If Not urlMatch.Success Then
            Exit Sub
        End If

        Dim titleMatch As Match = Regex.Match(
            responseText,
            "^<(?<kind>[A-Z_]+)\s+id=""[^""]*""\s+title=""(?<title>[^""]*)""[^>]*>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        Dim title As String = ""
        If titleMatch.Success Then
            title = titleMatch.Groups("title").Value.Trim()
        End If

        Dim url As String = urlMatch.Groups(1).Value.Trim()
        Dim source As String = GetM365SourceFromToolName(toolName)

        TryAddSourceLink(url, title, source, existingAnswer, seenUrls, results)
    End Sub

    Private Sub TryAddSourceLink(url As String,
                                 title As String,
                                 source As String,
                                 existingAnswer As String,
                                 seenUrls As HashSet(Of String),
                                 results As List(Of ToolSourceLink))
        Dim cleanUrl As String = If(url, "").Trim()
        If String.IsNullOrWhiteSpace(cleanUrl) Then
            Exit Sub
        End If

        If Not String.IsNullOrWhiteSpace(existingAnswer) AndAlso
           existingAnswer.IndexOf(cleanUrl, StringComparison.OrdinalIgnoreCase) >= 0 Then
            Exit Sub
        End If

        If Not seenUrls.Add(cleanUrl) Then
            Exit Sub
        End If

        results.Add(New ToolSourceLink With {
            .Url = cleanUrl,
            .Title = If(title, "").Trim(),
            .Source = If(source, "").Trim()
        })
    End Sub

    Private Function IsM365RetrievalToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Select Case toolName.Trim().ToLowerInvariant()
            Case "m365_get_mail",
                 "m365_get_mail_thread",
                 "m365_get_file",
                 "m365_get_event",
                 "m365_get_chat_thread",
                 "m365_get_onenote_page"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function GetM365SourceFromToolName(toolName As String) As String
        If String.IsNullOrWhiteSpace(toolName) Then
            Return ""
        End If

        Select Case toolName.Trim().ToLowerInvariant()
            Case "m365_get_mail", "m365_get_mail_thread"
                Return "mail"
            Case "m365_get_file"
                Return "file"
            Case "m365_get_event"
                Return "calendar"
            Case "m365_get_chat_thread"
                Return "teams"
            Case "m365_get_onenote_page"
                Return "onenote"
            Case Else
                Return "m365"
        End Select
    End Function

    Private Function BuildSourceLinkLabel(link As ToolSourceLink) As String
        If link Is Nothing Then Return "Open source"

        Dim title As String = If(link.Title, "").Trim()
        Dim source As String = If(link.Source, "").Trim().ToLowerInvariant()

        If String.IsNullOrWhiteSpace(title) Then
            title = "Open item"
        End If

        Select Case source
            Case "mail"
                Return title & " (e-mail)"
            Case "onedrive"
                Return title & " (OneDrive)"
            Case "sharepoint"
                Return title & " (SharePoint)"
            Case "file"
                Return title & " (file)"
            Case "teams"
                Return title & " (Teams)"
            Case "calendar"
                Return title & " (calendar)"
            Case "onenote"
                Return title & " (OneNote)"
            Case Else
                Return title
        End Select
    End Function

    Private Function EscapeMarkdownLinkText(value As String) As String
        Dim s As String = If(value, "")
        s = s.Replace("\", "\\")
        s = s.Replace("[", "\[")
        s = s.Replace("]", "\]")
        Return s
    End Function


End Class