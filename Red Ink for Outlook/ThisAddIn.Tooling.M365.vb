' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.M365.vb (Outlook host adapter)
' Purpose: Bridges Outlook's host-local ToolCall / ToolResponse / ToolExecutionContext
'          to the host-neutral SharedLibrary.M365ToolService.
'
'   Wire-up in ThisAddin.Tooling.vb:
'      In ExecuteToolCall (after the existing knowledge dispatch arm):
'          ElseIf SharedLibrary.M365ToolService.IsM365ToolName(toolCall.ToolName) Then
'              response = Await ExecuteInternalM365Tool(toolCall, context)
'              ToolingFileLogger.LogRawResponseStub(
'                  $"Internal tool ({toolCall.ToolName})", response.Response)
'
'      In GetAvailableTools (after the knowledge tools registration):
'          tools.AddRange(SharedLibrary.M365ToolService.GetTools(_context, InternalToolSuffix))
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary

Partial Public Class ThisAddIn

    Private _lastCompletedToolResponses As New List(Of ToolResponse)()

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

    Friend Function GetLastCompletedToolResponsesSnapshot() As List(Of ToolResponse)
        Dim result As New List(Of ToolResponse)()
        Try
            Debug.WriteLine("[AISearch] GetLastCompletedToolResponsesSnapshot: source count=" &
                            If(_lastCompletedToolResponses Is Nothing, -1, _lastCompletedToolResponses.Count).ToString())

            If _lastCompletedToolResponses Is Nothing Then Return result

            For Each r In _lastCompletedToolResponses
                result.Add(New ToolResponse() With {
                    .CallId = r.CallId,
                    .ToolName = r.ToolName,
                    .Response = r.Response,
                    .Success = r.Success,
                    .ErrorMessage = r.ErrorMessage,
                    .Timestamp = r.Timestamp,
                    .OriginalCallJson = r.OriginalCallJson
                })
            Next
        Catch
        End Try
        Return result
    End Function

End Class