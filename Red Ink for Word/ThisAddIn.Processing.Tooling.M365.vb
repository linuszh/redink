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

Imports System.Threading.Tasks
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

End Class