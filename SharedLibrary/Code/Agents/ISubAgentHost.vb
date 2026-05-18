' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ISubAgentHost.vb
' Purpose: Contract by which SharedLibrary asks the host add-in (Word or Outlook)
'          to run an isolated tooling-loop pass for a sub-agent. The host owns
'          the loop implementation; this interface only carries the inputs the
'          sub-agent needs and returns the final assistant text.
'
' The host implementation is expected to:
'   - Start a *clean* message history (no parent system prompt, no parent turns).
'   - Use SubAgentRunRequest.SystemPrompt as the only system prompt
'     (this already includes the AGENT.md body composed by SubAgentRunner).
'   - Use SubAgentRunRequest.UserMessage as the single user turn.
'   - Restrict tool availability to SubAgentRunRequest.AllowedToolNames
'     (a narrowing filter; if Nothing or empty => no narrowing).
'   - Honor SubAgentRunRequest.SpecialModelKey by resolving the corresponding
'     special-task-model and using it for this run only (restoring the previous
'     model afterwards), falling back to "agentdefaultmodel".
'   - Return the final assistant text (the runner will try to parse it as JSON).
'   - Restrict tool availability to SubAgentRunRequest.AllowedToolNames.
'     If AllowedToolNames is Nothing, the host may use its default active tool set.
'     If AllowedToolNames is an empty list, the sub-agent gets no callable tools.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    Public Class SubAgentRunRequest
        Public Property AgentName As String
        Public Property SystemPrompt As String
        Public Property UserMessage As String
        Public Property SpecialModelKey As String
        Public Property AllowedToolNames As IReadOnlyList(Of String)
        Public Property MaxIterations As Integer
        Public Property TimeoutSeconds As Integer
        Public Property WorkflowId As String
    End Class

    Public Interface ISubAgentHost
        Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                             cancellationToken As CancellationToken) As Task(Of String)
    End Interface

End Namespace