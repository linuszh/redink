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
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    Public Class SubAgentRunRequest
        ''' <summary>Logical agent name (for logs/diagnostics).</summary>
        Public Property AgentName As String
        ''' <summary>Full system prompt to use (already composed by SubAgentRunner).</summary>
        Public Property SystemPrompt As String
        ''' <summary>Single user message representing the task (and any caller-supplied context).</summary>
        Public Property UserMessage As String
        ''' <summary>Special-task-model key (e.g. "researchmodel"); empty => agentdefaultmodel.</summary>
        Public Property SpecialModelKey As String
        ''' <summary>
        ''' Narrowing filter; if non-empty, only these tool names may be offered to the model.
        ''' Nothing or empty means "no narrowing" (the host's full active tool set applies).
        ''' </summary>
        Public Property AllowedToolNames As IReadOnlyList(Of String)
        ''' <summary>Optional max iterations override; 0 = use host default.</summary>
        Public Property MaxIterations As Integer
        ''' <summary>Optional timeout override in seconds; 0 = use host default.</summary>
        Public Property TimeoutSeconds As Integer
    End Class

    Public Interface ISubAgentHost
        ''' <summary>
        ''' Runs one isolated tooling-loop pass as described by <paramref name="request"/>
        ''' and returns the final assistant text. Implementations MUST NOT carry any state
        ''' from the parent run.
        ''' </summary>
        Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                             cancellationToken As CancellationToken) As Task(Of String)
    End Interface

End Namespace