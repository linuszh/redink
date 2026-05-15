' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.AgentHost.vb
' Purpose: Agent host that manages sub-agent execution and tooling loops.
' =============================================================================

Option Strict Off
Option Explicit On

Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn
    Implements Agents.ISubAgentHost

    ''' <summary>
    ''' Sub-agent isolated tooling-loop pass. Snapshots the current model, swaps to the agent's
    ''' resolved special-task-model (fallback agentdefaultmodel), narrows tools, runs ONE
    ''' ExecuteToolingLoop with a clean message history, then restores the model.
    ''' </summary>
    Public Async Function RunIsolatedToolingLoopAsync(
            request As SharedLibrary.Agents.SubAgentRunRequest,
            ct As CancellationToken) As Task(Of String) _
            Implements SharedLibrary.Agents.ISubAgentHost.RunIsolatedToolingLoopAsync

        Dim backup As ModelConfig = GetCurrentConfig(_context)

        Try
            Dim modelKey As String = If(String.IsNullOrWhiteSpace(request.SpecialModelKey), "agentdefaultmodel", request.SpecialModelKey)
            Dim swapped As Boolean = False

            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                swapped = GetSpecialTaskModel(_context, INI_AlternateModelPath, modelKey)

                If Not swapped AndAlso Not modelKey.Equals("agentdefaultmodel", StringComparison.OrdinalIgnoreCase) Then
                    swapped = GetSpecialTaskModel(_context, INI_AlternateModelPath, "agentdefaultmodel")
                End If
            Else
                Debug.WriteLine("[WORD-AGENTHOST] INI_AlternateModelPath empty")
            End If

            Dim allTools As List(Of ModelConfig) = Nothing

            If _activeToolingContext IsNot Nothing AndAlso
               _activeToolingContext.AllowedToolRegistry IsNot Nothing Then

                allTools = _activeToolingContext.AllowedToolRegistry.MaterializeAll()
                ToolingFileLogger.LogStep("[subagent-host] Using parent tooling registry as sub-agent tool universe.")
            Else
                allTools = GetDiscussInkyEffectiveTools(includeImplicitWorkspaceTools:=True)
                ToolingFileLogger.LogWarn("[subagent-host] Parent tooling registry unavailable; falling back to DiscussInky effective tools.")
            End If

            allTools =
                If(allTools, New List(Of ModelConfig)()).
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
                    Select(Function(g) g.First()).
                    ToList()

            ToolingFileLogger.LogStep("[subagent-host] Requested allowed tools: " &
                                      If(request.AllowedToolNames Is Nothing OrElse request.AllowedToolNames.Count = 0,
                                         "(none)",
                                         String.Join(", ", request.AllowedToolNames)))

            ToolingFileLogger.LogStep("[subagent-host] Resolved sub-agent tool universe: " &
                                      If(allTools.Count = 0,
                                         "(none)",
                                         String.Join(", ", allTools.Select(Function(t) t.ToolName))))


            allTools.AddRange(SharedLibrary.Agents.WorkspaceTools.BuildAll())

            allTools =
                allTools.
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    GroupBy(Function(t) t.ToolName, StringComparer.OrdinalIgnoreCase).
                    Select(Function(g) g.First()).
                    ToList()

            Dim result = Await ExecuteToolingLoop(
                sysCommand:=request.SystemPrompt,
                userText:="",
                selectedTools:=allTools,
                useSecondAPI:=False,
                fullPromptOverride:=request.UserMessage,
                hideSplash:=True,
                hideLogWindow:=True,
                subAgentMode:=True,
                subAgentAllowedToolNames:=request.AllowedToolNames,
                subAgentSpecialModelKey:=request.SpecialModelKey).ConfigureAwait(False)

            Return If(result, "")
        Finally
            Try : RestoreDefaults(_context, backup) : Catch : End Try
        End Try
    End Function

End Class