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

        ' 1) Snapshot current model config.
        Dim backup As ModelConfig = GetCurrentConfig(_context)
        Try

            ' 2) Resolve and apply the requested special-task-model (fallback agentdefaultmodel).
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

            ' 3) Build the tool list available to the sub-agent. Base on the user's currently
            '    selected tools (so sources stay user-controlled) and always include memory_*.
            Dim allTools As New List(Of ModelConfig)()
            Try
                Dim available = GetAvailableTools()
                If SelectedToolNames IsNot Nothing AndAlso SelectedToolNames.Count > 0 Then
                    For Each mc In available
                        If mc Is Nothing OrElse String.IsNullOrWhiteSpace(mc.ToolName) Then Continue For
                        If SelectedToolNames.Contains(mc.ToolName) Then allTools.Add(mc)
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine($"[WORD-AGENTHOST] GetAvailableTools failed: {ex}")
            End Try
            allTools.AddRange(SharedLibrary.Agents.MemoryTools.BuildAll())

            ' Narrow per agent's allowed-tools (if provided).
            Dim narrowed As List(Of ModelConfig)
            If request.AllowedToolNames Is Nothing OrElse request.AllowedToolNames.Count = 0 Then
                narrowed = allTools
            Else
                Dim allow As New HashSet(Of String)(request.AllowedToolNames, StringComparer.OrdinalIgnoreCase)
                narrowed = allTools.Where(Function(t) t IsNot Nothing AndAlso allow.Contains(t.ToolName)).ToList()
            End If

            ' 4) Run ONE isolated tooling loop. Word's ExecuteToolingLoop has no cancellationToken
            '    parameter; cancellation is governed by the gate and the host's own flow.
            Dim result = Await ExecuteToolingLoop(
                sysCommand:=request.SystemPrompt,
                userText:="",
                selectedTools:=narrowed,
                useSecondAPI:=False,
                fullPromptOverride:=request.UserMessage,
                hideSplash:=True,
                hideLogWindow:=True,
                subAgentMode:=True).ConfigureAwait(False)
            Return If(result, "")
        Finally
            Try : RestoreDefaults(_context, backup) : Catch : End Try
        End Try
    End Function

End Class