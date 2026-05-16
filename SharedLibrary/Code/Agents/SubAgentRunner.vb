' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SubAgentRunner.vb
' Purpose: Orchestrates invocation of a sub-agent (Claude-style AGENT.md):
'           1. Locks the global AgentGate as the OWNER for the whole run so
'              nested LLM()/MCP calls inside this sub-agent's loop are
'              re-entrant against the same logical owner.
'           2. Composes a clean system prompt from the AGENT.md body
'              (no Inky.md, no parent system prompt) and runs ONE isolated
'              tooling-loop pass via the host (ISubAgentHost).
'           3. Parses the assistant's final text as {summary, result} JSON;
'              if missing, preserves direct JSON objects/arrays as structured
'              results instead of stringifying them.
'           4. Validates that the final output is not semantically empty.
'           5. Retries agent_empty_result exactly once with a stricter reminder.
'           6. Optionally stores the result in SessionMemory and returns a
'              compact tool-response JSON that includes the memory stub.
'
' Concurrency: The owner-scope on AgentGate means only one sub-agent can run
' globally at any moment, satisfying the "no parallel model calls" requirement.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SubAgentRunner

        Private Sub New()
        End Sub

        Private Const RetryReminderText As String =
            "Your previous response was empty or unusable. Return one usable final answer in the requested format. If you cannot complete the task, return a structured error object."

        Public Shared Async Function InvokeAsync(host As ISubAgentHost,
                                                 agentName As String,
                                                 task As String,
                                                 Optional contextBlob As String = Nothing,
                                                 Optional storeResultInMemory As Boolean = True,
                                                 Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)

            If host Is Nothing Then
                Return BuildInfrastructureErrorPayload(agentName, "no_host", "invoke", "No sub-agent host is available.")
            End If

            Dim ag = AgentResources.FindAgent(agentName)
            If ag Is Nothing Then
                Return BuildInfrastructureErrorPayload(agentName, "agent_not_found", "invoke", "The requested sub-agent was not found.")
            End If

            Return Await InvokeResolvedAsync(host, ag, task, contextBlob, storeResultInMemory, cancellationToken).ConfigureAwait(False)
        End Function

        Friend Shared Async Function InvokeResolvedAsync(host As ISubAgentHost,
                                                         ag As AgentDescriptor,
                                                         task As String,
                                                         Optional contextBlob As String = Nothing,
                                                         Optional storeResultInMemory As Boolean = True,
                                                         Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)

            If host Is Nothing Then
                Return BuildInfrastructureErrorPayload(If(ag?.Name, ""), "no_host", "invoke", "No sub-agent host is available.")
            End If

            If ag Is Nothing OrElse String.IsNullOrWhiteSpace(ag.Name) Then
                Return BuildInfrastructureErrorPayload(If(ag?.Name, ""), "agent_not_found", "invoke", "The requested sub-agent was not found.")
            End If

            Dim sys As New StringBuilder()

            If Not String.IsNullOrWhiteSpace(ag.Description) Then
                sys.AppendLine(ag.Description.Trim())
                sys.AppendLine()
            End If

            sys.Append(ag.LoadBody())
            sys.AppendLine()
            sys.AppendLine()
            sys.AppendLine("Output contract: your FINAL message MUST be one usable final answer in the requested format. If you cannot complete the task, return a structured error object.")

            Dim baseUserMessage As New StringBuilder()
            baseUserMessage.Append("Task:").AppendLine().Append(If(task, "").Trim())

            If Not String.IsNullOrWhiteSpace(contextBlob) Then
                baseUserMessage.AppendLine().AppendLine().AppendLine("Context (provided by parent):")
                baseUserMessage.Append(contextBlob.Trim())
            End If

            Dim allowedTools As IReadOnlyList(Of String) =
                If(ag.AllowedTools Is Nothing,
                   CType(Array.Empty(Of String)(), IReadOnlyList(Of String)),
                   ag.AllowedTools.AsReadOnly())

            Dim retryCount As Integer = 0
            Dim userMessageForRun As String = baseUserMessage.ToString()

            Await AgentGate.EnterAsync(cancellationToken).ConfigureAwait(False)
            AgentGate.MarkCurrentFlowAsOwner()

            Try
                Do
                    Dim req As New SubAgentRunRequest With {
                        .AgentName = ag.Name,
                        .SystemPrompt = sys.ToString(),
                        .UserMessage = userMessageForRun,
                        .SpecialModelKey = If(String.IsNullOrWhiteSpace(ag.Model), "agentdefaultmodel", ag.Model),
                        .AllowedToolNames = allowedTools,
                        .MaxIterations = 0,
                        .TimeoutSeconds = ag.TimeoutSeconds
                    }

                    Debug.WriteLine(
                        $"[SubAgentRunner] agent='{ag.Name}' allowed_tools={FormatAllowedTools(req.AllowedToolNames)} retry={retryCount}")

                    Dim finalText As String = Nothing

                    Try
                        finalText = Await host.RunIsolatedToolingLoopAsync(req, cancellationToken).ConfigureAwait(False)
                    Catch oce As OperationCanceledException
                        Throw
                    Catch ex As Exception
                        Return BuildInfrastructureErrorPayload(ag.Name, "agent_failed", "invoke", ex.Message)
                    End Try

                    Dim normalized = SubAgentRuntimeHardening.NormalizeFinalOutput(finalText, jsonRequired:=True)
                    LogFinalOutputDiagnostics(ag.Name, req.AllowedToolNames, normalized, retryCount)

                    If Not normalized.IsError Then
                        Dim resp As JObject = normalized.ToJObject()
                        resp("agent") = ag.Name

                        If storeResultInMemory Then
                            Try
                                Dim key As String = "agent_" & ag.Name & "_" & DateTime.UtcNow.ToString("yyyyMMddHHmmssfff")
                                Dim storedResult As JToken = If(normalized.Result Is Nothing, JValue.CreateNull(), normalized.Result.DeepClone())
                                Dim entry = SessionMemory.Put(
                                    key,
                                    If(normalized.Summary, "Result of sub-agent '" & ag.Name & "'."),
                                    storedResult,
                                    tags:={"agent", ag.Name})

                                resp("memory_key") = entry.Key
                                resp("stub") = SessionMemory.BuildStub(entry)
                            Catch
                            End Try
                        End If

                        Return resp.ToString(Formatting.None)
                    End If

                    Dim retryableErrorCodes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
    SubAgentRuntimeHardening.EmptyResultCode,
    SubAgentRuntimeHardening.ModelEmptyResponseCode
}

                    If retryCount = 0 AndAlso retryableErrorCodes.Contains(normalized.GetErrorCode()) Then
                        retryCount += 1

                        Debug.WriteLine(
        $"[SubAgentRunner] agent='{ag.Name}' retrying unusable output. retry={retryCount} errorCode={normalized.GetErrorCode()}")

                        userMessageForRun = BuildRetryUserMessage(baseUserMessage.ToString())
                        Continue Do
                    End If

                    Dim errResp As JObject = normalized.ToJObject()
                    errResp("agent") = ag.Name
                    errResp("retryCount") = retryCount
                    Return errResp.ToString(Formatting.None)
                Loop
            Finally
                AgentGate.UnmarkCurrentFlowAsOwner()
                AgentGate.Release()
            End Try
        End Function

        Private Shared Function BuildRetryUserMessage(baseUserMessage As String) As String
            Dim sb As New StringBuilder(If(baseUserMessage, "").TrimEnd())

            If sb.Length > 0 Then
                sb.AppendLine()
                sb.AppendLine()
            End If

            sb.Append(RetryReminderText)
            Return sb.ToString()
        End Function

        Private Shared Function BuildInfrastructureErrorPayload(agentName As String,
                                                                errorCode As String,
                                                                phase As String,
                                                                message As String) As String
            Dim obj As New JObject(
                New JProperty("summary", If(message, "Sub-agent failed.")),
                New JProperty("result", JValue.CreateNull()),
                New JProperty("resultKind", "error"),
                New JProperty("rawLength", 0),
                New JProperty("error", New JObject(
                    New JProperty("code", errorCode),
                    New JProperty("phase", phase),
                    New JProperty("message", If(message, "")))))

            If Not String.IsNullOrWhiteSpace(agentName) Then
                obj("agent") = agentName
            End If

            Return obj.ToString(Formatting.None)
        End Function

        Private Shared Sub LogFinalOutputDiagnostics(agentName As String,
                                                     allowedTools As IReadOnlyList(Of String),
                                                     normalized As SubAgentRuntimeHardening.NormalizedEnvelope,
                                                     retryCount As Integer)
            Dim errorCode As String = ""
            If normalized IsNot Nothing Then
                errorCode = normalized.GetErrorCode()
            End If

            Debug.WriteLine(
                $"[SubAgentRunner] agent='{agentName}' allowed_tools={FormatAllowedTools(allowedTools)} final_len={If(normalized?.RawLength, 0)} resultKind={If(normalized?.ResultKind, "")} retry={retryCount} errorCode={If(errorCode, "")}")
        End Sub

        Private Shared Function FormatAllowedTools(allowedTools As IReadOnlyList(Of String)) As String
            If allowedTools Is Nothing Then Return "(default-host-scope)"
            If allowedTools.Count = 0 Then Return "(none)"
            Return String.Join(", ", allowedTools)
        End Function

    End Class

End Namespace