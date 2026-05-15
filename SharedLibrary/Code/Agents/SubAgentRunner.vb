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
'              if missing, wraps the entire output into a {summary, result}.
'           4. Optionally stores the result in SessionMemory and returns a
'              compact tool-response JSON that includes the memory stub.
'
' Concurrency: The owner-scope on AgentGate means only one sub-agent can run
' globally at any moment, satisfying the "no parallel model calls" requirement.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class SubAgentRunner

        Private Sub New()
        End Sub

        ''' <summary>
        ''' Resolves the agent by name, runs it in isolation via <paramref name="host"/>,
        ''' and returns a JSON string suitable for use as a tool response in the parent
        ''' tooling loop. Shape:
        '''   { "agent": "...", "summary": "...", "memory_key": "mem_...", "stub": "[memory:KEY] SUMMARY", "result": ... }
        ''' </summary>
        Public Shared Async Function InvokeAsync(host As ISubAgentHost,
                                                 agentName As String,
                                                 task As String,
                                                 Optional contextBlob As String = Nothing,
                                                 Optional storeResultInMemory As Boolean = True,
                                                 Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)

            If host Is Nothing Then
                Return JsonConvert.SerializeObject(New With {Key .error = "no_host", Key .agent = agentName})
            End If

            Dim ag = AgentResources.FindAgent(agentName)
            If ag Is Nothing Then
                Return JsonConvert.SerializeObject(New With {Key .error = "agent_not_found", Key .agent = agentName})
            End If

            ' Compose the sub-agent's system prompt from AGENT.md body (clean: no Inky.md,
            ' no parent system prompt). Description goes first as a one-line role line.
            Dim sys As New StringBuilder()
            If Not String.IsNullOrWhiteSpace(ag.Description) Then
                sys.AppendLine(ag.Description.Trim())
                sys.AppendLine()
            End If
            sys.Append(ag.LoadBody())
            sys.AppendLine()
            sys.AppendLine()
            sys.AppendLine("Output contract: your FINAL message MUST be a single JSON object with the keys " &
                          """summary"" (string, one or two sentences) and ""result"" (string or JSON value with " &
                          "the actual deliverable). Do not wrap it in a code fence.")

            ' Single user turn: the task plus optional caller-supplied context.
            Dim usr As New StringBuilder()
            usr.Append("Task:").AppendLine().Append(If(task, "").Trim())
            If Not String.IsNullOrWhiteSpace(contextBlob) Then
                usr.AppendLine().AppendLine().AppendLine("Context (provided by parent):")
                usr.Append(contextBlob.Trim())
            End If

            Dim req As New SubAgentRunRequest With {
                .AgentName = ag.Name,
                .SystemPrompt = sys.ToString(),
                .UserMessage = usr.ToString(),
                .SpecialModelKey = If(String.IsNullOrWhiteSpace(ag.Model), "agentdefaultmodel", ag.Model),
                .AllowedToolNames = If(ag.AllowedTools Is Nothing OrElse ag.AllowedTools.Count = 0,
                                       Nothing, ag.AllowedTools.AsReadOnly()),
                .MaxIterations = 0,
                .TimeoutSeconds = ag.TimeoutSeconds
            }

            ' Take the gate as OWNER so nested LLM/MCP calls inside the host loop don't re-wait.
            Await AgentGate.EnterAsync(cancellationToken).ConfigureAwait(False)

            AgentGate.MarkCurrentFlowAsOwner()
            Try
                Dim finalText As String
                Try
                    finalText = Await host.RunIsolatedToolingLoopAsync(req, cancellationToken).ConfigureAwait(False)
                    If String.IsNullOrWhiteSpace(finalText) Then
                        Return JsonConvert.SerializeObject(New With {
                            Key .error = "agent_empty_result",
                            Key .agent = ag.Name,
                            Key .model = req.SpecialModelKey,
                            Key .allowed_tools = If(req.AllowedToolNames, Array.Empty(Of String)()),
                            Key .message = "The sub-agent tooling loop returned an empty final result."
                        })
                    End If
                Catch oce As OperationCanceledException
                    Throw
                Catch ex As Exception
                    Return JsonConvert.SerializeObject(New With {
                        Key .error = "agent_failed",
                        Key .agent = ag.Name,
                        Key .message = ex.Message
                    })
                End Try

                Dim parsedSummary As String = Nothing
                Dim parsedResult As JToken = Nothing
                ParseAgentOutput(finalText, parsedSummary, parsedResult)

                Dim resp As New JObject()
                resp("agent") = ag.Name
                resp("summary") = If(parsedSummary, "")
                resp("result") = If(parsedResult, JToken.FromObject(""))

                If storeResultInMemory Then
                    Try
                        Dim key As String = "agent_" & ag.Name & "_" & DateTime.UtcNow.ToString("yyyyMMddHHmmssfff")
                        Dim entry = SessionMemory.Put(key,
                                                      If(parsedSummary, "Result of sub-agent '" & ag.Name & "'."),
                                                      If(parsedResult, JToken.FromObject(finalText)),
                                                      tags:={"agent", ag.Name})
                        resp("memory_key") = entry.Key
                        resp("stub") = SessionMemory.BuildStub(entry)
                    Catch
                        ' best effort
                    End Try
                End If

                Return resp.ToString(Formatting.None)
            Finally
                AgentGate.UnmarkCurrentFlowAsOwner()
                AgentGate.Release()
            End Try
        End Function

        ''' <summary>
        ''' Tries to parse the sub-agent's final text as a JSON object with summary/result.
        ''' On failure, falls back to treating the entire text as the result and synthesising
        ''' a brief summary.
        ''' </summary>
        Private Shared Sub ParseAgentOutput(text As String, ByRef summary As String, ByRef result As JToken)
            summary = Nothing
            result = Nothing
            If String.IsNullOrWhiteSpace(text) Then
                summary = "(empty)"
                result = JToken.FromObject("")
                Return
            End If

            Dim candidate = StripCodeFence(text).Trim()
            Try
                Dim tok = JToken.Parse(candidate)
                If TypeOf tok Is JObject Then
                    Dim obj = CType(tok, JObject)
                    Dim s = obj.Value(Of String)("summary")
                    Dim r = obj("result")
                    If Not String.IsNullOrWhiteSpace(s) OrElse r IsNot Nothing Then
                        summary = If(s, "")
                        result = If(r, JToken.FromObject(""))
                        Return
                    End If
                End If
            Catch
                ' not JSON; fall through
            End Try

            ' Fallback: whole text as result.
            result = JToken.FromObject(text)
            summary = BuildFallbackSummary(text)
        End Sub

        Private Shared Function StripCodeFence(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return s
            Dim t = s.Trim()
            If t.StartsWith("```") Then
                Dim firstNl = t.IndexOf(ChrW(10))
                If firstNl > 0 Then t = t.Substring(firstNl + 1)
                If t.EndsWith("```") Then t = t.Substring(0, t.Length - 3)
            End If
            Return t.Trim()
        End Function

        Private Shared Function BuildFallbackSummary(text As String) As String
            Dim line = text.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If line.Length <= 160 Then Return line
            Return line.Substring(0, 157) & "..."
        End Function

    End Class

End Namespace