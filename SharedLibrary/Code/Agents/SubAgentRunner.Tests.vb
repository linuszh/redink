' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SubAgentRunner.Tests.vb
' Purpose: Self-tests for SubAgentRunner envelope handling, retry behavior,
'          and memory grounding. Tests verify correct parsing of sub-agent output
'          (envelope vs. direct JSON vs. plain text) and retry-on-empty logic.
'
' Tests:
'  - Envelope format: {summary, result} preserved as-is.
'  - Direct JSON objects/arrays: preserved as structured results.
'  - Plain text fallback: auto-generates summary when jsonRequired=false.
'  - Empty output: triggers agent_empty_result, retries exactly once.
'  - Retry success: second attempt recovers from empty first attempt.
'  - Model empty response: retry carries compact recovery prompt.
' =============================================================================

#If DEBUG Then

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq

Namespace AgentsXX

    Friend NotInheritable Class SubAgentRunnerTests

        Private Sub New()
        End Sub

        Public Shared Async Function RunAllAsync() As Task
            TestEnvelopeJsonResultStillWorks()
            TestDirectDomainJsonObjectIsPreservedAsJson()
            TestJsonArrayIsPreservedAsJson()
            TestPlainTextFallbackStillWorks()
            Await TestEmptyFinalTextReturnsAgentEmptyResultAsync().ConfigureAwait(False)
            Await TestEmptyJsonObjectReturnsAgentEmptyResultAsync().ConfigureAwait(False)
            Await TestEmptyEnvelopeReturnsAgentEmptyResultAsync().ConfigureAwait(False)
            Await TestRetryOccursExactlyOnceForAgentEmptyResultAsync().ConfigureAwait(False)
            Await TestModelEmptyResponseRetryCarriesCompactSummaryAsync().ConfigureAwait(False)
        End Function

        Private Shared Sub TestEnvelopeJsonResultStillWorks()
            Dim parsed = SubAgentRuntimeHardening.NormalizeFinalOutput(
                "{""summary"":""done"",""result"":{""path"":""a.txt"",""ok"":true}}",
                jsonRequired:=True)

            AssertFalse(parsed.IsError, "Envelope JSON should be usable.")
            AssertEqual("envelope", parsed.ResultKind, "Envelope resultKind mismatch.")
            AssertEqual("done", parsed.Summary, "Envelope summary mismatch.")
            AssertTrue(TypeOf parsed.Result Is JObject, "Envelope result should stay as JObject.")
            AssertEqual("a.txt", parsed.Result("path")?.ToString(), "Envelope result path mismatch.")
        End Sub

        Private Shared Sub TestDirectDomainJsonObjectIsPreservedAsJson()
            Dim parsed = SubAgentRuntimeHardening.NormalizeFinalOutput(
                "{""path"":""row-1.json"",""status"":""done"",""row"":{},""error"":null}",
                jsonRequired:=True)

            AssertFalse(parsed.IsError, "Direct JSON object should be usable.")
            AssertEqual("json_object", parsed.ResultKind, "Direct object resultKind mismatch.")
            AssertEqual("Sub-agent returned structured JSON.", parsed.Summary, "Direct object fallback summary mismatch.")
            AssertTrue(TypeOf parsed.Result Is JObject, "Direct object result should stay as JObject.")
            AssertEqual("done", parsed.Result("status")?.ToString(), "Direct object status mismatch.")
        End Sub

        Private Shared Sub TestJsonArrayIsPreservedAsJson()
            Dim parsed = SubAgentRuntimeHardening.NormalizeFinalOutput("[{""x"":1}]", jsonRequired:=True)

            AssertFalse(parsed.IsError, "JSON array should be usable.")
            AssertEqual("json_array", parsed.ResultKind, "JSON array resultKind mismatch.")
            AssertEqual("Sub-agent returned structured JSON.", parsed.Summary, "JSON array fallback summary mismatch.")
            AssertTrue(TypeOf parsed.Result Is JArray, "JSON array result should stay as JArray.")
            AssertEqual("1", parsed.Result(0)("x")?.ToString(), "JSON array item mismatch.")
        End Sub

        Private Shared Sub TestPlainTextFallbackStillWorks()
            Dim parsed = SubAgentRuntimeHardening.NormalizeFinalOutput("finished normally", jsonRequired:=False)

            AssertFalse(parsed.IsError, "Plain text fallback should be usable.")
            AssertEqual("text", parsed.ResultKind, "Plain text resultKind mismatch.")
            AssertEqual("finished normally", parsed.Summary, "Plain text summary mismatch.")
            AssertEqual("finished normally", parsed.Result?.ToString(), "Plain text result mismatch.")
        End Sub

        Private Shared Async Function TestEmptyFinalTextReturnsAgentEmptyResultAsync() As Task
            Dim host As New SequenceHost("", "")
            Dim agent As New AgentDescriptor() With {.Name = "empty_agent"}

            Dim raw = Await SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "test task",
                storeResultInMemory:=False).ConfigureAwait(False)

            Dim obj = JObject.Parse(raw)
            AssertEqual("error", obj.Value(Of String)("resultKind"), "Empty text should return error resultKind.")
            AssertEqual("agent_empty_result", obj("error")?.Value(Of String)("code"), "Empty text should return agent_empty_result.")
            AssertEqual("final_output_parse", obj("error")?.Value(Of String)("phase"), "Empty text phase mismatch.")
            AssertEqual(2, host.CallCount, "Empty text should retry exactly once.")
        End Function

        Private Shared Async Function TestEmptyJsonObjectReturnsAgentEmptyResultAsync() As Task
            Dim host As New SequenceHost("{}", "{}")
            Dim agent As New AgentDescriptor() With {.Name = "empty_object_agent"}

            Dim raw = Await SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "test task",
                storeResultInMemory:=False).ConfigureAwait(False)

            Dim obj = JObject.Parse(raw)
            AssertEqual("error", obj.Value(Of String)("resultKind"), "Empty object should return error resultKind.")
            AssertEqual("agent_empty_result", obj("error")?.Value(Of String)("code"), "Empty object should return agent_empty_result.")
            AssertEqual(2, host.CallCount, "Empty object should retry exactly once.")
        End Function

        Private Shared Async Function TestEmptyEnvelopeReturnsAgentEmptyResultAsync() As Task
            Dim host As New SequenceHost("{""summary"":"""",""result"":null}", "{""summary"":"""",""result"":null}")
            Dim agent As New AgentDescriptor() With {.Name = "empty_envelope_agent"}

            Dim raw = Await SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "test task",
                storeResultInMemory:=False).ConfigureAwait(False)

            Dim obj = JObject.Parse(raw)
            AssertEqual("error", obj.Value(Of String)("resultKind"), "Empty envelope should return error resultKind.")
            AssertEqual("agent_empty_result", obj("error")?.Value(Of String)("code"), "Empty envelope should return agent_empty_result.")
            AssertEqual(2, host.CallCount, "Empty envelope should retry exactly once.")
        End Function

        Private Shared Async Function TestRetryOccursExactlyOnceForAgentEmptyResultAsync() As Task
            Dim host As New SequenceHost(
                "",
                "{""summary"":""ok"",""result"":{""status"":""done""}}")
            Dim agent As New AgentDescriptor() With {.Name = "retry_agent"}

            Dim raw = Await SubAgentRunner.InvokeResolvedAsync(
                host,
                agent,
                "retry task",
                contextBlob:="ctx",
                storeResultInMemory:=False).ConfigureAwait(False)

            Dim obj = JObject.Parse(raw)
            AssertTrue(obj("error") Is Nothing, "Retry success should not return an error payload.")
            AssertEqual("envelope", obj.Value(Of String)("resultKind"), "Retry success resultKind mismatch.")
            AssertEqual(2, host.CallCount, "agent_empty_result should retry exactly once.")
            AssertTrue(host.UserMessages.Count = 2, "Expected two host invocations.")
            AssertTrue(host.UserMessages(1).IndexOf("empty or unusable", StringComparison.OrdinalIgnoreCase) >= 0, "Retry prompt reminder mismatch.")
            AssertEqual("done", obj("result")("status")?.ToString(), "Retry result mismatch.")
        End Function


        Private Shared Async Function TestModelEmptyResponseRetryCarriesCompactSummaryAsync() As Task
            Dim host As New SequenceHost(
        SubAgentRuntimeHardening.BuildModelEmptyResponsePayload(
            lastToolName:="workspace_extract_text",
            lastToolResultSummary:="workspace_extract_text succeeded with a compact excerpt; more text is available via next_offset.",
            compactedToolResponse:=True,
            retryHint:="Return final JSON or request a smaller chunk."),
        "{""summary"":""ok"",""result"":{""status"":""done""}}")

            Dim agent As New AgentDescriptor() With {
        .Name = "retry_after_model_empty"
    }
            agent.AllowedTools.Add("workspace_extract_text")
            agent.AllowedTools.Add("workspace_read")
            agent.AllowedTools.Add("js_run")
            agent.AllowedTools.Add("tool_loader")

            Dim raw = Await SubAgentRunner.InvokeResolvedAsync(
        host,
        agent,
        "retry task",
        contextBlob:="ctx",
        storeResultInMemory:=False).ConfigureAwait(False)

            Dim obj = JObject.Parse(raw)
            AssertTrue(obj("error") Is Nothing, "Retry success should not return an error payload.")
            AssertEqual(2, host.CallCount, "model_empty_response should retry exactly once.")
            AssertTrue(host.UserMessages(1).IndexOf("workspace_extract_text", StringComparison.OrdinalIgnoreCase) >= 0, "Retry prompt should include last tool name.")
            AssertTrue(host.UserMessages(1).IndexOf("smaller chunk", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               host.UserMessages(1).IndexOf("start_char", StringComparison.OrdinalIgnoreCase) >= 0, "Retry prompt should instruct chunked follow-up.")
            AssertTrue(host.UserMessages(1).Length < 4000, "Retry prompt should stay compact.")
        End Function


        Private Shared Sub AssertTrue(condition As Boolean, message As String)
            If Not condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub AssertFalse(condition As Boolean, message As String)
            If condition Then
                Throw New InvalidOperationException(message)
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As String, actual As String, message As String)
            If Not String.Equals(expected, actual, StringComparison.Ordinal) Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private Shared Sub AssertEqual(expected As Integer, actual As Integer, message As String)
            If expected <> actual Then
                Throw New InvalidOperationException($"{message} Expected='{expected}', Actual='{actual}'.")
            End If
        End Sub

        Private NotInheritable Class SequenceHost
            Implements ISubAgentHost

            Private ReadOnly _responses As Queue(Of String)

            Public ReadOnly Property UserMessages As New List(Of String)()
            Public Property CallCount As Integer

            Public Sub New(ParamArray responses() As String)
                _responses = New Queue(Of String)(If(responses, Array.Empty(Of String)()))
            End Sub

            Public Function RunIsolatedToolingLoopAsync(request As SubAgentRunRequest,
                                                        cancellationToken As CancellationToken) As Task(Of String) _
                Implements ISubAgentHost.RunIsolatedToolingLoopAsync

                CallCount += 1
                UserMessages.Add(If(request?.UserMessage, ""))

                If _responses.Count = 0 Then
                    Return Task.FromResult(String.Empty)
                End If

                Return Task.FromResult(_responses.Dequeue())
            End Function
        End Class

    End Class

End Namespace

#End If