' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.Tooling.Memory.vb

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn


    Private Function BuildMemoryGroundingClassifierInput(latestUserRequestRaw As String,
                                                         hostTaskSummary As String) As String
        If String.IsNullOrWhiteSpace(latestUserRequestRaw) Then
            Return ""
        End If

        Return SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingIntentClassifierUserPrompt(
            latestUserRequestRaw,
            hostTaskSummary)
    End Function


    Private Function HasMemoryGroundingClassifierInputsAvailable(context As ToolExecutionContext) As Boolean
        If context Is Nothing Then
            Return False
        End If

        Dim memoryToolsAvailable As Boolean = False

        If context.AuthoritativeToolRegistrySnapshot IsNot Nothing Then
            memoryToolsAvailable =
                context.AuthoritativeToolRegistrySnapshot.Contains(SharedLibrary.Agents.MemoryTools.ToolList) OrElse
                context.AuthoritativeToolRegistrySnapshot.Contains(SharedLibrary.Agents.MemoryTools.ToolGet)
        End If

        If Not memoryToolsAvailable AndAlso context.SelectedTools IsNot Nothing Then
            memoryToolsAvailable =
                context.SelectedTools.Any(
                    Function(tool)
                        Return tool IsNot Nothing AndAlso
                               SharedLibrary.Agents.MemoryTools.IsMemoryTool(tool.ToolName)
                    End Function)
        End If

        Dim workflowMemoryAvailable As Boolean = False

        Try
            workflowMemoryAvailable =
                SharedLibrary.Agents.SessionMemory.ListMostRecentWorkflowEntries(maxItems:=1).Count > 0
        Catch
            workflowMemoryAvailable = False
        End Try

        Return memoryToolsAvailable OrElse workflowMemoryAvailable
    End Function

    Private Async Function ResolveMemoryGroundingModeAsync(context As ToolExecutionContext,
                                                           userText As String,
                                                           otherPrompt As String,
                                                           fullPromptOverride As String,
                                                           useSecondAPI As Boolean,
                                                           hideSplash As Boolean,
                                                           explicitMemoryGroundingMode As SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode,
                                                           memoryGroundingModeIsExplicit As Boolean,
                                                           subAgentMode As Boolean) As Task
        If context Is Nothing OrElse context.SequencingState Is Nothing Then
            Return
        End If

        If subAgentMode Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped for sub-agent mode.", "diag")
            Return
        End If

        If memoryGroundingModeIsExplicit Then
            context.SequencingState.MemoryGroundingMode = explicitMemoryGroundingMode
            context.SequencingState.ShouldExposeRecentMemoryStubs =
                explicitMemoryGroundingMode <> SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None

            context.Log("Memory grounding mode applied from explicit override: " &
                SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState))
            Return
        End If

        If Not HasMemoryGroundingClassifierInputsAvailable(context) Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped: no memory tools and no workflow memory available.", "diag")
            Return
        End If

        Dim classifierInput As String =
            BuildMemoryGroundingClassifierInput(
                context.LatestUserRequestRaw,
                context.HostTaskSummary)

        If String.IsNullOrWhiteSpace(classifierInput) Then
            context.SequencingState.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.None
            context.SequencingState.ShouldExposeRecentMemoryStubs = False
            context.Log("Memory grounding classifier skipped: no classifier input was available.", "diag")
            Return
        End If

        context.Log("Memory grounding classifier invoked.", "diag")
        Dim raw As String = ""

        LogLatestUserRequestDiagnostic(context, "classifier")

        Try
            raw = Await LLM(
                SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingIntentClassifierSystemPrompt(),
                classifierInput,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                True)
        Catch ex As Exception
            context.LogWarn("Memory grounding classifier failed; defaulting to none.",
                            details:=$"host={context.HostKind}; error={ex.Message}")
        End Try

        context.Log("classifierRawOutput=" & If(raw, ""), "diag")

        Dim classifierNormalizedOutput As String = ""
        Dim classifierParseError As String = ""
        Dim decision =
            SharedLibrary.Agents.ToolCallSequencing.ParseMemoryGroundingIntentClassifierDecision(
                raw,
                classifierNormalizedOutput,
                classifierParseError)

        Dim normalizedOutputForLog As String = If(classifierNormalizedOutput, "")
        If normalizedOutputForLog.Length > 600 Then
            normalizedOutputForLog = normalizedOutputForLog.Substring(0, 600) & "..."
        End If

        context.Log("classifierNormalizedOutput=" & normalizedOutputForLog, "diag")
        context.Log("classifierParseSuccess=" & If(decision.IsValid, "true", "false"), "diag")

        If Not decision.IsValid Then
            context.Log("classifierParseError=" & If(classifierParseError, "invalid_classifier_output"), "diag")
            context.LogWarn(
                "classifierParseFailure",
                details:=$"host={context.HostKind}; classifierParseError={If(classifierParseError, "invalid_classifier_output")}")
        End If

        context.Log(
            "parsedMemoryGroundingMode=" &
            SharedLibrary.Agents.ToolCallSequencing.FormatMemoryGroundingMode(decision.MemoryGroundingMode) &
            "; shouldExposeRecentMemoryStubs=" &
            If(decision.ShouldExposeRecentMemoryStubs, "true", "false") &
            "; explicitStoredMemoryRequired=" &
            If(decision.ExplicitStoredMemoryRequired, "true", "false") &
            "; reason=" & If(decision.Reason, ""), "diag")

        If decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.Required AndAlso
           Not decision.ExplicitStoredMemoryRequired Then

            context.LogWarn(
                "classifierRequiredModeDowngraded",
                details:=$"host={context.HostKind}; reason=missing_explicit_user_demand_for_stored_memory; classifierReason={If(decision.Reason, "")}")

            decision.MemoryGroundingMode = SharedLibrary.Agents.ToolCallSequencing.MemoryGroundingMode.OptionalMode
        End If

        context.SequencingState.MemoryGroundingMode = decision.MemoryGroundingMode
        context.SequencingState.ShouldExposeRecentMemoryStubs = decision.ShouldExposeRecentMemoryStubs

        context.Log(
            "appliedMemoryGroundingMode=" &
            SharedLibrary.Agents.ToolCallSequencing.FormatMemoryGroundingMode(context.SequencingState.MemoryGroundingMode) &
            "; explicitStoredMemoryRequired=" &
            If(decision.ExplicitStoredMemoryRequired, "true", "false"), "diag")

        context.Log(
    "Memory grounding mode applied: " &
    SharedLibrary.Agents.ToolCallSequencing.BuildMemoryGroundingStateSummary(context.SequencingState),
    "diag")

    End Function

End Class
