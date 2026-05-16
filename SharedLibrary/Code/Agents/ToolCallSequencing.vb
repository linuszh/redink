Option Strict On
Option Explicit On


Imports System.Collections
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class ToolCallSequencing

        Private Sub New()
        End Sub

        Public Const DependentBatchingInstruction As String =
            "When using tools, do not emit multiple tool calls in one response if any later call depends on the result of an earlier call." & vbCrLf &
            "A tool call is dependent if it:" & vbCrLf &
            "- reads data that may have been changed by an earlier call;" & vbCrLf &
            "- uses the result of an earlier call;" & vbCrLf &
            "- selects the next item after a state update;" & vbCrLf &
            "- delegates work based on newly read or newly written state;" & vbCrLf &
            "- depends on whether an earlier call succeeded or failed." & vbCrLf &
            "In such cases, emit only the first required tool call. Wait for its result before deciding the next tool call." & vbCrLf &
            "Only batch multiple tool calls when they are independent read-only operations."

        Public Const UnresolvedToolFailureCode As String = "unresolved_tool_failure"
        Public Const InvalidTextOnlyFinalizationCode As String = "invalid_text_only_finalization"
        Public Const MissingRequiredMemoryAccessCode As String = "missing_required_memory_access"

        Public Enum TaskStatusKind
            None
            Complete
            Blocked
            ContinueTurn
        End Enum

        Public Enum ActiveToolingTurnKind
            InvalidTurn
            ToolCallTurn
            FinalCompleteTurn
            FinalBlockedTurn
        End Enum

        Public Enum MemoryGroundingMode
            None
            OptionalMode
            Required
        End Enum

        Public NotInheritable Class TaskStatusParseResult
            Public Property IsPresent As Boolean
            Public Property IsValid As Boolean
            Public Property Status As TaskStatusKind
            Public Property Reason As String
            Public Property FooterCount As Integer
            Public Property FailureReason As String
            Public Property FooterJson As String
            Public Property TextBeforeFooter As String

            Public ReadOnly Property Summary As String
                Get
                    If Not IsPresent Then Return "missing"
                    If Not IsValid Then Return "invalid:" & If(FailureReason, "")
                    Return Status.ToString().ToLowerInvariant()
                End Get
            End Property
        End Class

        Public NotInheritable Class ActiveToolingTurnValidationResult
            Public Property TurnKind As ActiveToolingTurnKind
            Public Property InvalidReason As String
            Public Property TaskStatus As TaskStatusParseResult

            Public ReadOnly Property TaskStatusSummary As String
                Get
                    If TaskStatus Is Nothing Then Return "missing"
                    Return TaskStatus.Summary
                End Get
            End Property
        End Class

        Public Enum ToolCallClassification
            ReadOnlyIndependent
            Mutating
            Stateful
            Skill
            Agent
            Unknown
        End Enum

        Public NotInheritable Class PlannedToolCall
            Public Property Index As Integer
            Public Property ToolName As String
            Public Property Classification As ToolCallClassification
            Public Property IsBarrier As Boolean
            Public Property WillExecute As Boolean
            Public Property SkipReason As String
        End Class

        Public NotInheritable Class ToolBatchPlan

            Public Sub New()
                Calls = New List(Of PlannedToolCall)()
            End Sub

            Public Property Calls As List(Of PlannedToolCall)

            Public ReadOnly Property TotalCallCount As Integer
                Get
                    Return Calls.Count
                End Get
            End Property

            Public ReadOnly Property ExecutedCount As Integer
                Get
                    Dim count As Integer = 0

                    For Each item In Calls
                        If item IsNot Nothing AndAlso item.WillExecute Then
                            count += 1
                        End If
                    Next

                    Return count
                End Get
            End Property

            Public ReadOnly Property DeferredCount As Integer
                Get
                    Dim count As Integer = 0

                    For Each item In Calls
                        If item IsNot Nothing AndAlso Not item.WillExecute Then
                            count += 1
                        End If
                    Next

                    Return count
                End Get
            End Property

            Public ReadOnly Property IsFullyBatchSafe As Boolean
                Get
                    If Calls.Count = 0 Then Return False

                    For Each item In Calls
                        If item Is Nothing Then Return False
                        If item.IsBarrier Then Return False
                        If Not item.WillExecute Then Return False
                    Next

                    Return True
                End Get
            End Property

        End Class

        Public NotInheritable Class ToolingRunState
            Public Property HasUnresolvedToolFailure As Boolean
            Public Property LastToolName As String
            Public Property LastErrorCode As String
            Public Property LastErrorMessage As String
            Public Property LastFailureSkippedByPolicy As Boolean
            Public Property LastFailureReturnedToParent As Boolean
            Public Property LastFailureRecoveredByToolCall As Boolean
            Public Property LastFailureHandledByBlockedFinal As Boolean
            Public Property LastFailureUltimatelyFatal As Boolean
            Public Property RecoveryToolName As String

            Public Property ActiveToolingSession As Boolean
            Public Property HasOpenToolWorkflow As Boolean
            Public Property LastStateFilePath As String
            Public Property LastOutputPath As String
            Public Property LastCollectionSize As Integer?
            Public Property LastProcessedItemCount As Integer?
            Public Property LastSuccessfulToolCall As String
            Public Property LastMutationToolCall As String
            Public Property LastAgentToolCall As String
            Public Property LastReadOnlyStateToolCall As String
            Public Property LastDetectedTurnType As String
            Public Property LastInvalidTurnReason As String
            Public Property FinalResponseOrigin As String
            Public Property ToolRequiredModeUsed As Boolean

            Public Property UserLanguage As String
            Public Property LastStructuredToolResult As String
            Public Property LastStructuredToolResultKind As String
            Public Property LastStructuredToolName As String
            Public Property LastKnownOutputReference As String

            Public Property MemoryGroundingMode As MemoryGroundingMode
            Public Property ShouldExposeRecentMemoryStubs As Boolean
            Public Property MemoryListCalledThisTurn As Boolean
            Public Property MemoryGetCalledThisTurn As Boolean
            Public Property FullMemoryValueAvailableThisTurn As Boolean
            Public Property MemoryListReturnedNoEntriesThisTurn As Boolean
            Public Property FinalCompleteRejectedForMissingMemoryAccess As Boolean

            Public ReadOnly Property RequiresParentRecovery As Boolean
                Get
                    Return HasUnresolvedToolFailure AndAlso
                   LastFailureSkippedByPolicy AndAlso
                   LastFailureReturnedToParent
                End Get
            End Property

            Public Sub NoteToolFailure(toolName As String,
                               Optional errorCode As String = "",
                               Optional errorMessage As String = "",
                               Optional skippedByPolicy As Boolean = False,
                               Optional returnedToParent As Boolean = False)
                HasUnresolvedToolFailure = True
                LastToolName = If(toolName, "")
                LastErrorCode = If(errorCode, "")
                LastErrorMessage = If(errorMessage, "")
                LastFailureSkippedByPolicy = skippedByPolicy
                LastFailureReturnedToParent = returnedToParent
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub

            Public Sub NoteRecoveryByLaterToolCall(toolName As String)
                If Not HasUnresolvedToolFailure Then Return

                HasUnresolvedToolFailure = False
                LastFailureRecoveredByToolCall = True
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = If(toolName, "")
            End Sub

            Public Sub NoteBlockedFinalHandled()
                If Not HasUnresolvedToolFailure Then Return

                HasUnresolvedToolFailure = False
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = True
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub

            Public Sub NoteFailureFatal()
                If Not HasUnresolvedToolFailure Then Return
                LastFailureUltimatelyFatal = True
            End Sub

            Public Sub NoteSuccessfulProgress()
                HasUnresolvedToolFailure = False
                LastToolName = ""
                LastErrorCode = ""
                LastErrorMessage = ""
                LastFailureSkippedByPolicy = False
                LastFailureReturnedToParent = False
                LastFailureRecoveredByToolCall = False
                LastFailureHandledByBlockedFinal = False
                LastFailureUltimatelyFatal = False
                RecoveryToolName = ""
            End Sub
        End Class

        Public Shared Function BuildExecutionPlan(toolNames As IEnumerable(Of String)) As ToolBatchPlan
            Dim plan As New ToolBatchPlan()

            If toolNames Is Nothing Then
                Return plan
            End If

            Dim barrierReached As Boolean = False
            Dim index As Integer = 0

            For Each rawName In toolNames
                Dim toolName As String = If(rawName, "").Trim()
                Dim classification = ClassifyToolName(toolName)
                Dim isBarrier As Boolean = IsBarrierClassification(classification)

                Dim item As New PlannedToolCall With {
                    .Index = index,
                    .ToolName = toolName,
                    .Classification = classification,
                    .IsBarrier = isBarrier,
                    .WillExecute = Not barrierReached,
                    .SkipReason = ""
                }

                If barrierReached Then
                    item.SkipReason = "deferred_after_sequencing_barrier"
                End If

                plan.Calls.Add(item)

                If isBarrier Then
                    barrierReached = True
                End If

                index += 1
            Next

            Return plan
        End Function

        Public Shared Function ClassifyToolName(toolName As String) As ToolCallClassification
            If String.IsNullOrWhiteSpace(toolName) Then
                Return ToolCallClassification.Unknown
            End If

            Dim name As String = toolName.Trim().ToLowerInvariant()

            If name.StartsWith("agent_", StringComparison.Ordinal) Then
                Return ToolCallClassification.Agent
            End If

            If name.StartsWith("skill_", StringComparison.Ordinal) OrElse
               name.Equals("skill_use", StringComparison.Ordinal) Then
                Return ToolCallClassification.Skill
            End If

            If name.Equals("tool_loader", StringComparison.Ordinal) OrElse
               name.StartsWith("memory_", StringComparison.Ordinal) Then
                Return ToolCallClassification.Stateful
            End If

            If HasAnyPhrase(name, "make_dir", "mkdir", "rmdir") Then
                Return ToolCallClassification.Mutating
            End If

            If HasAnyToken(name,
                           "state", "session", "cursor", "next", "queue", "loader") Then
                Return ToolCallClassification.Stateful
            End If

            If HasAnyToken(name,
                           "write", "save", "create", "delete", "remove", "move", "rename", "copy",
                           "append", "insert", "update", "set", "put", "apply", "stage",
                           "download", "upload", "commit", "send", "post") Then
                Return ToolCallClassification.Mutating
            End If

            If HasAnyToken(name,
                           "read", "get", "list", "inventory", "search", "find",
                           "extract", "query", "lookup", "retrieve", "fetch", "inspect") Then
                Return ToolCallClassification.ReadOnlyIndependent
            End If

            Return ToolCallClassification.Unknown
        End Function

        Public Shared Function IsBarrierClassification(classification As ToolCallClassification) As Boolean
            Select Case classification
                Case ToolCallClassification.ReadOnlyIndependent
                    Return False
                Case Else
                    Return True
            End Select
        End Function

        Public Shared Function ShouldBlockTextOnlyFinalization(runState As ToolingRunState,
                                                               retryCount As Integer,
                                                               maxRetryCount As Integer,
                                                               hasValidFinalAnswer As Boolean) As Boolean
            If retryCount < maxRetryCount Then
                Return False
            End If

            If runState IsNot Nothing AndAlso runState.HasUnresolvedToolFailure Then
                Return True
            End If

            Return Not hasValidFinalAnswer
        End Function

        Public Shared Function BuildBlockedResultPayload(errorCode As String,
                                                         phase As String,
                                                         message As String,
                                                         Optional lastToolName As String = "",
                                                         Optional lastToolErrorCode As String = "",
                                                         Optional lastToolErrorMessage As String = "") As String
            Dim obj As New JObject(
                New JProperty("status", "blocked"),
                New JProperty("error", New JObject(
                    New JProperty("code", If(errorCode, "")),
                    New JProperty("phase", If(phase, "")),
                    New JProperty("message", If(message, "")))))

            If Not String.IsNullOrWhiteSpace(lastToolName) OrElse
               Not String.IsNullOrWhiteSpace(lastToolErrorCode) OrElse
               Not String.IsNullOrWhiteSpace(lastToolErrorMessage) Then

                Dim lastTool As New JObject()

                If Not String.IsNullOrWhiteSpace(lastToolName) Then
                    lastTool("name") = lastToolName
                End If

                If Not String.IsNullOrWhiteSpace(lastToolErrorCode) Then
                    lastTool("errorCode") = lastToolErrorCode
                End If

                If Not String.IsNullOrWhiteSpace(lastToolErrorMessage) Then
                    lastTool("message") = lastToolErrorMessage
                End If

                obj("lastToolFailure") = lastTool
            End If

            Return obj.ToString(Formatting.None)
        End Function

        Public Shared Function ParseStrictTaskStatus(text As String) As TaskStatusParseResult
            Dim result As New TaskStatusParseResult() With {
                .Status = TaskStatusKind.None
            }

            Dim trimmedText As String = If(text, "")
            Dim trimmedEnd As String = trimmedText.TrimEnd()

            If trimmedEnd = "" Then
                result.FailureReason = "empty_response"
                Return result
            End If

            Dim matches As MatchCollection =
                Regex.Matches(
                    trimmedEnd,
                    "<TASK_STATUS>\s*(\{.*?\})\s*</TASK_STATUS>",
                    RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

            result.FooterCount = matches.Count

            If matches.Count = 0 Then
                result.FailureReason = "missing_task_status"
                Return result
            End If

            result.IsPresent = True

            If matches.Count <> 1 Then
                result.FailureReason = "multiple_task_status"
                Return result
            End If

            Dim match As Match = matches(0)

            If match.Index + match.Length <> trimmedEnd.Length Then
                result.FailureReason = "task_status_not_at_end"
                Return result
            End If

            Dim jsonText As String = match.Groups(1).Value.Trim()
            result.FooterJson = jsonText
            result.TextBeforeFooter = trimmedEnd.Substring(0, match.Index).TrimEnd()

            Try
                Dim obj As JObject = JObject.Parse(jsonText)
                Dim statusText As String = If(obj.Value(Of String)("status"), "").Trim().ToLowerInvariant()
                Dim reasonText As String = If(obj.Value(Of String)("reason"), "").Trim()

                If reasonText = "" Then
                    result.FailureReason = "task_status_missing_reason"
                    Return result
                End If

                If reasonText.Length > 160 Then
                    result.FailureReason = "task_status_reason_too_long"
                    Return result
                End If

                Select Case statusText
                    Case "complete"
                        result.Status = TaskStatusKind.Complete
                    Case "blocked"
                        result.Status = TaskStatusKind.Blocked
                    Case "continue"
                        result.Status = TaskStatusKind.ContinueTurn
                    Case Else
                        result.FailureReason = "task_status_invalid_status"
                        Return result
                End Select

                result.Reason = reasonText
                result.IsValid = True
                Return result
            Catch
                result.FailureReason = "malformed_task_status"
                Return result
            End Try
        End Function


        Public NotInheritable Class MemoryGroundingIntentDecision
            Public Property MemoryGroundingMode As MemoryGroundingMode = MemoryGroundingMode.None
            Public Property Reason As String = "invalid_classifier_output"
            Public Property ShouldExposeRecentMemoryStubs As Boolean
            Public Property IsValid As Boolean
        End Class

        Public Shared Function BuildMemoryGroundingIntentClassifierSystemPrompt() As String
            Return "Classify whether the assistant's next answer should be grounded in session memory or prior stored workflow results. " &
                "Return EXACTLY one raw JSON object and nothing else. " &
                "Do NOT use Markdown. Do NOT use code fences. Do NOT add explanations. Do NOT add surrounding text. " &
                "The output must be exactly one JSON object with exactly these fields: " &
                "{""memoryGroundingMode"":""none|optional|required"",""reason"":""short reason"",""shouldExposeRecentMemoryStubs"":true|false}. " &
                "Use ""required"" when the user explicitly requires an answer based on Memory, stored results, remembered results, prior saved information, or previous workflow outputs. " &
                "Use ""optional"" when prior stored workflow context may help but is not explicitly required. " &
                "Use ""none"" for an unrelated new task. " &
                "Base the decision on semantic meaning, not on language-specific keywords or surface wording."
        End Function

        Public Shared Function BuildMemoryGroundingIntentClassifierUserPrompt(userRequest As String) As String
            Return "<USER_REQUEST>" & If(userRequest, "").Trim() & "</USER_REQUEST>"
        End Function

        Public Shared Function ParseMemoryGroundingIntentClassifierDecision(raw As String) As MemoryGroundingIntentDecision
            Dim normalizedOutput As String = ""
            Dim parseError As String = ""
            Return ParseMemoryGroundingIntentClassifierDecision(raw, normalizedOutput, parseError)
        End Function

        Public Shared Function ParseMemoryGroundingIntentClassifierDecision(raw As String,
                                                                            ByRef normalizedOutput As String,
                                                                            ByRef parseError As String) As MemoryGroundingIntentDecision
            Dim result As New MemoryGroundingIntentDecision()

            normalizedOutput = NormalizeMemoryGroundingIntentClassifierOutput(raw)
            parseError = ""

            If String.IsNullOrWhiteSpace(normalizedOutput) Then
                parseError = "empty_classifier_output"
                Return result
            End If

            Try
                Dim obj As JObject = JObject.Parse(normalizedOutput)

                Dim parsedMode As MemoryGroundingMode
                If Not TryParseMemoryGroundingModeText(
                    If(obj.Value(Of String)("memoryGroundingMode"), ""),
                    parsedMode) Then
                    parseError = "invalid_memory_grounding_mode"
                    Return result
                End If

                Dim reasonToken As JToken = obj("reason")
                If reasonToken Is Nothing OrElse reasonToken.Type <> JTokenType.String Then
                    parseError = "missing_or_invalid_reason"
                    Return result
                End If

                Dim exposeToken As JToken = obj("shouldExposeRecentMemoryStubs")
                If exposeToken Is Nothing OrElse exposeToken.Type <> JTokenType.Boolean Then
                    parseError = "missing_or_invalid_shouldExposeRecentMemoryStubs"
                    Return result
                End If

                result.MemoryGroundingMode = parsedMode
                result.Reason = reasonToken.Value(Of String)().Trim()
                If result.Reason = "" Then
                    result.Reason = "parsed_classifier_output"
                End If

                result.ShouldExposeRecentMemoryStubs = exposeToken.Value(Of Boolean)()
                result.IsValid = True
                Return result
            Catch ex As Exception
                parseError = ex.Message
                Return result
            End Try
        End Function

        Private Shared Function NormalizeMemoryGroundingIntentClassifierOutput(raw As String) As String
            Dim trimmed As String = If(raw, "").Trim()
            If trimmed = "" Then
                Return ""
            End If

            Dim fencedMatch As Match =
                Regex.Match(
                    trimmed,
                    "^\s*```(?:[A-Za-z0-9_-]+)?\s*\r?\n(?<body>[\s\S]*?)\r?\n```\s*$",
                    RegexOptions.CultureInvariant)

            If fencedMatch.Success Then
                Return fencedMatch.Groups("body").Value.Trim()
            End If

            Return trimmed
        End Function

        Public Shared Function FormatMemoryGroundingMode(mode As MemoryGroundingMode) As String
            Select Case mode
                Case MemoryGroundingMode.Required
                    Return "required"
                Case MemoryGroundingMode.OptionalMode
                    Return "optional"
                Case Else
                    Return "none"
            End Select
        End Function

        Private Shared Function TryParseMemoryGroundingModeText(value As String,
                                                                ByRef mode As MemoryGroundingMode) As Boolean
            Select Case If(value, "").Trim().ToLowerInvariant()
                Case "required"
                    mode = MemoryGroundingMode.Required
                    Return True
                Case "optional"
                    mode = MemoryGroundingMode.OptionalMode
                    Return True
                Case "none"
                    mode = MemoryGroundingMode.None
                    Return True
                Case Else
                    mode = MemoryGroundingMode.None
                    Return False
            End Select
        End Function


        Public Shared Function ValidateActiveToolingTurn(responseText As String,
                                                         hasToolCalls As Boolean,
                                                         hasUnresolvedToolFailure As Boolean,
                                                         Optional runState As ToolingRunState = Nothing) As ActiveToolingTurnValidationResult
            Dim result As New ActiveToolingTurnValidationResult() With {
                .TurnKind = ActiveToolingTurnKind.InvalidTurn,
                .InvalidReason = "",
                .TaskStatus = Nothing
            }

            If hasToolCalls Then
                result.TurnKind = ActiveToolingTurnKind.ToolCallTurn
                Return result
            End If

            If String.IsNullOrWhiteSpace(responseText) Then
                result.InvalidReason = "empty_response"
                Return result
            End If

            If IsRawInternalJsonResponse(responseText) Then
                result.InvalidReason = "raw_internal_json"
                Return result
            End If

            Dim parsed As TaskStatusParseResult = ParseStrictTaskStatus(responseText)
            result.TaskStatus = parsed

            If Not parsed.IsPresent Then
                result.InvalidReason = parsed.FailureReason
                Return result
            End If

            If Not parsed.IsValid Then
                result.InvalidReason = parsed.FailureReason
                Return result
            End If

            If String.IsNullOrWhiteSpace(parsed.TextBeforeFooter) Then
                result.InvalidReason = "missing_user_facing_text"
                Return result
            End If

            Select Case parsed.Status
                Case TaskStatusKind.Complete
                    If hasUnresolvedToolFailure Then
                        result.InvalidReason = "complete_with_unresolved_tool_failure"
                        Return result
                    End If

                    If Not IsRequiredMemoryGroundingSatisfied(runState, ActiveToolingTurnKind.FinalCompleteTurn) Then
                        result.InvalidReason = MissingRequiredMemoryAccessCode
                        Return result
                    End If

                    result.TurnKind = ActiveToolingTurnKind.FinalCompleteTurn
                    Return result

                Case TaskStatusKind.Blocked
                    result.TurnKind = ActiveToolingTurnKind.FinalBlockedTurn
                    Return result

                Case TaskStatusKind.ContinueTurn
                    result.InvalidReason = "task_status_continue_not_final"
                    Return result

                Case Else
                    result.InvalidReason = "invalid_turn"
                    Return result
            End Select
        End Function

        Public Shared Function BuildActiveToolingRepairPrompt(Optional runState As ToolingRunState = Nothing) As String
            Dim prompt As String =
                "REPAIR: Your previous response was invalid for an active tooling session. " &
                "Your next response must be EXACTLY one of: " &
                "(1) the next required tool call and nothing else; " &
                "(2) a user-facing final prose answer ending with exactly one valid <TASK_STATUS>{""status"":""complete"",""reason"":""...""}</TASK_STATUS>; or " &
                "(3) a user-facing blocked explanation ending with exactly one valid <TASK_STATUS>{""status"":""blocked"",""reason"":""...""}</TASK_STATUS>. " &
                "Progress narration is invalid. Announcements of future work are invalid. Final prose without exactly one valid TASK_STATUS footer is invalid. TASK_STATUS continue is invalid during active tooling. Raw internal JSON is invalid as a user-facing answer."

            If runState Is Nothing Then
                Return prompt
            End If

            Dim additions As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(runState.LastStructuredToolResult) Then
                Dim toolLabel As String =
                    If(String.IsNullOrWhiteSpace(runState.LastStructuredToolName),
                       "the latest successful tool",
                       "'" & runState.LastStructuredToolName & "'")

                additions.Add(
                    "The latest successful structured tool result from " & toolLabel & " remains available during repair. " &
                    "Do not discard it, do not narrate it as progress, and do not surface it as raw JSON.")
            End If

            If Not String.IsNullOrWhiteSpace(runState.LastKnownOutputReference) Then
                additions.Add("A generic output or state reference remains available: " & runState.LastKnownOutputReference & ".")
            End If

            If additions.Count = 0 Then
                Return prompt
            End If

            Return prompt & " " & String.Join(" ", additions)
        End Function


        Public Shared Sub NoteMemoryGroundingToolResult(runState As ToolingRunState,
                                                        toolName As String,
                                                        rawResponse As String,
                                                        succeeded As Boolean)
            If runState Is Nothing OrElse String.IsNullOrWhiteSpace(toolName) Then
                Return
            End If

            Select Case toolName.Trim().ToLowerInvariant()
                Case MemoryTools.ToolList
                    runState.MemoryListCalledThisTurn = True
                    runState.MemoryListReturnedNoEntriesThisTurn =
                        succeeded AndAlso MemoryListHasNoEntries(rawResponse)

                Case MemoryTools.ToolGet
                    runState.MemoryGetCalledThisTurn = True

                    If succeeded AndAlso MemoryGetReturnedFullValue(rawResponse) Then
                        runState.FullMemoryValueAvailableThisTurn = True
                    End If
            End Select
        End Sub

        Public Shared Function IsRequiredMemoryGroundingSatisfied(runState As ToolingRunState,
                                                                  proposedTurnKind As ActiveToolingTurnKind) As Boolean
            If proposedTurnKind <> ActiveToolingTurnKind.FinalCompleteTurn Then
                Return True
            End If

            If runState Is Nothing Then
                Return True
            End If

            If runState.MemoryGroundingMode <> MemoryGroundingMode.Required Then
                Return True
            End If

            If runState.FullMemoryValueAvailableThisTurn Then
                Return True
            End If

            If runState.MemoryListCalledThisTurn AndAlso runState.MemoryListReturnedNoEntriesThisTurn Then
                Return True
            End If

            Return False
        End Function

        Public Shared Function BuildRequiredMemoryGroundingRepairPrompt() As String
            Return "The user requested an answer based on Memory. You have not accessed Memory in this turn. Call memory_list or memory_get before finalizing, or return a valid blocked response explaining why Memory cannot be accessed. In required memory-grounding mode, memory_list may be used to discover entries. If the answer requires the stored full content rather than only the list summaries, call memory_get for the relevant entries before finalizing."
        End Function

        Public Shared Function BuildMemoryGroundingStateSummary(runState As ToolingRunState) As String
            If runState Is Nothing Then
                Return "memoryGroundingMode=none; shouldExposeRecentMemoryStubs=false; memoryListCalledThisTurn=false; memoryGetCalledThisTurn=false; fullMemoryValueAvailableThisTurn=false; FinalCompleteRejectedForMissingMemoryAccess=false"
            End If

            Return "memoryGroundingMode=" & FormatMemoryGroundingMode(runState.MemoryGroundingMode) & ";" &
                " shouldExposeRecentMemoryStubs=" & If(runState.ShouldExposeRecentMemoryStubs, "true", "false") & ";" &
                " memoryListCalledThisTurn=" & If(runState.MemoryListCalledThisTurn, "true", "false") & ";" &
                " memoryGetCalledThisTurn=" & If(runState.MemoryGetCalledThisTurn, "true", "false") & ";" &
                " fullMemoryValueAvailableThisTurn=" & If(runState.FullMemoryValueAvailableThisTurn, "true", "false") & ";" &
                " FinalCompleteRejectedForMissingMemoryAccess=" & If(runState.FinalCompleteRejectedForMissingMemoryAccess, "true", "false")
        End Function

        Private Shared Function MemoryListHasNoEntries(rawResponse As String) As Boolean
            Dim trimmed As String = If(rawResponse, "").Trim()

            If trimmed = "" Then
                Return False
            End If

            If trimmed = "[]" Then
                Return True
            End If

            Try
                Dim token As JToken = JToken.Parse(trimmed)

                If TypeOf token Is JArray Then
                    Return DirectCast(token, JArray).Count = 0
                End If

                Dim obj As JObject = TryCast(token, JObject)
                If obj Is Nothing Then
                    Return False
                End If

                For Each propertyName In New String() {"items", "entries", "results"}
                    Dim child As JToken = obj(propertyName)

                    If TypeOf child Is JArray Then
                        Return DirectCast(child, JArray).Count = 0
                    End If
                Next
            Catch
            End Try

            Return False
        End Function

        Private Shared Function MemoryGetReturnedFullValue(rawResponse As String) As Boolean
            Dim trimmed As String = If(rawResponse, "").Trim()

            If trimmed = "" Then
                Return False
            End If

            Try
                Dim obj As JObject = TryCast(JToken.Parse(trimmed), JObject)
                If obj Is Nothing Then
                    Return False
                End If

                Dim valueToken As JToken = obj("value")
                Return valueToken IsNot Nothing AndAlso valueToken.Type <> JTokenType.Null
            Catch
                Return False
            End Try
        End Function


        Public Shared Sub NoteToolResultForRepair(runState As ToolingRunState,
                                                  toolName As String,
                                                  responseText As String,
                                                  Optional resultKind As String = "")
            If runState Is Nothing Then Return

            Dim raw As String = If(responseText, "").Trim()
            If raw = "" Then Return

            Dim normalizedKind As String = If(resultKind, "").Trim()
            If String.Equals(normalizedKind, "error", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            Try
                Dim token As JToken = JToken.Parse(raw)

                If Not TypeOf token Is JObject AndAlso Not TypeOf token Is JArray Then
                    Return
                End If

                runState.LastStructuredToolResult = raw
                runState.LastStructuredToolName = If(toolName, "")

                If normalizedKind = "" Then
                    normalizedKind = If(TypeOf token Is JObject, "json_object", "json_array")
                End If

                runState.LastStructuredToolResultKind = normalizedKind

                If TypeOf token Is JObject Then
                    TryNoteStructuredOutputReference(runState, DirectCast(token, JObject))
                End If
            Catch
                If normalizedKind <> "" AndAlso
                   Not String.Equals(normalizedKind, "text", StringComparison.OrdinalIgnoreCase) Then

                    runState.LastStructuredToolResult = raw
                    runState.LastStructuredToolName = If(toolName, "")
                    runState.LastStructuredToolResultKind = normalizedKind
                End If
            End Try
        End Sub

        Private Shared Sub TryNoteStructuredOutputReference(runState As ToolingRunState,
                                                            payload As JObject)
            If runState Is Nothing OrElse payload Is Nothing Then Return

            Dim reference As String =
                ExtractFirstStringValue(
                    payload,
                    "output_reference",
                    "state_reference",
                    "reference",
                    "output_path",
                    "state_path",
                    "path",
                    "file_path",
                    "workspace_path",
                    "memory_key",
                    "stub")

            If String.IsNullOrWhiteSpace(reference) Then
                reference =
                    ExtractFirstStringValue(
                        TryCast(payload("result"), JObject),
                        "output_reference",
                        "state_reference",
                        "reference",
                        "output_path",
                        "state_path",
                        "path",
                        "file_path",
                        "workspace_path",
                        "memory_key",
                        "stub")
            End If

            If Not String.IsNullOrWhiteSpace(reference) Then
                runState.LastKnownOutputReference = reference
            End If
        End Sub

        Private Shared Function ExtractFirstStringValue(payload As JObject,
                                                        ParamArray keys() As String) As String
            If payload Is Nothing OrElse keys Is Nothing Then Return ""

            For Each key In keys
                If String.IsNullOrWhiteSpace(key) Then Continue For

                Dim token As JToken = payload(key)
                If token Is Nothing Then Continue For

                Dim value As String = token.ToString().Trim()
                If value <> "" Then
                    Return value
                End If
            Next

            Return ""
        End Function


        Public Shared Sub NoteToolExecutionMetadata(runState As ToolingRunState,
                                                    toolName As String,
                                                    arguments As IDictionary(Of String, Object),
                                                    success As Boolean)
            If runState Is Nothing Then Return

            runState.ActiveToolingSession = True
            runState.HasOpenToolWorkflow = True

            Dim classification As ToolCallClassification = ClassifyToolName(toolName)

            If success Then
                runState.LastSuccessfulToolCall = If(toolName, "")
            End If

            Select Case classification
                Case ToolCallClassification.Mutating
                    runState.LastMutationToolCall = If(toolName, "")
                Case ToolCallClassification.Agent
                    runState.LastAgentToolCall = If(toolName, "")
                Case ToolCallClassification.ReadOnlyIndependent, ToolCallClassification.Stateful
                    runState.LastReadOnlyStateToolCall = If(toolName, "")
            End Select

            Dim knownPath As String = ExtractFirstPathArgument(arguments)
            If Not String.IsNullOrWhiteSpace(knownPath) Then
                runState.LastKnownOutputReference = knownPath
                runState.LastStateFilePath = knownPath

                If classification = ToolCallClassification.Mutating Then
                    runState.LastOutputPath = knownPath
                End If
            End If

            Dim collectionSize As Integer? = InferCollectionSize(arguments)
            If collectionSize.HasValue Then
                runState.LastCollectionSize = collectionSize
            End If

            If success AndAlso runState.LastCollectionSize.HasValue AndAlso runState.LastCollectionSize.Value > 1 Then
                runState.LastProcessedItemCount = If(runState.LastProcessedItemCount, 0) + 1
            End If
        End Sub

        Public Shared Function BuildTaskStatusFooter(status As String, reason As String) As String
            Dim footerObject As New JObject(
                New JProperty("status", If(status, "").Trim().ToLowerInvariant()),
                New JProperty("reason", NormalizeFooterReason(reason)))

            Return "<TASK_STATUS>" & footerObject.ToString(Formatting.None) & "</TASK_STATUS>"
        End Function

        Public Shared Function BuildUserSafeBlockedFinalMessage(runState As ToolingRunState,
                                                                errorCode As String,
                                                                message As String,
                                                                successCount As Integer,
                                                                failedCount As Integer,
                                                                Optional userLanguage As String = "",
                                                                Optional appendTaskStatusFooter As Boolean = True) As String
            Dim effectiveUserLanguage As String = If(userLanguage, "").Trim()
            If effectiveUserLanguage = "" Then
                effectiveUserLanguage = If(runState?.UserLanguage, "").Trim()
            End If

            Dim languageKey As String =
                ResolveBlockedMessageLanguage(effectiveUserLanguage)

            Dim parts As New List(Of String)()

            parts.Add(
                GetBlockedMessageText(
                    languageKey,
                    "The tool workflow stopped before a valid final answer was produced.",
                    "Der Tool-Ablauf wurde beendet, bevor eine gültige Schlussantwort erzeugt wurde."))

            If String.Equals(errorCode, MissingRequiredMemoryAccessCode, StringComparison.OrdinalIgnoreCase) Then
                parts.Add(
                    GetBlockedMessageText(
                        languageKey,
                        "Memory access was required before finalization, but no relevant Memory content was accessed in this run.",
                        "Vor der Finalisierung war ein Zugriff auf Memory erforderlich, aber in diesem Lauf wurde kein relevanter Memory-Inhalt geladen."))
            End If

            If runState IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(runState.LastSuccessfulToolCall) Then
                parts.Add(
                    String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        GetBlockedMessageText(
                            languageKey,
                            "Last successful tool call: {0}.",
                            "Letzter erfolgreicher Tool-Aufruf: {0}."),
                        runState.LastSuccessfulToolCall))
            End If

            If runState IsNot Nothing AndAlso runState.HasUnresolvedToolFailure Then
                Dim failureText As String =
                    GetBlockedMessageText(
                        languageKey,
                        "An unresolved tool failure remains",
                        "Ein ungelöster Tool-Fehler besteht weiterhin")

                If Not String.IsNullOrWhiteSpace(runState.LastToolName) Then
                    failureText &= String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        GetBlockedMessageText(languageKey, " in {0}", " in {0}"),
                        runState.LastToolName)
                End If

                If Not String.IsNullOrWhiteSpace(runState.LastErrorCode) Then
                    failureText &= String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        GetBlockedMessageText(languageKey, " ({0})", " ({0})"),
                        runState.LastErrorCode)
                End If

                If Not String.IsNullOrWhiteSpace(runState.LastErrorMessage) Then
                    failureText &= String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        GetBlockedMessageText(languageKey, ": {0}", ": {0}"),
                        runState.LastErrorMessage)
                End If

                parts.Add(failureText & ".")
            Else
                parts.Add(
                    GetBlockedMessageText(
                        languageKey,
                        "No unresolved tool failure was recorded before response-contract enforcement stopped the run.",
                        "Vor dem Stopp durch die Antwortvertrags-Prüfung wurde kein ungelöster Tool-Fehler erfasst."))
            End If

            parts.Add(
                String.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    GetBlockedMessageText(
                        languageKey,
                        "Successful tool-call count: {0}. Failed tool-call count: {1}. These are tool-call counts only; they do not indicate completed business items.",
                        "Erfolgreiche Tool-Aufruf-Zahl: {0}. Fehlgeschlagene Tool-Aufruf-Zahl: {1}. Diese Zahlen betreffen nur Tool-Aufrufe und bedeuten keine abgeschlossenen Arbeitseinheiten."),
                    Math.Max(successCount, 0),
                    Math.Max(failedCount, 0)))

            Dim referenceText As String = GetLastKnownReference(runState)
            If referenceText <> "" Then
                parts.Add(
                    String.Format(
                        System.Globalization.CultureInfo.InvariantCulture,
                        GetBlockedMessageText(
                            languageKey,
                            "Partial output or state reference: {0}.",
                            "Teilweise Ausgabe- oder Statusreferenz: {0}."),
                        referenceText))
            End If

            Dim hint As String = BuildPracticalHint(runState, languageKey)
            If hint <> "" Then
                parts.Add(
                    GetBlockedMessageText(
                        languageKey,
                        "Suggested next action: ",
                        "Empfohlener nächster Schritt: ") & hint)
            End If

            Dim finalText As String =
                String.Join(" ", parts.Where(Function(p) Not String.IsNullOrWhiteSpace(p))).Trim()

            If appendTaskStatusFooter Then
                finalText &= " " & BuildTaskStatusFooter("blocked", If(errorCode, "host_generated_blocked"))
            End If

            Return finalText.Trim()
        End Function

        Private Shared Function IsRawInternalJsonResponse(text As String) As Boolean
            Dim trimmed As String = If(text, "").Trim()
            If trimmed = "" Then Return False

            Try
                Dim token As JToken = JToken.Parse(trimmed)
                Dim obj As JObject = TryCast(token, JObject)
                If obj Is Nothing Then Return False

                Return obj("status") IsNot Nothing OrElse
                       obj("error") IsNot Nothing OrElse
                       obj("resultKind") IsNot Nothing
            Catch
                Return False
            End Try
        End Function

        Private Shared Function NormalizeFooterReason(reason As String) As String
            Dim normalized As String = If(reason, "").Trim()
            If normalized = "" Then normalized = "blocked"
            If normalized.Length > 80 Then normalized = normalized.Substring(0, 80).Trim()
            Return normalized
        End Function

        Private Shared Function ResolveBlockedMessageLanguage(userLanguage As String) As String
            Dim normalized As String = If(userLanguage, "").Trim()
            If normalized = "" Then Return "en"

            Try
                Dim culture = System.Globalization.CultureInfo.GetCultureInfo(normalized)
                If String.Equals(culture.TwoLetterISOLanguageName, "de", StringComparison.OrdinalIgnoreCase) Then
                    Return "de"
                End If
            Catch
                If normalized.StartsWith("de", StringComparison.OrdinalIgnoreCase) Then
                    Return "de"
                End If
            End Try

            Return "en"
        End Function

        Private Shared Function GetBlockedMessageText(languageKey As String,
                                                      englishText As String,
                                                      germanText As String) As String
            If String.Equals(languageKey, "de", StringComparison.OrdinalIgnoreCase) Then
                Return germanText
            End If

            Return englishText
        End Function

        Private Shared Function GetLastKnownReference(runState As ToolingRunState) As String
            If runState Is Nothing Then Return ""

            If Not String.IsNullOrWhiteSpace(runState.LastKnownOutputReference) Then
                Return runState.LastKnownOutputReference
            End If

            If Not String.IsNullOrWhiteSpace(runState.LastOutputPath) Then
                Return runState.LastOutputPath
            End If

            If Not String.IsNullOrWhiteSpace(runState.LastStateFilePath) Then
                Return runState.LastStateFilePath
            End If

            Return ""
        End Function

        Private Shared Function BuildPracticalHint(runState As ToolingRunState,
                                                   languageKey As String) As String
            If runState Is Nothing Then
                Return GetBlockedMessageText(
                    languageKey,
                    "Retry with a smaller, deterministic next step.",
                    "Erneut mit einem kleineren, deterministischen nächsten Schritt versuchen.")
            End If

            If Not String.IsNullOrWhiteSpace(GetLastKnownReference(runState)) Then
                Return GetBlockedMessageText(
                    languageKey,
                    "Review the last known output/state reference and resume from that point.",
                    "Die letzte bekannte Ausgabe- oder Statusreferenz prüfen und von dort aus fortsetzen.")
            End If

            If Not String.IsNullOrWhiteSpace(runState.LastSuccessfulToolCall) Then
                Return GetBlockedMessageText(
                    languageKey,
                    "Retry from the last successful tool call or narrow the next requested tool step.",
                    "Vom letzten erfolgreichen Tool-Aufruf erneut ansetzen oder den nächsten angeforderten Tool-Schritt enger fassen.")
            End If

            Return GetBlockedMessageText(
                languageKey,
                "Retry with a smaller, deterministic next step.",
                "Erneut mit einem kleineren, deterministischen nächsten Schritt versuchen.")
        End Function

        Private Shared Function ExtractFirstPathArgument(arguments As IDictionary(Of String, Object)) As String
            If arguments Is Nothing Then Return ""

            Dim keys As String() = {
                "path",
                "file_path",
                "source_path",
                "target_path",
                "output_path",
                "state_path",
                "workspace_path"
            }

            For Each key In keys
                If Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then Continue For

                Dim value As String = TryGetScalarString(arguments(key))
                If Not String.IsNullOrWhiteSpace(value) Then
                    Return value.Trim()
                End If
            Next

            Return ""
        End Function

        Private Shared Function InferCollectionSize(arguments As IDictionary(Of String, Object)) As Integer?
            If arguments Is Nothing Then Return Nothing

            For Each pair In arguments
                If pair.Value Is Nothing OrElse TypeOf pair.Value Is String Then Continue For

                If TypeOf pair.Value Is JArray Then
                    Return DirectCast(pair.Value, JArray).Count
                End If

                If TypeOf pair.Value Is IEnumerable Then
                    Dim count As Integer = 0
                    For Each item In DirectCast(pair.Value, IEnumerable)
                        count += 1
                    Next
                    Return count
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function TryGetScalarString(value As Object) As String
            If value Is Nothing Then Return ""

            If TypeOf value Is JValue Then
                Return DirectCast(value, JValue).ToString()
            End If

            If TypeOf value Is String Then
                Return CStr(value)
            End If

            Return value.ToString()
        End Function

        Private Shared Function HasAnyPhrase(name As String, ParamArray phrases() As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            If phrases Is Nothing Then Return False

            For Each phrase In phrases
                If String.IsNullOrWhiteSpace(phrase) Then Continue For

                If name.IndexOf(phrase, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function HasAnyToken(name As String, ParamArray expectedTokens() As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            If expectedTokens Is Nothing Then Return False

            Dim tokens = name.Split(New Char() {"_"c, "-"c, "."c}, StringSplitOptions.RemoveEmptyEntries)

            For Each token In tokens
                For Each expected In expectedTokens
                    If String.IsNullOrWhiteSpace(expected) Then Continue For

                    If token.Equals(expected, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Next

            Return False
        End Function

    End Class

End Namespace