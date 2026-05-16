Option Strict On
Option Explicit On

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