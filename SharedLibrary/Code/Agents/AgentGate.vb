' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: AgentGate.vb
' Purpose: Global serialization gate ensuring that no two LLM / MCP model calls
'          execute concurrently. Required because sub-agents are not yet
'          isolated; only one model interaction may run at a time.
'
' Usage:
'   Await AgentGate.EnterAsync(cancellationToken)
'   Try
'       ' ... single model call ...
'   Finally
'       AgentGate.Release()
'   End Try
'
'   The gate is re-entrant for the SAME logical owner via BeginOwnedScope /
'   EndOwnedScope. Tooling loops that already hold the gate for an outer
'   transaction can suppress nested waits within the same async flow.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    Public NotInheritable Class AgentGate

        Private Sub New()
        End Sub

        Private Shared ReadOnly _gate As New SemaphoreSlim(1, 1)

        ' Re-entrancy support. We hold the flag in a *mutable* holder so that
        ' writes performed inside awaited helpers are observable in the caller.
        ' (Plain AsyncLocal(Of Boolean) writes only flow DOWN into child async
        ' frames; they do NOT flow back UP to the calling method, which would
        ' otherwise cause the same async flow to dead-lock against itself when
        ' a sub-agent runner takes the gate and then calls into LLM().)
        Private NotInheritable Class OwnerHolder
            Public Owned As Boolean
        End Class

        Private Shared ReadOnly _ownerHolder As New AsyncLocal(Of OwnerHolder)

        Private Shared Function EnsureHolder() As OwnerHolder
            Dim h = _ownerHolder.Value
            If h Is Nothing Then
                h = New OwnerHolder()
                _ownerHolder.Value = h
            End If
            Return h
        End Function

        Private Shared Function IsOwner() As Boolean
            Dim h = _ownerHolder.Value
            Return h IsNot Nothing AndAlso h.Owned
        End Function

        ''' <summary>Acquires the global model-call gate. Honors cancellation.</summary>
        Public Shared Async Function EnterAsync(Optional cancellationToken As CancellationToken = Nothing) As Task
            Dim owner = IsOwner()
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] EnterAsync ENTER owner={owner} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
            If owner Then
                System.Diagnostics.Debug.WriteLine("[AGENTGATE] EnterAsync BYPASS (owner)")
                Return
            End If

            System.Diagnostics.Debug.WriteLine("[AGENTGATE] EnterAsync WAITING")
            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] EnterAsync ACQUIRED busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
        End Function

        ''' <summary>Releases the gate. Safe to call only if EnterAsync was awaited.</summary>
        Public Shared Sub Release()
            If IsOwner() Then Return
            Try
                _gate.Release()
            Catch
                ' Ignore over-release; defensive.
            End Try
        End Sub

        ''' <summary>
        ''' Marks the current async flow as the owner of the gate for the duration
        ''' of a nested scope. Pair with <see cref="EndOwnedScope"/> in Finally.
        ''' Useful when a tooling loop wants to hold the gate across many internal
        ''' LLM/MCP calls without serializing against itself.
        ''' </summary>
        Public Shared Async Function BeginOwnedScopeAsync(Optional cancellationToken As CancellationToken = Nothing) As Task
            Dim holder = EnsureHolder()
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] BeginOwnedScopeAsync ENTER holderExists={holder IsNot Nothing} owned={holder.Owned} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
            If holder.Owned Then
                System.Diagnostics.Debug.WriteLine("[AGENTGATE] BeginOwnedScopeAsync BYPASS (already owner)")
                Return
            End If

            System.Diagnostics.Debug.WriteLine("[AGENTGATE] BeginOwnedScopeAsync WAITING")
            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            holder.Owned = True
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] BeginOwnedScopeAsync ACQUIRED owned={holder.Owned} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
        End Function

        Public Shared Sub EndOwnedScope()
            Dim holder = _ownerHolder.Value
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] EndOwnedScope ENTER holderIsNothing={holder Is Nothing} owned={If(holder Is Nothing, False, holder.Owned)} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
            If holder Is Nothing OrElse Not holder.Owned Then
                System.Diagnostics.Debug.WriteLine("[AGENTGATE] EndOwnedScope BYPASS (not owner)")
                Return
            End If

            holder.Owned = False
            Try
                _gate.Release()
                System.Diagnostics.Debug.WriteLine($"[AGENTGATE] EndOwnedScope RELEASED busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[AGENTGATE] EndOwnedScope RELEASE ERROR: {ex.Message}")
            End Try
        End Sub

        Public Shared Sub MarkCurrentFlowAsOwner()
            Dim holder = EnsureHolder()
            holder.Owned = True
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] MarkCurrentFlowAsOwner owned={holder.Owned} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
        End Sub

        Public Shared Sub UnmarkCurrentFlowAsOwner()
            Dim holder = _ownerHolder.Value
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] UnmarkCurrentFlowAsOwner ENTER holderIsNothing={holder Is Nothing} owned={If(holder Is Nothing, False, holder.Owned)} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
            If holder Is Nothing Then Return
            holder.Owned = False
            System.Diagnostics.Debug.WriteLine($"[AGENTGATE] UnmarkCurrentFlowAsOwner EXIT owned={holder.Owned} busy={IsBusy} thread={Thread.CurrentThread.ManagedThreadId}")
        End Sub

        ''' <summary>True if the gate is currently held by some caller.</summary>
        Public Shared ReadOnly Property IsBusy As Boolean
            Get
                Return _gate.CurrentCount = 0
            End Get
        End Property

    End Class

End Namespace