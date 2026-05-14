' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillAuthorMode.vb
' Purpose: Process-wide flag + helpers for "skill author" mode. When active, the
'          agent layer is allowed to read AND write inside any discovered skill's
'          own folder (so the model can create/edit SKILL.md, references/, scripts/).
'
' Activation:
'   SkillAuthorMode.Enable()          ' persists until Disable() or process exit
'   SkillAuthorMode.Disable()
'
' Scoped activation (preferred for tool calls):
'   Using SkillAuthorMode.BeginScope()
'       ... await router.TryHandleAsync(...)
'   End Using
'
' Notes:
'   - The flag is consumed by PathPolicy (read path) and SkillAuthorPathPolicy
'     (write path) below.
'   - Active mode is reflected in the agent-layer system-prompt addendum so the
'     model is aware it may author skills.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading

Namespace Agents

    Public NotInheritable Class SkillAuthorMode

        Private Sub New()
        End Sub

        Private Shared _persistent As Integer = 0
        Private Shared ReadOnly _scope As New AsyncLocal(Of Integer)

        ''' <summary>True when the mode is active (persistent and/or in a scope).</summary>
        Public Shared ReadOnly Property IsActive As Boolean
            Get
                Return Volatile.Read(_persistent) > 0 OrElse _scope.Value > 0
            End Get
        End Property

        Public Shared Sub Enable()
            Interlocked.Increment(_persistent)
        End Sub

        Public Shared Sub Disable()
            If Volatile.Read(_persistent) > 0 Then Interlocked.Decrement(_persistent)
        End Sub

        ''' <summary>Push/pop a scope. Best when scoping per-call (chat surface).</summary>
        Public Shared Function BeginScope() As IDisposable
            _scope.Value = _scope.Value + 1
            Return New Releaser()
        End Function

        Private Class Releaser
            Implements IDisposable
            Private _done As Boolean
            Public Sub Dispose() Implements IDisposable.Dispose
                If _done Then Return
                _done = True
                _scope.Value = Math.Max(0, _scope.Value - 1)
            End Sub
        End Class

    End Class

End Namespace