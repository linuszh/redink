' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.ChatAgentScope.vb
' Purpose: Provides RAII-style scope token for temporarily elevating _chatAgentActive flag.
'          Enables built-in M365 tools during isolated tooling loops without modifying core tooling logic.
' 
' Architecture:
'  - EnterChatAgentScope(): Returns disposable token that sets _chatAgentActive = True on entry.
'  - ChatAgentScopeToken: Nested IDisposable class that restores previous flag state on Dispose.
'  - Used by ExecuteToolingLoop to isolate sub-agent M365 tool access without side effects.
' =============================================================================

Option Explicit On
Option Strict On

Imports System

Partial Public Class ThisAddIn

    Public Function EnterChatAgentScope() As IDisposable
        Return New ChatAgentScopeToken(Me)
    End Function

    Private NotInheritable Class ChatAgentScopeToken
        Implements IDisposable

        Private ReadOnly _owner As ThisAddIn
        Private ReadOnly _previous As Boolean
        Private _disposed As Boolean

        Public Sub New(owner As ThisAddIn)
            _owner = owner
            _previous = owner._chatAgentActive
            owner._chatAgentActive = True
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            Try
                _owner._chatAgentActive = _previous
            Catch
            End Try
        End Sub
    End Class

End Class