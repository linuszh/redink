' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
'
' Provides a small RAII-style helper that flips the existing
' _chatAgentActive flag for the lifetime of the returned token, so the
' built-in M365 tools become available to ExecuteToolingLoop without
' modifying ThisAddIn.Tooling.vb.

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