' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose: Wires the Knowledge Store background indexing service into Outlook's
'          startup, idle timer, and shutdown lifecycle.
'
' Lifecycle:
'  - Startup: InitializeKnowledgeStoreService() called from DelayedStartupTasks.
'  - Idle: KsTimer_Tick fires every 60s and drives background indexing
'          only when Outlook is genuinely idle (no AutoPilot, Chat, or Agent).
'  - Shutdown: ShutdownKnowledgeStoreService() called from ThisAddIn_Shutdown.
'
' The user controls background indexing via My.Settings.EnableKBBackgroundIndexing
' (default: False). The setting is toggled from the Settings UI.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private _ksTimer As System.Windows.Forms.Timer
    Private Const KS_IDLE_INTERVAL_MS As Integer = 60000

    Public Sub InitializeKnowledgeStoreService()
        Try
            If Not KnowledgeStoreCatalog.IsConfigured(_context) Then Return

            KnowledgeStoreIdleService.Initialize(_context)

            _ksTimer = New System.Windows.Forms.Timer()
            _ksTimer.Interval = KS_IDLE_INTERVAL_MS
            AddHandler _ksTimer.Tick, AddressOf KsTimer_Tick
            _ksTimer.Start()
        Catch ex As Exception
            Debug.WriteLine($"KS Wiring: Init error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns True when AutoPilot is currently doing real work that should block
    ''' Knowledge Store background processing.
    ''' </summary>
    Private Function IsAutoPilotBusy() As Boolean
        If Not _apActive Then
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(_apCurrentProcessingEntryId) Then
            Return True
        End If

        If _apMailQueue.Count > 0 Then
            Return True
        End If

        If System.Threading.Interlocked.CompareExchange(_apSchedulerCheckRunning, 0, 0) <> 0 Then
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Returns True when the host is genuinely idle — no active AutoPilot mail/voicemail/scheduler
    ''' work, no chat LLM jobs, no chat agent execution, and no power transitions.
    ''' </summary>
    Private Function IsOutlookIdle() As Boolean
        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
            Return False
        End If

        If IsAutoPilotBusy() Then
            Return False
        End If

        If _chatAgentActive Then
            Return False
        End If

        If System.Threading.Interlocked.CompareExchange(activeJobs, 0, 0) > 0 Then
            Return False
        End If

        If _activeToolingContext IsNot Nothing Then
            Return False
        End If

        Return True
    End Function

    Private Async Sub KsTimer_Tick(sender As Object, e As EventArgs)
        Try
            If Not KnowledgeStoreIdleService.CanRunNow(_context) Then
                Debug.WriteLine("KS Wiring: Skipping tick — outside configured Knowledge Store processing window.")
                Return
            End If

            If Not IsOutlookIdle() Then
                Debug.WriteLine("KS Wiring: Skipping tick — Outlook is busy (active AutoPilot work, Chat, or Agent).")
                Return
            End If

            Debug.WriteLine("KS Wiring: Timer tick fired (60s).")
            Await KnowledgeStoreIdleService.OnIdleTickAsync().ConfigureAwait(False)
        Catch
            ' Never let timer callbacks crash the host
        End Try
    End Sub

    Public Sub ShutdownKnowledgeStoreService()
        Try
            If _ksTimer IsNot Nothing Then
                _ksTimer.Stop()
                RemoveHandler _ksTimer.Tick, AddressOf KsTimer_Tick
                _ksTimer.Dispose()
                _ksTimer = Nothing
            End If

            KnowledgeStoreIdleService.Shutdown()
        Catch
        End Try
    End Sub

End Class