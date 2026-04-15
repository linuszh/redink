' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose: Wires the Knowledge Store background indexing service and foreground
'          indexer into Word's startup, idle timer, and shutdown lifecycle.
'          Also provides the ConsultKnowledgeStore function used by Freestyle
'          and DiscussInky for (kb:...) trigger resolution.
'
' Lifecycle:
'  - Startup: InitializeKnowledgeStoreService() called from DelayedStartupTasks.
'  - Idle: _ksTimer_Tick fires every 60s and drives background indexing.
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

    ''' <summary>Timer driving background Knowledge Store indexing.</summary>
    Private _ksTimer As System.Windows.Forms.Timer

    ''' <summary>Interval between idle ticks in milliseconds (60 seconds).</summary>
    Private Const KS_IDLE_INTERVAL_MS As Integer = 60000

    ''' <summary>
    ''' Initializes the Knowledge Store idle service. Call from DelayedStartupTasks.
    ''' </summary>
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
    ''' Idle timer tick — drives background indexing.
    ''' </summary>
    Private Async Sub KsTimer_Tick(sender As Object, e As EventArgs)
        Try
            If Not KnowledgeStoreIdleService.CanRunNow(_context) Then
                Debug.WriteLine("KS Wiring: Skipping tick — outside configured Knowledge Store processing window.")
                Return
            End If

            Debug.WriteLine("KS Wiring: Timer tick fired (60s).")
            Await KnowledgeStoreIdleService.OnIdleTickAsync().ConfigureAwait(False)
        Catch
            ' Never let timer callbacks crash the host
        End Try
    End Sub

    ''' <summary>
    ''' Runs a foreground index of all (or a specific) Knowledge Store with progress bar.
    ''' Called from menu or settings UI.
    ''' </summary>
    ''' <param name="storeName">Optional store name to restrict indexing to. Empty = all stores.</param>
    ''' <param name="forceReindex">If True, re-indexes all files even if already indexed.</param>
    Public Async Function RunForegroundKnowledgeStoreIndexAsync(
            Optional storeName As String = "",
            Optional forceReindex As Boolean = False) As Task
        Try
            Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(_context, storeName, forceReindex).ConfigureAwait(False)

            Dim msg As String
            If result.WasCancelled Then
                msg = $"Indexing was cancelled. Indexed {result.IndexedFiles} of {result.TotalFiles} file(s)."
            Else
                msg = $"Indexing complete. Indexed: {result.IndexedFiles}, Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}."
            End If

            ShowCustomMessageBox(msg, $"{AN} Knowledge Store")
        Catch ex As Exception
            ShowCustomMessageBox($"Error during foreground indexing: {ex.Message}", $"{AN} Knowledge Store")
        End Try
    End Function

    ''' <summary>
    ''' Shuts down the Knowledge Store service. Call from ThisAddIn_Shutdown.
    ''' </summary>
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