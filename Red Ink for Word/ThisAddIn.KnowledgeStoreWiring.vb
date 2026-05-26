' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.KnowledgeStoreWiring.vb
' Purpose:
'   Wires Knowledge Store services into the Word add-in lifecycle.
'
' Responsibilities:
'   - Initialize and shut down the shared Knowledge Store idle service.
'   - Drive periodic background indexing from a WinForms timer.
'   - Register Word-specific host-idle logic with `KnowledgeStoreHostGate` so
'     gated Knowledge Store AI work waits while Word chat activity is visible.
'   - Expose the foreground Knowledge Store indexing command used by UI entry
'     points.
'
' Lifecycle:
'   - Startup: `InitializeKnowledgeStoreService()` from delayed startup.
'   - Idle: `KsTimer_Tick` drives background indexing.
'   - Shutdown: `ShutdownKnowledgeStoreService()` clears timer and gate wiring.
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

    Private Function IsWordIdle() As Boolean
        Try
            If chatForm IsNot Nothing AndAlso
               Not chatForm.IsDisposed AndAlso
               chatForm.Visible Then
                Return False
            End If
        Catch
        End Try

        Try
            For Each openForm As System.Windows.Forms.Form In System.Windows.Forms.Application.OpenForms
                If openForm Is Nothing OrElse openForm.IsDisposed OrElse Not openForm.Visible Then
                    Continue For
                End If

                If TypeOf openForm Is frmAIChat Then
                    Return False
                End If
            Next
        Catch
        End Try

        Return True
    End Function

    Public Sub InitializeKnowledgeStoreService()
        Try
            KnowledgeStoreHostGate.RegisterHostIdleProvider("Word", Function() IsWordIdle())

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
        Finally
            KnowledgeStoreHostGate.ClearHostIdleProvider()
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


End Class