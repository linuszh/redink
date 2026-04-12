' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreIdleService.vb
' Purpose:
'   Hosts the shared background-indexing lifecycle used by Office hosts
'   (Word, Outlook, Excel) from their idle or timer callbacks.
'
' Responsibilities:
'   - Initialize and retain the shared Knowledge Store watcher instance.
'   - Enable or disable background indexing based on persisted user settings.
'   - Trigger periodic scans and incremental processing during idle ticks.
'   - Keep the shared context synchronized with the current enabled state.
'   - Shut down watcher resources cleanly when the host closes.
'
' Notes:
'   - Hosts call `Initialize`, `SetEnabled`, `OnIdleTickAsync`, and `Shutdown`.
'   - The host remains responsible for deciding when it is truly idle enough to
'     call this service.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Singleton service that manages the Knowledge Store watcher lifecycle.
    ''' </summary>
    Public Class KnowledgeStoreIdleService

        Private Shared _watcher As KnowledgeStoreWatcher = Nothing
        Private Shared _context As ISharedContext = Nothing
        Private Shared _initialized As Boolean = False
        Private Shared _enabled As Boolean = False
        Private Shared _lastPeriodicScan As DateTime = DateTime.MinValue
        Private Shared ReadOnly _lock As New Object()

        ''' <summary>Interval between full periodic scans (minutes).</summary>
        Private Const PeriodicScanIntervalMinutes As Integer = 15

        ''' <summary>
        ''' Initializes the service. Does NOT start watching unless isEnabled = True.
        ''' </summary>
        ''' <param name="context">Shared context for configuration.</param>        
        Public Shared Sub Initialize(context As ISharedContext)
            SyncLock _lock

                If _initialized Then Return

                _context = context
                _initialized = True

                context.INI_KnowledgeStoreBackgroundIndexing = False

                If My.Settings.EnableKBBackgroundIndexing Then
                    StartInternal()
                    context.INI_KnowledgeStoreBackgroundIndexing = True
                    
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Enables or disables background indexing at runtime.
        ''' </summary>
        Public Shared Sub SetEnabled(value As Boolean)
            SyncLock _lock
                If value = _enabled Then Return

                If value Then
                    StartInternal()
                Else
                    StopInternal()
                End If

                ' Keep the live memory context perfectly in sync with the service's state
                If _context IsNot Nothing Then
                    _context.INI_KnowledgeStoreBackgroundIndexing = _enabled
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Returns True if background indexing is currently active.
        ''' </summary>
        Public Shared ReadOnly Property IsEnabled As Boolean
            Get
                Return _enabled
            End Get
        End Property

        Private Shared Sub StartInternal()
            If _context Is Nothing Then Return
            If Not KnowledgeStoreCatalog.IsConfigured(_context) Then Return
            If _watcher IsNot Nothing Then Return

            Try
                _watcher = New KnowledgeStoreWatcher(_context)
                _watcher.StartWatching()
                _watcher.RunPeriodicScan()
                _lastPeriodicScan = DateTime.UtcNow
                _enabled = True
                Debug.WriteLine($"KSIdleService: Started")
            Catch ex As Exception
                Debug.WriteLine($"KSIdleService: Start error: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub StopInternal()
            If _watcher IsNot Nothing Then
                _watcher.Dispose()
                _watcher = Nothing
                Debug.WriteLine($"KSIdleService: Stopped")
            End If
            _enabled = False
        End Sub

        ''' <summary>
        ''' Called by the host's idle timer. No-op if disabled.
        ''' Returns the number of files indexed this tick.
        ''' </summary>
        Public Shared Async Function OnIdleTickAsync() As Task(Of Integer)
            Debug.WriteLine($"KSIdleService: OnIdleTickAsync called. Enabled={_enabled}, WatcherIsNothing={_watcher Is Nothing}")
            If Not _enabled OrElse _watcher Is Nothing Then Return 0

            Try
                If (DateTime.UtcNow - _lastPeriodicScan).TotalMinutes >= PeriodicScanIntervalMinutes Then
                    Debug.WriteLine("KSIdleService: Running full periodic watcher scan.")
                    _watcher.RunPeriodicScan()
                    _lastPeriodicScan = DateTime.UtcNow
                End If

                Dim results = Await _watcher.ProcessPendingAsync().ConfigureAwait(False)
                For Each r In results
                    If r.Success Then
                        Debug.WriteLine($"KSIdleService: Indexed '{r.Title}'")
                    Else
                        Debug.WriteLine($"KSIdleService: Failed '{r.FilePath}': {r.ErrorMessage}")
                    End If
                Next

                Return Enumerable.Count(results, Function(r) r.Success)
            Catch ex As Exception
                Debug.WriteLine($"KSIdleService: Tick error: {ex.Message}")
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Returns the number of files currently queued for indexing.
        ''' </summary>
        Public Shared ReadOnly Property PendingCount As Integer
            Get
                If _watcher Is Nothing Then Return 0
                Return _watcher.PendingCount
            End Get
        End Property

        ''' <summary>
        ''' Stops watching and releases resources. Called at host shutdown.
        ''' </summary>
        Public Shared Sub Shutdown()
            SyncLock _lock
                StopInternal()
                _initialized = False
                _context = Nothing
            End SyncLock
        End Sub

    End Class

End Namespace