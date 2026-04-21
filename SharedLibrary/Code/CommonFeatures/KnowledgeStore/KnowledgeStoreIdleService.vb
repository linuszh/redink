' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreIdleService.vb
' Purpose:
'   Hosts the shared background-indexing lifecycle used by Office hosts from
'   their idle or timer callbacks.
'
' Responsibilities:
'   - Initialize and retain the shared `KnowledgeStoreWatcher` instance.
'   - Enable or disable background indexing based on persisted settings.
'   - Trigger periodic scans and incremental queue processing during idle ticks.
'   - Keep the shared context synchronized with the current enabled state.
'   - Shut down watcher resources cleanly when the host closes.
'
' Notes:
'   - Hosts remain responsible for deciding when they are idle enough to call
'     this service.
'   - Word, Outlook, and Excel can all use this shared lifecycle service.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Globalization
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

        Private Const BackgroundEnabledSettingName As String = "EnableKBBackgroundIndexing"
        Private Const BackgroundWindowSettingName As String = "KnowledgeStoreBackgroundIndexingWindow"

        ''' <summary>
        ''' Initializes the service. Does NOT start watching unless background indexing is enabled.
        ''' </summary>
        ''' <param name="context">Shared context for configuration.</param>
        Public Shared Sub Initialize(context As ISharedContext)
            SyncLock _lock
                If _initialized Then Return

                _context = context
                _initialized = True

                If _context IsNot Nothing Then
                    _context.INI_KnowledgeStoreBackgroundIndexing = GetPersistedBackgroundEnabled(_context)
                    _context.INI_KnowledgeStoreBackgroundIndexingWindow = GetPersistedBackgroundWindow(_context)
                End If

                If _context IsNot Nothing AndAlso _context.INI_KnowledgeStoreBackgroundIndexing Then
                    StartInternal()
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' Enables or disables background indexing at runtime.
        ''' </summary>
        Public Shared Sub SetEnabled(value As Boolean)
            SyncLock _lock
                If value = _enabled Then
                    If _context IsNot Nothing Then
                        _context.INI_KnowledgeStoreBackgroundIndexing = value
                        _context.INI_KnowledgeStoreBackgroundIndexingWindow = GetPersistedBackgroundWindow(_context)
                    End If
                    Return
                End If

                If value Then
                    StartInternal()
                Else
                    StopInternal()
                End If

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

        ''' <summary>
        ''' Returns True when the configured processing window currently allows background work.
        ''' Empty configuration means always allowed.
        ''' </summary>
        Public Shared Function CanRunNow(Optional context As ISharedContext = Nothing,
                                         Optional localNow As DateTime? = Nothing) As Boolean
            Dim effectiveContext = If(context, _context)
            Dim windowSpec = GetPersistedBackgroundWindow(effectiveContext)

            If String.IsNullOrWhiteSpace(windowSpec) Then
                Return True
            End If

            Dim workingSpec = windowSpec.Trim()
            Dim allowMode As Boolean = True

            If workingSpec.StartsWith("allow:", StringComparison.OrdinalIgnoreCase) Then
                workingSpec = workingSpec.Substring("allow:".Length).Trim()
                allowMode = True
            ElseIf workingSpec.StartsWith("deny:", StringComparison.OrdinalIgnoreCase) Then
                workingSpec = workingSpec.Substring("deny:".Length).Trim()
                allowMode = False
            End If

            If String.IsNullOrWhiteSpace(workingSpec) Then
                Return True
            End If

            Dim nowValue = If(localNow.HasValue, localNow.Value, DateTime.Now)
            Dim nowTime = nowValue.TimeOfDay
            Dim parsedAny As Boolean = False
            Dim matchedAny As Boolean = False

            For Each rawPart In workingSpec.Split({";"c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                Dim part = rawPart.Trim()
                If String.IsNullOrWhiteSpace(part) Then Continue For

                Dim bounds = part.Split({"-"c}, 2, StringSplitOptions.None)
                If bounds.Length <> 2 Then Continue For

                Dim startTime As TimeSpan
                Dim endTime As TimeSpan

                If Not TryParseWindowTime(bounds(0).Trim(), startTime) Then Continue For
                If Not TryParseWindowTime(bounds(1).Trim(), endTime) Then Continue For

                parsedAny = True

                Dim isInWindow As Boolean
                If startTime = endTime Then
                    isInWindow = True
                ElseIf startTime < endTime Then
                    isInWindow = nowTime >= startTime AndAlso nowTime < endTime
                Else
                    isInWindow = nowTime >= startTime OrElse nowTime < endTime
                End If

                If isInWindow Then
                    matchedAny = True
                    Exit For
                End If
            Next

            If Not parsedAny Then
                Return True
            End If

            If allowMode Then
                Return matchedAny
            End If

            Return Not matchedAny
        End Function

        Private Shared Sub StartInternal()
            If _context Is Nothing Then Return
            If Not KnowledgeStoreCatalog.IsConfigured(_context) Then Return
            If _watcher IsNot Nothing Then Return

            Try
                _watcher = New KnowledgeStoreWatcher(_context)
                _watcher.StartWatching()
                _enabled = True

                If CanRunNow(_context) Then
                    _watcher.RunPeriodicScan()
                    _lastPeriodicScan = DateTime.UtcNow
                Else
                    _lastPeriodicScan = DateTime.MinValue
                    Debug.WriteLine("KSIdleService: Started outside allowed processing window; awaiting next allowed tick.")
                End If

                Debug.WriteLine("KSIdleService: Started")
            Catch ex As Exception
                Debug.WriteLine($"KSIdleService: Start error: {ex.Message}")
            End Try
        End Sub

        Private Shared Sub StopInternal()
            If _watcher IsNot Nothing Then
                _watcher.Dispose()
                _watcher = Nothing
                Debug.WriteLine("KSIdleService: Stopped")
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

            If Not CanRunNow(_context) Then
                Debug.WriteLine("KSIdleService: Skipping tick — outside configured processing window.")
                Return 0
            End If

            Try
                Dim mustRescanNow As Boolean =
                    HasWritableSharedStores() OrElse
                    (DateTime.UtcNow - _lastPeriodicScan).TotalMinutes >= PeriodicScanIntervalMinutes

                If mustRescanNow Then
                    Debug.WriteLine("KSIdleService: Running periodic watcher scan.")
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

        Private Shared Function HasWritableSharedStores() As Boolean
            If _context Is Nothing Then Return False

            Return KnowledgeStoreCatalog.GetActiveStores(_context).Any(
                Function(store) store IsNot Nothing AndAlso
                                KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) AndAlso
                                (store.IsFromCentralCatalog OrElse
                                 String.Equals(store.Role, "shared", StringComparison.OrdinalIgnoreCase)))
        End Function

        Private Shared Function GetPersistedBackgroundEnabled(context As ISharedContext) As Boolean
            Try
                Dim rawValue = My.Settings.Item(BackgroundEnabledSettingName)
                If rawValue IsNot Nothing Then
                    Return CBool(rawValue)
                End If
            Catch
            End Try

            If context IsNot Nothing Then
                Return context.INI_KnowledgeStoreBackgroundIndexing
            End If

            Return False
        End Function

        Private Shared Function GetPersistedBackgroundWindow(context As ISharedContext) As String
            Try
                Dim rawValue = My.Settings.Item(BackgroundWindowSettingName)
                If rawValue IsNot Nothing Then
                    Return rawValue.ToString().Trim()
                End If
            Catch
            End Try

            If context IsNot Nothing Then
                Return If(context.INI_KnowledgeStoreBackgroundIndexingWindow, "").Trim()
            End If

            Return ""
        End Function

        Private Shared Function TryParseWindowTime(value As String, ByRef result As TimeSpan) As Boolean
            Dim formats As String() = {"h\:mm", "hh\:mm", "h\:mm\:ss", "hh\:mm\:ss"}

            Return TimeSpan.TryParseExact(
                value,
                formats,
                CultureInfo.InvariantCulture,
                result)
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