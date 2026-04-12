' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreWatcher.vb
' Purpose:
'   Monitors active, writable Knowledge Store directories for file changes and
'   queues supported documents for incremental background indexing.
'
' Responsibilities:
'   - Create and manage `FileSystemWatcher` instances for active stores.
'   - Detect newly created or modified source documents.
'   - Maintain an in-memory pending queue of files awaiting indexing.
'   - Perform periodic scans to recover missed events and discover files not yet
'     captured by watcher notifications.
'   - Process a limited number of queued files per idle tick and return
'     per-file indexing results.
'   - Keep failures isolated so one bad file or watcher error does not stop the
'     overall background indexing pipeline.
'
' Notes:
'   - The queue is intentionally in-memory; missed items are recovered by later
'     scans.
'   - Actual document processing is delegated to
'     `KnowledgeStoreProcessingService`.
'   - Designed for use from host idle services such as
'     `KnowledgeStoreIdleService`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Collections.Concurrent
Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Monitors Knowledge Store directories and queues new/modified documents for indexing.
    ''' Thread-safe — designed to be called from host idle timers.
    ''' </summary>
    Public Class KnowledgeStoreWatcher
        Implements IDisposable

#Region "Fields"

        Private ReadOnly _context As ISharedContext
        Private ReadOnly _watchers As New List(Of FileSystemWatcher)()
        Private ReadOnly _pendingFiles As New ConcurrentQueue(Of PendingFile)()
        Private ReadOnly _processLock As New SemaphoreSlim(1, 1)
        Private _isWatching As Boolean = False
        Private _disposed As Boolean = False

        ''' <summary>Maximum files to index per ProcessPendingAsync call.</summary>
        Public Const MaxFilesPerTick As Integer = 3

        ''' <summary>Supported file extensions for monitoring (lowercase, with dot).</summary>
        Private Shared ReadOnly SupportedExtensions As New HashSet(Of String)(
            KnowledgeIndexer.SupportedExtensions, StringComparer.OrdinalIgnoreCase)

#End Region

#Region "Data Model"

        ''' <summary>
        ''' Represents a file detected as new or modified, pending indexing.
        ''' </summary>
        Public Class PendingFile
            Public Property FilePath As String = ""
            Public Property StoreName As String = ""
            Public Property DetectedUtc As DateTime = DateTime.UtcNow
        End Class

        ''' <summary>
        ''' Result of a single indexing operation.
        ''' </summary>
        Public Class IndexResult
            Public Property FilePath As String = ""
            Public Property Success As Boolean = False
            Public Property Title As String = ""
            Public Property ErrorMessage As String = ""
        End Class

#End Region

#Region "Constructor"

        Public Sub New(context As ISharedContext)
            _context = context
        End Sub

#End Region

#Region "Start / Stop Watching"

        ''' <summary>
        ''' Starts watching all active, writable Knowledge Store directories.
        ''' Safe to call multiple times; subsequent calls are no-ops.
        ''' </summary>
        Public Sub StartWatching()
            If _isWatching Then Return

            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then Continue For
                If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then Continue For
                If Not Directory.Exists(store.ResolvedSourcePath) Then Continue For

                Try
                    Dim watcher As New FileSystemWatcher() With {
                        .Path = store.ResolvedSourcePath,
                        .IncludeSubdirectories = store.ScanSubdirectories,
                        .NotifyFilter = NotifyFilters.FileName Or NotifyFilters.LastWrite Or NotifyFilters.CreationTime,
                        .InternalBufferSize = 16384,
                        .EnableRaisingEvents = True
                    }

                    Dim storeName = store.Name

                    AddHandler watcher.Created, Sub(s, e) OnFileDetected(e.FullPath, storeName)
                    AddHandler watcher.Changed, Sub(s, e) OnFileDetected(e.FullPath, storeName)
                    AddHandler watcher.Renamed, Sub(s, e) OnFileDetected(e.FullPath, storeName)
                    AddHandler watcher.Error, Sub(s, e)
                                                  Debug.WriteLine($"KSWatcher error on '{storeName}': {e.GetException().Message}")
                                              End Sub

                    _watchers.Add(watcher)
                Catch ex As Exception
                    Debug.WriteLine($"KSWatcher: Failed to watch '{store.ResolvedSourcePath}': {ex.Message}")
                End Try
            Next

            _isWatching = True
        End Sub

        ''' <summary>
        ''' Stops all file watchers. The pending queue is intentionally NOT drained —
        ''' on restart, the periodic scan will re-detect any files that were pending.
        ''' </summary>
        Public Sub StopWatching()
            For Each w In _watchers
                Try
                    w.EnableRaisingEvents = False
                    w.Dispose()
                Catch
                End Try
            Next
            _watchers.Clear()
            _isWatching = False
        End Sub

#End Region

#Region "File Detection"

        Private Sub OnFileDetected(filePath As String, storeName As String)
            Try
                If IsMetadataPath(filePath) Then Return

                Dim ext = Path.GetExtension(filePath)
                If String.IsNullOrWhiteSpace(ext) Then Return
                If Not SupportedExtensions.Contains(ext.ToLowerInvariant()) Then Return

                ' Debounce: skip if already in queue
                Dim normalizedPath = filePath.ToUpperInvariant()
                If _pendingFiles.Any(Function(p) p.FilePath.ToUpperInvariant() = normalizedPath) Then Return

                _pendingFiles.Enqueue(New PendingFile With {
                    .FilePath = filePath,
                    .StoreName = storeName,
                    .DetectedUtc = DateTime.UtcNow
                })
            Catch
                ' Watcher callbacks must never throw
            End Try
        End Sub

        ''' <summary>
        ''' Returns True if the path is inside a .redink/ metadata folder.
        ''' </summary>
        Private Shared Function IsMetadataPath(filePath As String) As Boolean
            Dim sep1 = $"{Path.DirectorySeparatorChar}.redink{Path.DirectorySeparatorChar}"
            Dim sep2 = $"{Path.AltDirectorySeparatorChar}.redink{Path.AltDirectorySeparatorChar}"
            Return filePath.IndexOf(sep1, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   filePath.IndexOf(sep2, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

#End Region

#Region "Periodic Full Scan"

        ''' <summary>
        ''' Performs a full directory scan of all active stores and enqueues files
        ''' that are not yet in the manifest (or have been modified since last index).
        ''' </summary>
        Public Sub RunPeriodicScan()
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then Continue For
                If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then Continue For
                If Not Directory.Exists(store.ResolvedSourcePath) Then Continue For

                Try
                    Dim manifest = KnowledgeStoreManifest.Load(store)
                    Dim dirSearchOption As SearchOption = If(store.ScanSubdirectories,
                                          SearchOption.AllDirectories,
                                          SearchOption.TopDirectoryOnly)

                    For Each ext In KnowledgeIndexer.SupportedExtensions
                        Dim files As String()
                        Try
                            files = Directory.GetFiles(store.ResolvedSourcePath, "*" & ext, dirSearchOption)
                        Catch
                            Continue For
                        End Try

                        For Each filePath In files
                            If IsMetadataPath(filePath) Then Continue For

                            Dim entry = manifest.FindByPath(filePath)
                            If entry Is Nothing Then
                                Dim legacyRawCandidate = Path.Combine(store.ResolvedSourcePath, ".redink", "Raw", Path.GetFileName(filePath))
                                entry = manifest.FindByPath(legacyRawCandidate)
                            End If

                            If entry Is Nothing Then
                                OnFileDetected(filePath, store.Name)
                            Else
                                Try
                                    Dim lastWrite = File.GetLastWriteTimeUtc(filePath)
                                    If lastWrite > entry.IndexedDate.ToUniversalTime() Then
                                        OnFileDetected(filePath, store.Name)
                                    End If
                                Catch
                                End Try
                            End If
                        Next
                    Next
                Catch ex As Exception
                    Debug.WriteLine($"KSWatcher: Scan error for '{store.Name}': {ex.Message}")
                End Try
            Next
        End Sub

#End Region

#Region "Process Pending Queue"

        ''' <summary>
        ''' Dequeues up to maxFiles pending files and indexes them.
        ''' Thread-safe via SemaphoreSlim.
        ''' </summary>
        Public Async Function ProcessPendingAsync(Optional maxFiles As Integer = MaxFilesPerTick) As Task(Of List(Of IndexResult))
            Dim results As New List(Of IndexResult)()
            If Not Await _processLock.WaitAsync(0).ConfigureAwait(False) Then
                Return results
            End If

            Dim trayIcon As System.Windows.Forms.NotifyIcon = Nothing

            Try
                Dim processed As Integer = 0
                Dim startCount As Integer = _pendingFiles.Count

                If startCount > 0 AndAlso _context.INI_KnowledgeStoreUseLLMIndex Then
                    Try
                        trayIcon = New System.Windows.Forms.NotifyIcon()
                        trayIcon.Icon = System.Drawing.SystemIcons.Information
                        trayIcon.Visible = True
                        trayIcon.Text = $"Knowledge Store: Preparing to index {startCount} file(s)..."
                    Catch
                    End Try
                End If

                While processed < maxFiles AndAlso processed < startCount
                    Dim pending As PendingFile = Nothing
                    If Not _pendingFiles.TryDequeue(pending) Then Exit While

                    Dim safeFileName = Path.GetFileName(pending.FilePath)

                    If trayIcon IsNot Nothing Then
                        Dim statusText = $"Knowledge Store ({processed + 1}/{startCount}): {safeFileName}"
                        trayIcon.Text = If(statusText.Length > 63, statusText.Substring(0, 60) & "...", statusText)
                    End If

                    Dim result As New IndexResult() With {.FilePath = pending.FilePath}

                    Try
                        If Not File.Exists(pending.FilePath) Then
                            result.ErrorMessage = "File no longer exists."
                            results.Add(result)
                            processed += 1
                            Continue While
                        End If

                        Dim store = KnowledgeStoreCatalog.GetStoreByName(pending.StoreName, _context)
                        If store Is Nothing OrElse Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                            result.ErrorMessage = "Store not found or not writable."
                            results.Add(result)
                            processed += 1
                            Continue While
                        End If

                        Dim processResult = Await KnowledgeStoreProcessingService.ProcessDocumentAsync(
                            store:=store,
                            filePath:=pending.FilePath,
                            context:=_context,
                            useLlmIndex:=_context.INI_KnowledgeStoreUseLLMIndex,
                            isBackground:=True).ConfigureAwait(False)

                        If Not processResult.Success OrElse processResult.Entry Is Nothing Then
                            result.ErrorMessage = processResult.ErrorMessage
                            results.Add(result)
                            processed += 1
                            Continue While
                        End If

                        result.Success = True
                        result.Title = processResult.Entry.Title

                    Catch ex As Exception
                        result.ErrorMessage = ex.Message
                        KnowledgeWikiService.LogWikiError(
                            KnowledgeStoreCatalog.GetStoreByName(pending.StoreName, _context)?.ResolvedSourcePath,
                            pending.FilePath,
                            ex.Message)
                    End Try

                    results.Add(result)
                    processed += 1
                End While

            Finally
                If trayIcon IsNot Nothing Then
                    trayIcon.Visible = False
                    trayIcon.Dispose()
                End If
                _processLock.Release()
            End Try

            Return results
        End Function

#End Region

#Region "Wiki Page Maintenance"

        ' Future: Replace with LLM-generated rich summaries including entity cross-references,
        ' contradiction flags, and related-document links.

        Friend Shared Sub WriteWikiSummaryPage(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                 entry As KnowledgeStoreManager.KnowledgeEntry)
            Dim wikiDir = KnowledgeStoreCatalog.GetWikiPath(store)
            If String.IsNullOrWhiteSpace(wikiDir) Then Return

            Dim safeName = SanitizeFileName(If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "unknown"))))
            Dim wikiFile = Path.Combine(wikiDir, safeName & "-summary.md")
            Dim tmpFile = wikiFile & ".tmp"

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"# {entry.Title}")
            sb.AppendLine()
            sb.AppendLine($"**Source:** `{entry.FilePath}`")
            sb.AppendLine($"**Indexed:** {entry.IndexedDate:yyyy-MM-dd HH:mm}")
            If entry.Tags IsNot Nothing AndAlso entry.Tags.Length > 0 Then
                sb.AppendLine($"**Tags:** {String.Join(", ", entry.Tags)}")
            End If
            If entry.Keywords IsNot Nothing AndAlso entry.Keywords.Length > 0 Then
                sb.AppendLine($"**Keywords:** {String.Join(", ", entry.Keywords.Take(15))}")
            End If
            sb.AppendLine()
            sb.AppendLine("## Summary")
            sb.AppendLine()
            sb.AppendLine(If(entry.Summary, "(no summary available)"))
            sb.AppendLine()
            ' Future: Entity cross-references, related documents, contradiction notes
            sb.AppendLine("<!-- TODO: Entity cross-references, related documents, contradiction notes -->")

            ' Atomic write for wiki page too
            File.WriteAllText(tmpFile, sb.ToString(), System.Text.Encoding.UTF8)
            If File.Exists(wikiFile) Then File.Delete(wikiFile)
            File.Move(tmpFile, wikiFile)
        End Sub

        Friend Shared Sub UpdateIndexFile(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                           entry As KnowledgeStoreManager.KnowledgeEntry)
            Dim metaDir = KnowledgeStoreCatalog.GetMetadataPath(store)
            If String.IsNullOrWhiteSpace(metaDir) Then Return

            Dim indexPath = Path.Combine(metaDir, KnowledgeStoreCatalog.IndexFile)
            Dim safeName = SanitizeFileName(If(entry.Title, "unknown"))
            Dim line = $"- [{entry.Title}](wiki/{safeName}-summary.md) — {TruncateSummary(entry.Summary, 120)} ({entry.IndexedDate:yyyy-MM-dd})"

            Dim lines As New List(Of String)()
            If File.Exists(indexPath) Then
                Try
                    lines.AddRange(File.ReadAllLines(indexPath, System.Text.Encoding.UTF8))
                Catch
                End Try
            End If

            If lines.Count = 0 Then
                lines.Add($"# {store.Name} — Knowledge Store Index")
                lines.Add("")
                lines.Add("Auto-maintained catalog of all indexed documents.")
                lines.Add("")
                ' Future: Organize by category/entity when entity extraction is implemented
                lines.Add("<!-- TODO: Organize by category/entity when entity extraction is implemented -->")
                lines.Add("")
            End If

            Dim prefix = $"- [{entry.Title}]("
            Dim existingIdx = lines.FindIndex(Function(l) l.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            If existingIdx >= 0 Then
                lines(existingIdx) = line
            Else
                lines.Add(line)
            End If

            ' Atomic write
            Dim tmpPath = indexPath & ".tmp"
            File.WriteAllLines(tmpPath, lines, System.Text.Encoding.UTF8)
            If File.Exists(indexPath) Then File.Delete(indexPath)
            File.Move(tmpPath, indexPath)
        End Sub

        Friend Shared Sub AppendLogEntry(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                          entry As KnowledgeStoreManager.KnowledgeEntry)
            Dim metaDir = KnowledgeStoreCatalog.GetMetadataPath(store)
            If String.IsNullOrWhiteSpace(metaDir) Then Return

            Dim logPath = Path.Combine(metaDir, KnowledgeStoreCatalog.LogFile)
            Dim logLine = $"## [{DateTime.UtcNow:yyyy-MM-dd HH:mm}] ingest | {entry.Title}" & vbCrLf &
                          $"File: `{entry.FilePath}`" & vbCrLf &
                          $"Keywords: {If(entry.Keywords IsNot Nothing, String.Join(", ", entry.Keywords.Take(10)), "(none)")}" & vbCrLf

            ' AppendAllText is already atomic at the OS level for reasonable sizes
            File.AppendAllText(logPath, logLine & vbCrLf, System.Text.Encoding.UTF8)
        End Sub

#End Region

#Region "Helpers"

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(name.Where(Function(c) Not invalid.Contains(c)).ToArray())
            If cleaned.Length > 100 Then cleaned = cleaned.Substring(0, 100)
            Return If(String.IsNullOrWhiteSpace(cleaned), "unnamed", cleaned.Trim())
        End Function

        Private Shared Function TruncateSummary(summary As String, maxLen As Integer) As String
            If String.IsNullOrWhiteSpace(summary) Then Return "(no summary)"
            If summary.Length <= maxLen Then Return summary
            Dim cut = summary.LastIndexOf(" "c, maxLen)
            If cut < maxLen \ 2 Then cut = maxLen
            Return summary.Substring(0, cut).Trim() & "..."
        End Function

        ''' <summary>Returns the number of files currently queued for indexing.</summary>
        Public ReadOnly Property PendingCount As Integer
            Get
                Return _pendingFiles.Count
            End Get
        End Property

#End Region

#Region "IDisposable"

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            StopWatching()
            _processLock.Dispose()
        End Sub

#End Region

    End Class

End Namespace