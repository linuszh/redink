' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreForegroundIndexer.vb
' Purpose: Runs a foreground (user-initiated) full index of all or a specific
'          Knowledge Store with a progress dialog. Uses the shared
'          ProgressBarModule + DPIProgressForm pattern.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Runs a user-initiated foreground index of Knowledge Store(s) with progress UI.
    ''' </summary>
    Public Class KnowledgeStoreForegroundIndexer

#Region "Result Model"

        ''' <summary>
        ''' Summary of a foreground indexing run.
        ''' </summary>
        Public Class IndexRunResult
            Public Property TotalFiles As Integer = 0
            Public Property IndexedFiles As Integer = 0
            Public Property SkippedFiles As Integer = 0
            Public Property FailedFiles As Integer = 0
            Public Property WasCancelled As Boolean = False
        End Class

#End Region

#Region "Public Entry Point"

        ''' <summary>
        ''' Runs the foreground indexer asynchronously with a progress dialog.
        ''' </summary>
        ''' <param name="context">Shared context for configuration.</param>
        ''' <param name="storeName">Restrict to a single store by name. Empty = all active stores.</param>
        ''' <param name="forceReindex">If True, re-index files even if already in the manifest.</param>
        Public Shared Async Function RunAsync(
                context As ISharedContext,
                Optional storeName As String = "",
                Optional forceReindex As Boolean = False) As Task(Of IndexRunResult)

            Dim result As New IndexRunResult()

            ' ── Collect target stores ──
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(context)
            If Not String.IsNullOrWhiteSpace(storeName) Then
                stores = stores.Where(
                    Function(s) s.Name.Equals(storeName, StringComparison.OrdinalIgnoreCase)).ToList()
            End If

            If stores.Count = 0 Then Return result

            ' ── Collect all files to index ──
            Dim fileQueue As New List(Of (FilePath As String, Store As KnowledgeStoreCatalog.KnowledgeStoreDefinition))()

            For Each store In stores
                If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, context) Then Continue For
                If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then Continue For
                If Not Directory.Exists(store.ResolvedSourcePath) Then Continue For

                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim searchOpt = If(store.ScanSubdirectories,
                                   SearchOption.AllDirectories,
                                   SearchOption.TopDirectoryOnly)

                For Each ext In KnowledgeIndexer.SupportedExtensions
                    Dim files As String()
                    Try
                        files = Directory.GetFiles(store.ResolvedSourcePath, "*" & ext, searchOpt)
                    Catch
                        Continue For
                    End Try

                    For Each filePath In files
                        If IsMetadataPath(filePath) Then Continue For

                        If Not forceReindex Then
                            Dim existing = manifest.FindByPath(filePath)
                            If existing IsNot Nothing Then
                                Try
                                    Dim lastWrite = File.GetLastWriteTimeUtc(filePath)
                                    If lastWrite <= existing.IndexedDate.ToUniversalTime() Then
                                        result.SkippedFiles += 1
                                        Continue For
                                    End If
                                Catch
                                End Try
                            End If
                        End If

                        fileQueue.Add((filePath, store))
                    Next
                Next
            Next

            result.TotalFiles = fileQueue.Count + result.SkippedFiles
            If fileQueue.Count = 0 Then Return result

            ' ── Show progress UI ──
            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressMax = fileQueue.Count
            ProgressBarModule.GlobalProgressLabel = "Preparing ..."
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.ShowProgressBarInSeparateThread(
                $"{AN} Knowledge Store — Indexing", "Indexing documents ...")

            ' ── Process files ──
            Dim processed As Integer = 0

            For Each item In fileQueue
                If ProgressBarModule.CancelOperation Then
                    result.WasCancelled = True
                    Exit For
                End If

                processed += 1
                Dim fileName = Path.GetFileName(item.FilePath)
                ProgressBarModule.GlobalProgressLabel = $"({processed}/{fileQueue.Count}) {fileName}"
                ProgressBarModule.GlobalProgressValue = processed

                Try
                    Try
                        Dim entry = Await KnowledgeIndexer.IndexDocumentAsync(
                        item.FilePath, context, context.INI_KnowledgeStoreUseLLMIndex).ConfigureAwait(False)

                        If entry Is Nothing Then
                            result.FailedFiles += 1
                            Continue For
                        End If

                        Dim manifest = KnowledgeStoreManifest.Load(item.Store)
                        manifest.AddOrUpdate(entry)
                        manifest.Save(item.Store)

                        ' Wiki pages, index.md, log.md — failures here do NOT affect the manifest
                        Try : KnowledgeStoreWatcher.WriteWikiSummaryPage(item.Store, entry) : Catch : End Try
                        Try : KnowledgeStoreWatcher.UpdateIndexFile(item.Store, entry) : Catch : End Try
                        Try : KnowledgeStoreWatcher.AppendLogEntry(item.Store, entry) : Catch : End Try

                        result.IndexedFiles += 1
                    Catch
                        result.FailedFiles += 1
                    End Try
                Catch
                    result.FailedFiles += 1
                End Try
            Next

            ' ── Close progress UI ──
            ProgressBarModule.CancelOperation = True

            Return result
        End Function

#End Region

#Region "Helpers"

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

    End Class

End Namespace