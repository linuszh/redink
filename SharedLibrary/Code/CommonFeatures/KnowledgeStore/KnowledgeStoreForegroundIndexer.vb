' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreForegroundIndexer.vb
' Purpose:
'   Runs a user-initiated foreground indexing pass for one or more Knowledge
'   Stores with progress reporting and cancellation support.
'
' Responsibilities:
'   - Collect target stores and enumerate supported source documents.
'   - Skip metadata files and, unless forced, already up-to-date files.
'   - Drive the shared document-processing pipeline for each queued file.
'   - Update progress UI through `ProgressBarModule`.
'   - Return aggregate indexing statistics including indexed, skipped, failed,
'     and cancelled counts.
'
' Notes:
'   - This is the user-visible counterpart to watcher-driven background indexing.
'   - Per-file processing is delegated to `KnowledgeStoreProcessingService`.
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
                Dim exactStore = KnowledgeStoreCatalog.GetStoreById(storeName, context)

                If exactStore IsNot Nothing Then
                    stores = New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) From {exactStore}
                Else
                    stores = KnowledgeStoreCatalog.GetStoresByName(storeName, context).
            Where(Function(s) s.Active).ToList()
                End If
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
                        Dim processResult = Await KnowledgeStoreProcessingService.ProcessDocumentAsync(
                        store:=item.Store,
                        filePath:=item.FilePath,
                        context:=context,
                        useLlmIndex:=context.INI_KnowledgeStoreUseLLMIndex).ConfigureAwait(False)

                        If Not processResult.Success OrElse processResult.Entry Is Nothing Then
                            result.FailedFiles += 1
                            Continue For
                        End If

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