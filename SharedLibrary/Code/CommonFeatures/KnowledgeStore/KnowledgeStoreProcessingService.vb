' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreProcessingService.vb
' Purpose:
'   Provides the shared document-processing pipeline used by both foreground and
'   background Knowledge Store indexing so both modes behave identically.
'
' Responsibilities:
'   - Validate store, source file, and write-permission preconditions.
'   - Invoke `KnowledgeIndexer.IndexDocumentAsync` to build a
'     `KnowledgeStoreManager.KnowledgeEntry`.
'   - Load the target store manifest, upsert the processed entry, and save it.
'   - Return a structured `ProcessResult` with success state, entry payload, and
'     error details.
'   - Log processing failures into the Knowledge Wiki error log.
'
' Notes:
'   - This file centralizes the indexing write path shared by
'     `KnowledgeStoreForegroundIndexer` and watcher/idle-based background flows.
'   - It does not enumerate files itself; callers supply the target document.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Shared processing pipeline used by both foreground and background indexing.
    ''' This ensures both modes behave identically.
    ''' </summary>
    Public Class KnowledgeStoreProcessingService

        Public Class ProcessResult
            Public Property Success As Boolean = False
            Public Property Entry As KnowledgeStoreManager.KnowledgeEntry = Nothing
            Public Property ErrorMessage As String = ""
        End Class

        Public Shared Async Function ProcessDocumentAsync(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                          filePath As String,
                                                          context As ISharedContext,
                                                          Optional useLlmIndex As Boolean = True,
                                                          Optional isBackground As Boolean = False) As Task(Of ProcessResult)

            Dim result As New ProcessResult()

            If store Is Nothing Then
                result.ErrorMessage = "Store is missing."
                Return result
            End If

            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                result.ErrorMessage = "Source file does not exist."
                Return result
            End If

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, context) Then
                result.ErrorMessage = "Store is not writable for current user."
                Return result
            End If

            Try
                Dim entry = Await KnowledgeIndexer.IndexDocumentAsync(
                    filePath:=filePath,
                    kbRootPath:=store.ResolvedSourcePath,
                    context:=context,
                    useLLMIndex:=useLlmIndex,
                    isBackground:=isBackground).ConfigureAwait(False)

                If entry Is Nothing Then
                    result.ErrorMessage = "Indexer returned no entry."
                    Return result
                End If

                Dim manifest = KnowledgeStoreManifest.Load(store)
                manifest.AddOrUpdate(entry)
                manifest.Save(store)

                result.Success = True
                result.Entry = entry
                Return result
            Catch ex As Exception
                result.ErrorMessage = ex.Message
                KnowledgeWikiService.LogWikiError(store.ResolvedSourcePath, filePath, ex.Message)
                Return result
            End Try
        End Function

    End Class

End Namespace