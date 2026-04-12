' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreManifest.vb
' Purpose:
'   Manages the per-store manifest file `.redink\manifest.json`, which tracks
'   indexed documents and their metadata inside a single Knowledge Store.
'
' Responsibilities:
'   - Load manifest entries for a specific store.
'   - Tolerate missing, empty, locked, or malformed manifest files safely.
'   - Add, update, remove, and locate manifest entries by normalized file path.
'   - Persist manifest changes with atomic-write semantics.
'   - Preserve per-store indexing state independently from catalog metadata.
'
' Notes:
'   - This manifest is store-local and distinct from broader catalog/index
'     configuration files.
'   - Corrupted manifests are treated as empty and can be rebuilt on save.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Manages the per-store manifest file listing all indexed documents and their metadata.
    ''' </summary>
    Public Class KnowledgeStoreManifest

#Region "Fields"

        Private ReadOnly _entries As New List(Of KnowledgeStoreManager.KnowledgeEntry)()

#End Region

#Region "Properties"

        ''' <summary>All entries in this manifest.</summary>
        Public ReadOnly Property Entries As List(Of KnowledgeStoreManager.KnowledgeEntry)
            Get
                Return _entries
            End Get
        End Property

#End Region

#Region "Load"

        ''' <summary>
        ''' Loads the manifest for a given store definition.
        ''' Returns an empty manifest if the file doesn't exist or is corrupted.
        ''' </summary>
        Public Shared Function Load(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As KnowledgeStoreManifest
            Dim manifest As New KnowledgeStoreManifest()
            If store Is Nothing Then Return manifest

            Try
                Dim metaDir = KnowledgeStoreCatalog.GetMetadataPath(store, createIfMissing:=False)
                If String.IsNullOrWhiteSpace(metaDir) Then Return manifest

                Dim manifestPath = Path.Combine(metaDir, KnowledgeStoreCatalog.ManifestFile)
                If Not File.Exists(manifestPath) Then Return manifest

                Dim json As String
                Try
                    json = File.ReadAllText(manifestPath, System.Text.Encoding.UTF8)
                Catch ex As IOException
                    ' File locked or inaccessible — return empty rather than crashing
                    Debug.WriteLine($"KnowledgeStoreManifest.Load: IO error reading manifest: {ex.Message}")
                    Return manifest
                End Try

                If String.IsNullOrWhiteSpace(json) Then Return manifest

                Dim arr As JArray
                Try
                    arr = JArray.Parse(json)
                Catch ex As JsonReaderException
                    ' Corrupted JSON — return empty manifest; next Save() will overwrite cleanly
                    Debug.WriteLine($"KnowledgeStoreManifest.Load: Corrupted JSON in manifest, will rebuild on next save: {ex.Message}")
                    Return manifest
                End Try

                For Each token As JToken In arr
                    Try
                        If token.Type <> JTokenType.Object Then Continue For
                        Dim obj = CType(token, JObject)

                        Dim entry As New KnowledgeStoreManager.KnowledgeEntry() With {
                            .FilePath = If(obj.Value(Of String)("FilePath"), ""),
                            .Title = If(obj.Value(Of String)("Title"), ""),
                            .Summary = If(obj.Value(Of String)("Summary"), ""),
                            .IsFromCentralIndex = False
                        }

                        Dim dateStr = obj.Value(Of String)("IndexedDate")
                        If Not String.IsNullOrWhiteSpace(dateStr) Then
                            Dim parsed As DateTime
                            If DateTime.TryParse(dateStr, System.Globalization.CultureInfo.InvariantCulture,
                                                 System.Globalization.DateTimeStyles.RoundtripKind, parsed) Then
                                entry.IndexedDate = parsed
                            End If
                        End If

                        Dim kwToken = obj("Keywords")
                        If kwToken IsNot Nothing AndAlso kwToken.Type = JTokenType.Array Then
                            entry.Keywords = kwToken.ToObject(Of String())()
                        End If

                        Dim tagToken = obj("Tags")
                        If tagToken IsNot Nothing AndAlso tagToken.Type = JTokenType.Array Then
                            entry.Tags = tagToken.ToObject(Of String())()
                        End If

                        manifest._entries.Add(entry)
                    Catch
                        ' Skip malformed entries — do not fail the whole manifest
                    End Try
                Next
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreManifest.Load: Unexpected error: {ex.Message}")
            End Try

            Return manifest
        End Function

#End Region

#Region "Save (Atomic)"

        ''' <summary>
        ''' Saves this manifest to the store's .redink/manifest.json using atomic write.
        ''' Writes to a .tmp file first, then replaces the original to prevent corruption.
        ''' </summary>
        Public Sub Save(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition)
            If store Is Nothing Then Return

            Dim metaDir = KnowledgeStoreCatalog.GetMetadataPath(store, createIfMissing:=True)
            If String.IsNullOrWhiteSpace(metaDir) Then Return

            Dim manifestPath = Path.Combine(metaDir, KnowledgeStoreCatalog.ManifestFile)
            Dim tmpPath = manifestPath & ".tmp"

            Dim arr As New JArray()
            For Each entry In _entries
                Dim obj As New JObject()
                obj("FilePath") = If(entry.FilePath, "")
                obj("Title") = If(entry.Title, "")
                obj("Summary") = If(entry.Summary, "")
                obj("IndexedDate") = entry.IndexedDate.ToString("o")
                If entry.Keywords IsNot Nothing AndAlso entry.Keywords.Length > 0 Then
                    obj("Keywords") = New JArray(entry.Keywords)
                End If
                If entry.Tags IsNot Nothing AndAlso entry.Tags.Length > 0 Then
                    obj("Tags") = New JArray(entry.Tags)
                End If
                arr.Add(obj)
            Next

            Try
                ' Write to temp file first
                File.WriteAllText(tmpPath, arr.ToString(Formatting.Indented), System.Text.Encoding.UTF8)

                ' Atomic replace: rename .tmp over the original
                ' File.Move with overwrite requires .NET Framework 4.0+ / .NET Core 3.0+
                ' For safety, delete then move if overwrite not available
                If File.Exists(manifestPath) Then
                    File.Delete(manifestPath)
                End If
                File.Move(tmpPath, manifestPath)
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreManifest.Save: Error during atomic write: {ex.Message}")
                ' Clean up temp file if it exists
                Try
                    If File.Exists(tmpPath) Then File.Delete(tmpPath)
                Catch
                End Try
            End Try
        End Sub

#End Region

#Region "Query / Mutate"

        ''' <summary>
        ''' Finds an entry by file path (case-insensitive, environment-variable-aware).
        ''' </summary>
        Public Function FindByPath(filePath As String) As KnowledgeStoreManager.KnowledgeEntry
            If String.IsNullOrWhiteSpace(filePath) Then Return Nothing
            Dim normalized = NormalizePath(KnowledgeStoreCatalog.StripQuotes(filePath))
            Return _entries.FirstOrDefault(
                Function(e) NormalizePath(KnowledgeStoreCatalog.StripQuotes(If(e.FilePath, ""))) = normalized)
        End Function

        ''' <summary>
        ''' Adds a new entry or updates an existing one (matched by FilePath).
        ''' </summary>
        Public Sub AddOrUpdate(entry As KnowledgeStoreManager.KnowledgeEntry)
            If entry Is Nothing Then Return
            Dim existing = FindByPath(entry.FilePath)
            If existing IsNot Nothing Then
                _entries.Remove(existing)
            End If
            _entries.Add(entry)
        End Sub

        ''' <summary>
        ''' Removes an entry by file path.
        ''' </summary>
        Public Function RemoveByPath(filePath As String) As Boolean
            Dim entry = FindByPath(filePath)
            If entry Is Nothing Then Return False
            Return _entries.Remove(entry)
        End Function

        Private Shared Function NormalizePath(rawPath As String) As String
            If String.IsNullOrWhiteSpace(rawPath) Then Return ""
            Try
                Return ExpandEnvironmentVariables(rawPath.Trim()).ToUpperInvariant()
            Catch
                Return rawPath.Trim().ToUpperInvariant()
            End Try
        End Function

#End Region

    End Class

End Namespace