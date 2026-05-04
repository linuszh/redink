' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreManifest.vb
' Purpose:
'   Manages the per-store manifest file `.redink\manifest.json`, which tracks
'   indexed documents and their metadata within a single Knowledge Store.
'
' Responsibilities:
'   - Load manifest entries for a specific store.
'   - Tolerate missing, empty, locked, or malformed manifest files safely.
'   - Add, update, remove, and locate manifest entries by normalized file path.
'   - Persist manifest changes using safe write semantics.
'   - Preserve per-store indexing state independently from catalog metadata.
'
' Notes:
'   - The manifest is store-local and distinct from the global store catalog.
'   - Corrupted manifests are treated defensively and can be rebuilt on save.
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
        Private ReadOnly _storeRoot As String = ""

#End Region

#Region "Construction"

        Private Sub New(Optional storeRoot As String = "")
            _storeRoot = NormalizeStoreRoot(storeRoot)
        End Sub

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
            Dim manifest As New KnowledgeStoreManifest(If(store IsNot Nothing, store.ResolvedSourcePath, ""))
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
                    Debug.WriteLine($"KnowledgeStoreManifest.Load: IO error reading manifest: {ex.Message}")
                    Return manifest
                End Try

                If String.IsNullOrWhiteSpace(json) Then Return manifest

                Dim arr As JArray
                Try
                    arr = JArray.Parse(json)
                Catch ex As JsonReaderException
                    Debug.WriteLine($"KnowledgeStoreManifest.Load: Corrupted JSON in manifest, will rebuild on next save: {ex.Message}")
                    Return manifest
                End Try

                For Each token As JToken In arr
                    Try
                        If token.Type <> JTokenType.Object Then Continue For
                        Dim obj = CType(token, JObject)

                        Dim entry As New KnowledgeStoreManager.KnowledgeEntry() With {
                            .FilePath = NormalizeStoredPathForStore(
                                If(store.ResolvedSourcePath, ""),
                                If(obj.Value(Of String)("FilePath"), "")),
                            .Title = If(obj.Value(Of String)("Title"), ""),
                            .Summary = If(obj.Value(Of String)("Summary"), ""),
                            .IsFromCentralIndex = False
                        }

                        Dim dateStr = obj.Value(Of String)("IndexedDate")
                        If Not String.IsNullOrWhiteSpace(dateStr) Then
                            Dim parsed As DateTime
                            If DateTime.TryParse(
                                dateStr,
                                System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.RoundtripKind,
                                parsed) Then

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
                Dim persistedPath = NormalizeStoredPathForStore(store.ResolvedSourcePath, entry.FilePath)
                entry.FilePath = persistedPath

                Dim obj As New JObject()
                obj("FilePath") = persistedPath
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
                File.WriteAllText(tmpPath, arr.ToString(Formatting.Indented), System.Text.Encoding.UTF8)

                If File.Exists(manifestPath) Then
                    File.Delete(manifestPath)
                End If

                File.Move(tmpPath, manifestPath)
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreManifest.Save: Error during atomic write: {ex.Message}")
                Try
                    If File.Exists(tmpPath) Then File.Delete(tmpPath)
                Catch
                End Try
            End Try
        End Sub

#End Region

#Region "Query / Mutate"

        ''' <summary>
        ''' Finds an entry by file path using exact match first, then relocated-path fallback.
        ''' </summary>
        Public Function FindByPath(filePath As String) As KnowledgeStoreManager.KnowledgeEntry
            If String.IsNullOrWhiteSpace(filePath) Then Return Nothing

            Dim normalized = NormalizeComparisonPath(_storeRoot, KnowledgeStoreCatalog.StripQuotes(filePath))

            Dim exact = _entries.FirstOrDefault(
                Function(e) NormalizeComparisonPath(_storeRoot, KnowledgeStoreCatalog.StripQuotes(If(e.FilePath, ""))) = normalized)

            If exact IsNot Nothing Then
                Return exact
            End If

            Return FindByRelocatedPath(filePath)
        End Function

        ''' <summary>
        ''' Adds a new entry or updates an existing one (matched by FilePath).
        ''' </summary>
        Public Sub AddOrUpdate(entry As KnowledgeStoreManager.KnowledgeEntry)
            If entry Is Nothing Then Return

            entry.FilePath = NormalizeStoredPathForStore(_storeRoot, entry.FilePath)

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

        Private Function FindByRelocatedPath(filePath As String) As KnowledgeStoreManager.KnowledgeEntry
            Dim targetAbsolute = ResolveToComparableAbsolutePath(_storeRoot, filePath)
            Dim targetFileName = Path.GetFileName(If(targetAbsolute, ""))

            If String.IsNullOrWhiteSpace(targetFileName) Then
                Return Nothing
            End If

            Dim candidates = _entries.
                Where(Function(e) Path.GetFileName(If(e.FilePath, "")).Equals(targetFileName, StringComparison.OrdinalIgnoreCase)).
                Select(
                    Function(e)
                        Dim candidateAbsolute = ResolveToComparableAbsolutePath(_storeRoot, e.FilePath)
                        Dim score = GetTrailingSegmentMatchScore(targetAbsolute, candidateAbsolute)
                        Return Tuple.Create(e, score)
                    End Function).
                Where(Function(x) x.Item2 > 0).
                OrderByDescending(Function(x) x.Item2).
                ToList()

            If candidates.Count = 0 Then
                Return Nothing
            End If

            If candidates.Count = 1 Then
                Return candidates(0).Item1
            End If

            If candidates(0).Item2 > candidates(1).Item2 Then
                Return candidates(0).Item1
            End If

            Return Nothing
        End Function

        Friend Shared Function NormalizeStoredPathForStore(storeRoot As String, rawPath As String) As String
            Dim cleaned = KnowledgeStoreCatalog.StripQuotes(rawPath)
            If String.IsNullOrWhiteSpace(cleaned) Then Return ""

            Dim expanded As String
            Try
                expanded = ExpandEnvironmentVariables(cleaned.Trim())
            Catch
                expanded = cleaned.Trim()
            End Try

            If Path.IsPathRooted(expanded) Then
                Dim relative = TryGetRelativeToStoreRoot(storeRoot, expanded)
                If Not String.IsNullOrWhiteSpace(relative) Then
                    Return relative
                End If

                Try
                    Return Path.GetFullPath(expanded)
                Catch
                    Return expanded
                End Try
            End If

            Dim relativePath = expanded.Trim().TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If relativePath.Length = 0 Then Return ""

            Return relativePath.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
        End Function

        Private Shared Function NormalizeComparisonPath(storeRoot As String, rawPath As String) As String
            Dim normalized = NormalizeStoredPathForStore(storeRoot, rawPath)
            If String.IsNullOrWhiteSpace(normalized) Then Return ""

            If Path.IsPathRooted(normalized) Then
                Try
                    Return Path.GetFullPath(normalized).ToUpperInvariant()
                Catch
                    Return normalized.Trim().ToUpperInvariant()
                End Try
            End If

            Return normalized.
                Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).
                Trim().
                ToUpperInvariant()
        End Function

        Private Shared Function ResolveToComparableAbsolutePath(storeRoot As String, rawPath As String) As String
            Dim normalized = NormalizeStoredPathForStore(storeRoot, rawPath)
            If String.IsNullOrWhiteSpace(normalized) Then Return ""

            Try
                If Path.IsPathRooted(normalized) Then
                    Return Path.GetFullPath(normalized)
                End If

                If String.IsNullOrWhiteSpace(storeRoot) Then
                    Return normalized
                End If

                Return Path.GetFullPath(Path.Combine(storeRoot, normalized))
            Catch
                Return normalized
            End Try
        End Function

        Private Shared Function GetTrailingSegmentMatchScore(pathA As String, pathB As String) As Integer
            If String.IsNullOrWhiteSpace(pathA) OrElse String.IsNullOrWhiteSpace(pathB) Then
                Return 0
            End If

            Dim partsA = pathA.
                Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).
                Split(New Char() {Path.DirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries)

            Dim partsB = pathB.
                Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar).
                Split(New Char() {Path.DirectorySeparatorChar}, StringSplitOptions.RemoveEmptyEntries)

            Dim idxA = partsA.Length - 1
            Dim idxB = partsB.Length - 1
            Dim score As Integer = 0

            While idxA >= 0 AndAlso idxB >= 0
                If Not partsA(idxA).Equals(partsB(idxB), StringComparison.OrdinalIgnoreCase) Then
                    Exit While
                End If

                score += 1
                idxA -= 1
                idxB -= 1
            End While

            Return score
        End Function

        Private Shared Function NormalizeStoreRoot(storeRoot As String) As String
            If String.IsNullOrWhiteSpace(storeRoot) Then Return ""

            Try
                Return Path.GetFullPath(ExpandEnvironmentVariables(storeRoot.Trim())).
                    TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Catch
                Return storeRoot.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            End Try
        End Function

        Private Shared Function TryGetRelativeToStoreRoot(storeRoot As String, candidatePath As String) As String
            Dim normalizedStoreRoot = NormalizeStoreRoot(storeRoot)
            If String.IsNullOrWhiteSpace(normalizedStoreRoot) Then Return ""
            If String.IsNullOrWhiteSpace(candidatePath) Then Return ""

            Try
                Dim fullCandidate = Path.GetFullPath(candidatePath.Trim())
                Dim rootWithSeparator = normalizedStoreRoot

                If Not rootWithSeparator.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                    rootWithSeparator &= Path.DirectorySeparatorChar
                End If

                If Not fullCandidate.StartsWith(rootWithSeparator, StringComparison.OrdinalIgnoreCase) Then
                    Return ""
                End If

                Return GetRelativePathCompat(normalizedStoreRoot, fullCandidate).
                    TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetRelativePathCompat(basePath As String, targetPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(basePath) Then Return targetPath
                If String.IsNullOrWhiteSpace(targetPath) Then Return ""

                Dim normalizedBase = Path.GetFullPath(basePath)
                Dim normalizedTarget = Path.GetFullPath(targetPath)

                If Not normalizedBase.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                    normalizedBase &= Path.DirectorySeparatorChar
                End If

                Dim baseUri As New Uri(normalizedBase, UriKind.Absolute)
                Dim targetUri As New Uri(normalizedTarget, UriKind.Absolute)

                Dim relativeUri = baseUri.MakeRelativeUri(targetUri)
                Dim relativePath = Uri.UnescapeDataString(relativeUri.ToString())

                Return relativePath.Replace("/"c, Path.DirectorySeparatorChar)
            Catch
                Return Path.GetFileName(targetPath)
            End Try
        End Function

#End Region

    End Class

End Namespace