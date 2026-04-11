' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreCatalog.vb
' Purpose: Manages the catalog of multiple named Knowledge Stores.
'          Each store points to a watched directory tree and carries metadata
'          (name, owner restriction, role, active flag).
'
' Architecture / How it works:
'  - Two catalog files exist (mirroring the central/local INI pattern):
'      * Central catalog  = INI_KnowledgeStorePath  → read-only for non-owners
'      * Local catalog     = INI_KnowledgeStorePathLocal → user-writable
'  - Both are JSON arrays of KnowledgeStoreDefinition objects.
'  - LoadAll() merges both catalogs; local definitions override central ones
'    with the same Name (case-insensitive).
'  - Per-store metadata is kept in a .redink/ subfolder inside each store's
'    SourcePath: manifest.json, wiki/ pages, index.md, log.md.
'
' JSON Format (per catalog file):
'  [
'    {
'      "Name": "Contracts",
'      "SourcePath": "\\\\server\\legal\\contracts",
'      "Owner": "jdoe",
'      "Role": "shared",
'      "Active": true,
'      "ScanSubdirectories": true
'    }, ...
'  ]
'
' Future enhancements (planned for next iteration):
'  - Entity/topic cross-linking pages in the wiki layer.
'  - Contradiction detection across wiki pages.
'  - Full lint/health-check operations.
'  - LLM-driven query-time page generation and filing.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Manages the catalog of named Knowledge Store definitions, loaded from central and local JSON files.
    ''' </summary>
    Public Class KnowledgeStoreCatalog

#Region "Data Model"

        ''' <summary>
        ''' Defines a single Knowledge Store: a watched directory with metadata.
        ''' </summary>
        Public Class KnowledgeStoreDefinition
            ''' <summary>Human-readable unique name (e.g., "Contracts", "Research").</summary>
            Public Property Name As String = ""

            ''' <summary>Root directory path to watch for source documents (supports env variables).</summary>
            Public Property SourcePath As String = ""

            ''' <summary>
            ''' Owner username (Environment.UserName). Only this user may write to the store's
            ''' index if Role = "shared". Empty = no restriction.
            ''' </summary>
            Public Property Owner As String = ""

            ''' <summary>
            ''' Role: "personal" (single-user), "shared" (multi-user, owner-restricted writes),
            ''' "readonly" (no indexing, query only).
            ''' </summary>
            Public Property Role As String = "personal"

            ''' <summary>Whether this store is active (monitored and queryable).</summary>
            Public Property Active As Boolean = True

            ''' <summary>Whether to scan subdirectories of SourcePath.</summary>
            Public Property ScanSubdirectories As Boolean = True

            ' ── Runtime-only fields (not serialized) ──

            ''' <summary>True if this definition came from the central catalog file.</summary>
            <JsonIgnore>
            Public Property IsFromCentralCatalog As Boolean = False

            ''' <summary>Expanded, resolved SourcePath (computed at load time).</summary>
            <JsonIgnore>
            Public Property ResolvedSourcePath As String = ""
        End Class

#End Region

#Region "Constants"

        ''' <summary>Subfolder inside each store's SourcePath that holds index metadata and wiki pages.</summary>
        Public Const MetadataFolder As String = ".redink"

        ''' <summary>Manifest file inside .redink/ listing all indexed documents.</summary>
        Public Const ManifestFile As String = "manifest.json"

        ''' <summary>Wiki subfolder inside .redink/ containing LLM-generated summary pages.</summary>
        Public Const WikiFolder As String = "wiki"

        ''' <summary>Auto-maintained catalog of all wiki pages with one-line summaries.</summary>
        Public Const IndexFile As String = "index.md"

        ''' <summary>Append-only chronological record of ingest/query/lint operations.</summary>
        Public Const LogFile As String = "log.md"

#End Region

#Region "Load / Merge"

        ''' <summary>
        ''' Loads and merges Knowledge Store definitions from both catalog files.
        ''' Local definitions override central ones with the same Name (case-insensitive).
        ''' Auto-creates missing local catalog files with an empty array.
        ''' </summary>
        Public Shared Function LoadAll(context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            Dim merged As New List(Of KnowledgeStoreDefinition)()
            Dim seenNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Ensure catalog files exist (auto-create local; validate both)
            EnsureCatalogFile(context.INI_KnowledgeStorePathLocal, isLocal:=True)
            EnsureCatalogFile(context.INI_KnowledgeStorePath, isLocal:=False)

            ' Local catalog first (takes precedence)
            Dim localDefs = LoadCatalogFile(context.INI_KnowledgeStorePathLocal, isFromCentral:=False)
            For Each def In localDefs
                If String.IsNullOrWhiteSpace(def.Name) Then Continue For
                merged.Add(def)
                seenNames.Add(def.Name)
            Next

            ' Central catalog — skip duplicates
            Dim centralDefs = LoadCatalogFile(context.INI_KnowledgeStorePath, isFromCentral:=True)
            For Each def In centralDefs
                If String.IsNullOrWhiteSpace(def.Name) Then Continue For
                If seenNames.Contains(def.Name) Then Continue For
                merged.Add(def)
                seenNames.Add(def.Name)
            Next

            Return merged
        End Function

        ''' <summary>
        ''' Ensures a catalog file exists. If the path is configured but the file is missing,
        ''' auto-creates it with an empty JSON array (local only). If the file exists but
        ''' contains invalid JSON, logs an error and resets it (local only).
        ''' </summary>
        Private Shared Sub EnsureCatalogFile(rawPath As String, isLocal As Boolean)
            If String.IsNullOrWhiteSpace(rawPath) Then Return

            Try
                Dim expanded = ExpandEnvironmentVariables(rawPath.Trim())
                If String.IsNullOrWhiteSpace(expanded) Then Return

                ' Auto-create missing file (local catalog only — central is admin-managed)
                If Not File.Exists(expanded) Then
                    If isLocal Then
                        Dim dir = Path.GetDirectoryName(expanded)
                        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                            Directory.CreateDirectory(dir)
                        End If
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Created empty local catalog at '{expanded}'")
                    Else
                        Debug.WriteLine($"KnowledgeStoreCatalog: Central catalog not found at '{expanded}' — skipping.")
                    End If
                    Return
                End If

                ' Check for empty/whitespace file — but try to recover from .tmp first
                Dim json = File.ReadAllText(expanded, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then
                    ' The file exists but is empty — likely a crash during write.
                    ' Try to recover from .tmp file (atomic write leftovers)
                    Dim tmpPath = expanded & ".tmp"
                    If File.Exists(tmpPath) Then
                        Dim tmpJson = File.ReadAllText(tmpPath, System.Text.Encoding.UTF8)
                        If Not String.IsNullOrWhiteSpace(tmpJson) Then
                            Try
                                Dim tmpToken = JToken.Parse(tmpJson)
                                If tmpToken.Type = JTokenType.Array Then
                                    File.WriteAllText(expanded, tmpJson, System.Text.Encoding.UTF8)
                                    Debug.WriteLine($"KnowledgeStoreCatalog: Recovered catalog from '{tmpPath}'")
                                    Return
                                End If
                            Catch
                                ' .tmp is also bad — fall through
                            End Try
                        End If
                    End If

                    ' Also try .bak
                    Dim bakPath = expanded & ".bak"
                    If File.Exists(bakPath) Then
                        Dim bakJson = File.ReadAllText(bakPath, System.Text.Encoding.UTF8)
                        If Not String.IsNullOrWhiteSpace(bakJson) Then
                            Try
                                Dim bakToken = JToken.Parse(bakJson)
                                If bakToken.Type = JTokenType.Array Then
                                    File.WriteAllText(expanded, bakJson, System.Text.Encoding.UTF8)
                                    Debug.WriteLine($"KnowledgeStoreCatalog: Recovered catalog from '{bakPath}'")
                                    Return
                                End If
                            Catch
                            End Try
                        End If
                    End If

                    If isLocal Then
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Empty local catalog — no recovery source found, reset to '[]'.")
                    End If
                    Return
                End If

                ' Validate JSON structure
                Try
                    Dim token = JToken.Parse(json)
                    If token.Type <> JTokenType.Array Then
                        Dim msg = $"KnowledgeStoreCatalog: '{expanded}' does not contain a JSON array (found {token.Type}). " &
                                  "Expected format: [ {{ ""Name"": ""..."", ""SourcePath"": ""..."", ... }} ]"
                        Debug.WriteLine(msg)
                        If isLocal Then
                            Dim backup = expanded & ".bak"
                            Try : File.Copy(expanded, backup, overwrite:=True) : Catch : End Try
                            File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                            Debug.WriteLine($"KnowledgeStoreCatalog: Local catalog was invalid — backed up to '{backup}' and reset.")
                        End If
                    End If
                Catch ex As JsonReaderException
                    Dim msg = $"KnowledgeStoreCatalog: '{expanded}' contains invalid JSON: {ex.Message}. " &
                              "Expected format: [ {{ ""Name"": ""..."", ""SourcePath"": ""..."", ... }} ]"
                    Debug.WriteLine(msg)
                    If isLocal Then
                        Dim backup = expanded & ".bak"
                        Try : File.Copy(expanded, backup, overwrite:=True) : Catch : End Try
                        File.WriteAllText(expanded, "[]", System.Text.Encoding.UTF8)
                        Debug.WriteLine($"KnowledgeStoreCatalog: Local catalog was malformed — backed up to '{backup}' and reset.")
                    End If
                End Try
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreCatalog: EnsureCatalogFile error for '{rawPath}': {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Loads store definitions from a single catalog JSON file.
        ''' </summary>
        Private Shared Function LoadCatalogFile(rawPath As String, isFromCentral As Boolean) As List(Of KnowledgeStoreDefinition)
            Dim result As New List(Of KnowledgeStoreDefinition)()
            If String.IsNullOrWhiteSpace(rawPath) Then Return result

            Try
                Dim expanded = ExpandEnvironmentVariables(rawPath.Trim())
                If String.IsNullOrWhiteSpace(expanded) OrElse Not File.Exists(expanded) Then Return result

                Dim json = File.ReadAllText(expanded, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then Return result

                Dim arr = JArray.Parse(json)
                For Each token As JToken In arr
                    Try
                        If token.Type <> JTokenType.Object Then Continue For
                        Dim obj = CType(token, JObject)

                        Dim def As New KnowledgeStoreDefinition() With {
                            .Name = If(obj.Value(Of String)("Name"), "").Trim(),
                            .SourcePath = StripQuotes(If(obj.Value(Of String)("SourcePath"), "")),
                            .Owner = If(obj.Value(Of String)("Owner"), "").Trim(),
                            .Role = If(obj.Value(Of String)("Role"), "personal").Trim(),
                            .Active = If(obj("Active") IsNot Nothing, obj.Value(Of Boolean)("Active"), True),
                            .ScanSubdirectories = If(obj("ScanSubdirectories") IsNot Nothing, obj.Value(Of Boolean)("ScanSubdirectories"), True),
                            .IsFromCentralCatalog = isFromCentral
                        }
                        ' Resolve the source path
                        def.ResolvedSourcePath = ExpandEnvironmentVariables(def.SourcePath)

                        If Not String.IsNullOrWhiteSpace(def.Name) Then
                            result.Add(def)
                        End If
                    Catch
                        ' Skip malformed entries
                    End Try
                Next
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeStoreCatalog: Error loading '{rawPath}': {ex.Message}")
            End Try

            Return result
        End Function

#End Region

#Region "Save"

        ''' <summary>
        ''' Saves local store definitions to the local catalog file.
        ''' Only definitions where IsFromCentralCatalog = False are written.
        ''' Uses atomic write (tmp + move) to prevent corruption.
        ''' </summary>
        Public Shared Sub SaveLocalCatalog(definitions As List(Of KnowledgeStoreDefinition), context As ISharedContext)
            If String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal is not configured.")
            End If

            Dim expanded = ExpandEnvironmentVariables(context.INI_KnowledgeStorePathLocal.Trim())
            If String.IsNullOrWhiteSpace(expanded) Then
                Throw New InvalidOperationException("KnowledgeStorePathLocal could not be expanded.")
            End If

            Dim localDefs = definitions.Where(Function(d) Not d.IsFromCentralCatalog).ToList()

            Dim arr As New JArray()
            For Each def In localDefs
                Dim obj As New JObject()
                obj("Name") = def.Name
                obj("SourcePath") = def.SourcePath
                obj("Owner") = If(def.Owner, "")
                obj("Role") = If(def.Role, "personal")
                obj("Active") = def.Active
                obj("ScanSubdirectories") = def.ScanSubdirectories
                arr.Add(obj)
            Next

            ' Ensure directory exists
            Dim dir = Path.GetDirectoryName(expanded)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            ' Atomic write: write to .tmp, then replace original
            Dim tmpPath = expanded & ".tmp"
            File.WriteAllText(tmpPath, arr.ToString(Formatting.Indented), System.Text.Encoding.UTF8)
            If File.Exists(expanded) Then
                File.Delete(expanded)
            End If
            File.Move(tmpPath, expanded)
        End Sub

#End Region

#Region "Query Helpers"


        ''' <summary>
        ''' Returns True if at least one catalog file is configured.
        ''' </summary>
        Public Shared Function IsConfigured(context As ISharedContext) As Boolean
            Return Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePath) OrElse
                   Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStorePathLocal)
        End Function

        ''' <summary>
        ''' Returns all active store definitions.
        ''' </summary>
        Public Shared Function GetActiveStores(context As ISharedContext) As List(Of KnowledgeStoreDefinition)
            Return LoadAll(context).Where(Function(d) d.Active).ToList()
        End Function

        ''' <summary>
        ''' Returns a specific store by name (case-insensitive).
        ''' </summary>
        Public Shared Function GetStoreByName(name As String, context As ISharedContext) As KnowledgeStoreDefinition
            If String.IsNullOrWhiteSpace(name) Then Return Nothing
            Return LoadAll(context).FirstOrDefault(
                Function(d) d.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' Returns True if the current user is allowed to write (index) to this store.
        ''' Checks both Environment.UserName and the configured INI_KnowledgeStoreOwner.
        ''' </summary>
        Public Shared Function CanCurrentUserWrite(def As KnowledgeStoreDefinition, Optional context As ISharedContext = Nothing) As Boolean
            If def Is Nothing Then Return False
            If String.Equals(def.Role, "readonly", StringComparison.OrdinalIgnoreCase) Then Return False
            If String.IsNullOrWhiteSpace(def.Owner) Then Return True

            ' Match against current Windows username
            If String.Equals(def.Owner, Environment.UserName, StringComparison.OrdinalIgnoreCase) Then Return True

            ' Match against configured owner identity (if different from Windows username)
            If context IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(context.INI_KnowledgeStoreOwner) AndAlso
               String.Equals(def.Owner, context.INI_KnowledgeStoreOwner, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Returns the .redink/ metadata folder path for a given store.
        ''' Creates it if it does not exist and the user has write permission.
        ''' </summary>
        Public Shared Function GetMetadataPath(def As KnowledgeStoreDefinition, Optional createIfMissing As Boolean = True) As String
            If def Is Nothing OrElse String.IsNullOrWhiteSpace(def.ResolvedSourcePath) Then Return ""
            Dim metaDir = Path.Combine(def.ResolvedSourcePath, MetadataFolder)
            If createIfMissing AndAlso Not Directory.Exists(metaDir) Then
                Try
                    Directory.CreateDirectory(metaDir)
                Catch
                    Return ""
                End Try
            End If
            Return metaDir
        End Function

        ''' <summary>
        ''' Returns the wiki/ folder path for a given store.
        ''' Creates it if it does not exist and the user has write permission.
        ''' </summary>
        Public Shared Function GetWikiPath(def As KnowledgeStoreDefinition, Optional createIfMissing As Boolean = True) As String
            Dim metaDir = GetMetadataPath(def, createIfMissing)
            If String.IsNullOrWhiteSpace(metaDir) Then Return ""
            Dim wikiDir = Path.Combine(metaDir, WikiFolder)
            If createIfMissing AndAlso Not Directory.Exists(wikiDir) Then
                Try
                    Directory.CreateDirectory(wikiDir)
                Catch
                    Return ""
                End Try
            End If
            Return wikiDir
        End Function

        ''' <summary>
        ''' Creates a new KnowledgeStoreDefinition with the Owner defaulted from
        ''' <see cref="ISharedContext.INI_KnowledgeStoreOwner"/> (falls back to %USERNAME%).
        ''' </summary>
        Public Shared Function CreateDefinition(name As String, sourcePath As String, context As ISharedContext) As KnowledgeStoreDefinition
            Dim cleanPath = StripQuotes(sourcePath)
            Return New KnowledgeStoreDefinition() With {
                .Name = If(name, "").Trim(),
                .SourcePath = cleanPath,
                .Owner = If(String.IsNullOrWhiteSpace(context.INI_KnowledgeStoreOwner),
                            Environment.UserName,
                            context.INI_KnowledgeStoreOwner),
                .Role = "personal",
                .Active = True,
                .ScanSubdirectories = True,
                .ResolvedSourcePath = ExpandEnvironmentVariables(cleanPath)
            }
        End Function


        ''' <summary>
        ''' Strips leading/trailing quotes and whitespace from a path string.
        ''' Paths pasted from Windows Explorer or wrapped in INI values often contain quotes
        ''' that cause File.Exists, Directory.GetFiles, and Path.Combine to fail.
        ''' </summary>
        Friend Shared Function StripQuotes(path As String) As String
            If String.IsNullOrWhiteSpace(path) Then Return ""
            Return path.Trim().Trim(""""c, "'"c).Trim()
        End Function


#End Region

    End Class

End Namespace