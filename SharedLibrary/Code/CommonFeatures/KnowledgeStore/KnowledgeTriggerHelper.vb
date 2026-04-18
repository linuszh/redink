' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeTriggerHelper.vb
' Purpose:
'   Parses Knowledge Store inline triggers such as `(kb)` and `(kb:...)`,
'   resolves matching store content, and assembles prompt-ready knowledge blocks
'   for LLM injection.
'
' Responsibilities:
'   - Trigger detection:
'       * Detect simple and parameterized Knowledge Store triggers in prompts.
'       * Distinguish between load-all, store-filtered, tag-filtered, and
'         keyword-search requests.
'   - Trigger parsing:
'       * Parse `(kb)` and `(kb:...)` syntax into structured
'         `KnowledgeRequest` objects.
'       * Support store filters, tag filters, comma-separated values, and
'         free-text fallback search semantics.
'   - Prompt preparation:
'       * Strip Knowledge Store trigger syntax from the original user prompt
'         when needed.
'       * Resolve the requested content from active stores using wiki content
'         first and raw/manifests as fallback.
'       * Wrap returned knowledge in dedicated markup blocks suitable for prompt
'         injection.
'   - Safety and limits:
'       * Cap the number of included documents and total injected characters.
'       * Avoid failures when stores, manifests, or source files are missing or
'         inaccessible.
'
' Notes:
'   - This helper is the bridge between user-facing `(kb:...)` syntax and the
'     underlying Knowledge Store retrieval pipeline.
'   - It is used by prompt-processing flows such as Freestyle and other
'     Knowledge Store-enabled entry points.
' =============================================================================
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Security
Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Parses (kb:...) triggers from user prompts and resolves them to Knowledge Store content.
    ''' </summary>
    Public Class KnowledgeTriggerHelper

#Region "Constants"

        ''' <summary>Simple trigger for Knowledge Store invocation.</summary>
        Public Const KbTrigger As String = "(kb)"

        ''' <summary>Trigger prefix with parameter for filtered invocation.</summary>
        Public Const KbTriggerPrefix As String = "(kb:"

        ''' <summary>Trigger suffix.</summary>
        Private Const KbTriggerSuffix As String = ")"

        ''' <summary>Explicit store prefix inside (kb:...).</summary>
        Private Const StorePrefix As String = "store:"

        ''' <summary>Explicit tag prefix inside (kb:...).</summary>
        Private Const TagPrefix As String = "tag:"

        Private Const MaxDocuments As Integer = 10
        Private Const MaxTotalChars As Integer = 200000

        Private Const KnowledgeOpenTag As String = "<KNOWLEDGESTORE>"
        Private Const KnowledgeCloseTag As String = "</KNOWLEDGESTORE>"

#End Region

#Region "Data Model"

        ''' <summary>
        ''' Describes a parsed Knowledge Store request extracted from a user prompt.
        ''' </summary>
        Public Class KnowledgeRequest
            Public Property LoadAll As Boolean
            Public Property Tags As String()
            Public Property StoreName As String
            Public Property SearchQuery As String
            Public Property RawTrigger As String = ""
            Public Property OriginalParameter As String = ""
            Public Property HasExplicitStoreFilter As Boolean = False
            Public Property HasExplicitTagFilter As Boolean = False
        End Class

#End Region

#Region "Trigger Detection"

        ''' <summary>
        ''' Returns True if the prompt contains a Knowledge Store trigger.
        ''' </summary>
        Public Shared Function HasKnowledgeTrigger(prompt As String) As Boolean
            If String.IsNullOrWhiteSpace(prompt) Then Return False
            Return prompt.IndexOf(KbTrigger, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   prompt.IndexOf(KbTriggerPrefix, StringComparison.OrdinalIgnoreCase) >= 0
        End Function

#End Region

#Region "Trigger Parsing"

        ''' <summary>
        ''' Parses a (kb:...) trigger from the prompt and returns a KnowledgeRequest.
        ''' Returns Nothing if no trigger is found.
        ''' </summary>
        Public Shared Function TryParseKnowledgeTrigger(prompt As String) As KnowledgeRequest
            If String.IsNullOrWhiteSpace(prompt) Then Return Nothing

            ' Check for parameterized trigger first: (kb:...)
            Dim prefixIdx = prompt.IndexOf(KbTriggerPrefix, StringComparison.OrdinalIgnoreCase)
            If prefixIdx >= 0 Then
                Dim closeIdx = prompt.IndexOf(KbTriggerSuffix, prefixIdx + KbTriggerPrefix.Length, StringComparison.OrdinalIgnoreCase)
                If closeIdx > prefixIdx Then
                    Dim rawTrigger = prompt.Substring(prefixIdx, closeIdx - prefixIdx + KbTriggerSuffix.Length)
                    Dim parameter = prompt.Substring(prefixIdx + KbTriggerPrefix.Length, closeIdx - prefixIdx - KbTriggerPrefix.Length).Trim()

                    If String.IsNullOrWhiteSpace(parameter) Then
                        Return New KnowledgeRequest() With {
                            .LoadAll = True,
                            .RawTrigger = rawTrigger,
                            .OriginalParameter = ""
                        }
                    End If

                    Dim parsed = ParseParameter(parameter, rawTrigger)
                    parsed.OriginalParameter = parameter
                    Return parsed
                End If
            End If

            ' Check for simple trigger: (kb)
            Dim simpleIdx = prompt.IndexOf(KbTrigger, StringComparison.OrdinalIgnoreCase)
            If simpleIdx >= 0 Then
                Return New KnowledgeRequest() With {
                    .LoadAll = True,
                    .RawTrigger = KbTrigger,
                    .OriginalParameter = ""
                }
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Parses the parameter string inside (kb:...).
        ''' Supports: "store:X", "tag:Y", "store:X tag:Y", plain CSV tags/store names.
        ''' Store and tag values may be quoted to allow spaces: store:"My Store"
        ''' </summary>
        Private Shared Function ParseParameter(parameter As String, rawTrigger As String) As KnowledgeRequest
            Dim req As New KnowledgeRequest() With {
                .RawTrigger = rawTrigger,
                .OriginalParameter = parameter
            }

            Dim hasStorePrefix = parameter.IndexOf(StorePrefix, StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasTagPrefix = parameter.IndexOf(TagPrefix, StringComparison.OrdinalIgnoreCase) >= 0

            req.HasExplicitStoreFilter = hasStorePrefix
            req.HasExplicitTagFilter = hasTagPrefix

            If hasStorePrefix OrElse hasTagPrefix Then
                ' Tokenize respecting quoted values: store:"My Store" tag:"Some Tag" free text
                Dim tokens = TokenizeWithQuotes(parameter)
                For Each token In tokens
                    If token.StartsWith(StorePrefix, StringComparison.OrdinalIgnoreCase) Then
                        req.StoreName = Unquote(token.Substring(StorePrefix.Length)).Trim()
                    ElseIf token.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Dim tagVal = Unquote(token.Substring(TagPrefix.Length)).Trim()
                        Dim tagParts = tagVal.Split(","c).
                            Select(Function(t) t.Trim()).
                            Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                            ToArray()
                        req.Tags = tagParts
                    Else
                        req.SearchQuery = If(String.IsNullOrWhiteSpace(req.SearchQuery), token, req.SearchQuery & " " & token)
                    End If
                Next

                Return req
            End If

            Dim csvParts = parameter.Split(","c).
                Select(Function(p) p.Trim()).
                Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                ToArray()

            If csvParts.Length > 0 Then
                req.Tags = csvParts
                req.SearchQuery = parameter
            End If

            Return req
        End Function

        ''' <summary>
        ''' Splits a string by spaces but keeps quoted segments (double quotes) as single tokens.
        ''' E.g. <c>store:"My Store" hello</c> → <c>{"store:My Store", "hello"}</c>.
        ''' The prefix (e.g. "store:") is preserved and the quotes are stripped from the value portion.
        ''' </summary>
        Public Shared Function TokenizeWithQuotes(input As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(input) Then Return result

            ' Match: prefix:"quoted value" | prefix:'quoted value' | unquoted_token
            Dim pattern As String = "(\w+:""[^""]*""|\w+:'[^']*'|\S+)"
            For Each m As Match In Regex.Matches(input, pattern)
                Dim token = m.Value

                ' For prefixed tokens like store:"My Store", strip the quotes from the value part
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 AndAlso colonIdx < token.Length - 1 Then
                    Dim prefix = token.Substring(0, colonIdx + 1)
                    Dim value = token.Substring(colonIdx + 1)
                    result.Add(prefix & Unquote(value))
                Else
                    result.Add(token)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Removes surrounding double or single quotes from a string if present.
        ''' </summary>
        Private Shared Function Unquote(value As String) As String
            If String.IsNullOrEmpty(value) OrElse value.Length < 2 Then Return value
            If (value.StartsWith("""", StringComparison.Ordinal) AndAlso value.EndsWith("""", StringComparison.Ordinal)) OrElse
               (value.StartsWith("'", StringComparison.Ordinal) AndAlso value.EndsWith("'", StringComparison.Ordinal)) Then
                Return value.Substring(1, value.Length - 2)
            End If
            Return value
        End Function

        ''' <summary>
        ''' Removes the Knowledge Store trigger from the prompt and returns the cleaned text.
        ''' </summary>
        Public Shared Function StripKnowledgeTrigger(prompt As String, request As KnowledgeRequest) As String
            If request Is Nothing OrElse String.IsNullOrWhiteSpace(request.RawTrigger) Then Return prompt
            Dim idx = prompt.IndexOf(request.RawTrigger, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then Return prompt
            Dim result = (prompt.Substring(0, idx) & prompt.Substring(idx + request.RawTrigger.Length)).Trim()
            result = Regex.Replace(result, " {2,}", " ")
            Return result
        End Function

#End Region

#Region "Knowledge Resolution"

        Private Shared Function ShouldUseSemanticFirst(request As KnowledgeRequest) As Boolean
            If request Is Nothing Then Return False
            If request.LoadAll Then Return False
            If String.IsNullOrWhiteSpace(request.SearchQuery) Then Return False

            If request.HasExplicitTagFilter Then Return False

            Dim parameter = If(request.OriginalParameter, "").Trim()

            If String.IsNullOrWhiteSpace(parameter) Then Return False

            If parameter.IndexOf(","c) >= 0 Then Return False

            If parameter.IndexOf(" "c) >= 0 Then Return True

            Return False
        End Function

        ''' <summary>
        ''' Resolves a KnowledgeRequest to assembled document content ready for LLM injection.
        ''' Synchronous compatibility wrapper over the async implementation.
        ''' </summary>
        Public Shared Function ResolveKnowledge(
                request As KnowledgeRequest,
                context As ISharedContext) As (Content As String, StatusMessage As String)

            Return ResolveKnowledgeAsync(request, context).GetAwaiter().GetResult()
        End Function

        Private Shared Function FindWikiPagePathForEntry(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                         entry As KnowledgeStoreManager.KnowledgeEntry) As String
            If store Is Nothing OrElse entry Is Nothing Then Return ""

            Dim wikiDir = Path.Combine(store.ResolvedSourcePath, KnowledgeStoreCatalog.MetadataFolder, KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiDir) Then Return ""

            Dim normalizedEntryPath = NormalizePathForCompare(entry.FilePath)
            Dim normalizedTitle = NormalizeKey(If(entry.Title, ""))

            For Each file In Directory.GetFiles(wikiDir, "*.md", SearchOption.AllDirectories)
                Dim name = Path.GetFileName(file)
                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    Dim content = System.IO.File.ReadAllText(file, Encoding.UTF8)
                    Dim frontMatter = GetFrontMatter(content)

                    Dim pageTitle = GetFrontMatterScalar(frontMatter, "title")
                    If NormalizeKey(pageTitle) = normalizedTitle AndAlso Not String.IsNullOrWhiteSpace(normalizedTitle) Then
                        Return file
                    End If

                    Dim sourcePaths = GetFrontMatterList(frontMatter, "source_paths")
                    For Each sourcePath In sourcePaths
                        If NormalizePathForCompare(sourcePath) = normalizedEntryPath AndAlso
                           Not String.IsNullOrWhiteSpace(normalizedEntryPath) Then
                            Return file
                        End If
                    Next
                Catch
                End Try
            Next

            Return ""
        End Function

        Private Shared Function GetFrontMatter(content As String) As String
            If String.IsNullOrWhiteSpace(content) Then Return ""
            Dim normalized = content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If Not normalized.StartsWith("---" & vbLf, StringComparison.Ordinal) Then Return ""

            Dim endMarker = vbLf & "---" & vbLf
            Dim endPos = normalized.IndexOf(endMarker, 4, StringComparison.Ordinal)
            If endPos < 0 Then Return ""

            Return normalized.Substring(4, endPos - 4)
        End Function

        Private Shared Function GetFrontMatterScalar(frontMatter As String, key As String) As String
            If String.IsNullOrWhiteSpace(frontMatter) Then Return ""

            For Each rawLine In frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim line = rawLine.Trim()
                If line.StartsWith(key & ":", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = line.Substring(key.Length + 1).Trim()
                    If value.StartsWith("""", StringComparison.Ordinal) AndAlso
                       value.EndsWith("""", StringComparison.Ordinal) AndAlso
                       value.Length >= 2 Then
                        value = value.Substring(1, value.Length - 2)
                    End If
                    Return value.Trim()
                End If
            Next

            Return ""
        End Function

        Private Shared Function GetFrontMatterList(frontMatter As String, key As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(frontMatter) Then Return result

            Dim lines = frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim inTarget As Boolean = False

            For Each rawLine In lines
                Dim line = rawLine.TrimEnd()

                If line.Trim().StartsWith(key & ":", StringComparison.OrdinalIgnoreCase) Then
                    inTarget = True
                    Continue For
                End If

                If inTarget Then
                    Dim trimmed = line.Trim()
                    If trimmed.StartsWith("- ", StringComparison.Ordinal) Then
                        Dim value = trimmed.Substring(2).Trim()
                        If value.StartsWith("""", StringComparison.Ordinal) AndAlso
                               value.EndsWith("""", StringComparison.Ordinal) AndAlso
                               value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2)
                        End If
                        ' Strip surrounding single quotes (YAML single-quoted scalars)
                        If value.StartsWith("'", StringComparison.Ordinal) AndAlso
                               value.EndsWith("'", StringComparison.Ordinal) AndAlso
                               value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2).Replace("''", "'")
                        End If
                        If Not String.IsNullOrWhiteSpace(value) Then result.Add(value)
                    ElseIf trimmed.Contains(":"c) Then
                        Exit For
                    End If
                End If
            Next

            Return result
        End Function

        Private Shared Function NormalizePathForCompare(pathValue As String) As String
            If String.IsNullOrWhiteSpace(pathValue) Then Return ""
            Try
                Dim expanded = SharedMethods.ExpandEnvironmentVariables(pathValue)
                Return Path.GetFullPath(expanded).Trim().ToUpperInvariant()
            Catch
                Return pathValue.Trim().ToUpperInvariant()
            End Try
        End Function

        Private Shared Function NormalizeKey(value As String) As String
            Return Regex.Replace(If(value, "").ToLowerInvariant(), "[^a-z0-9]+", "").Trim()
        End Function

#End Region

#Region "Knowledge Resolution"

        ''' <summary>
        ''' Resolves a KnowledgeRequest to assembled document content ready for LLM injection.
        ''' Uses the multi-store catalog and manifest infrastructure. If a search query is provided,
        ''' it uses semantic vector search for ranking.
        ''' </summary>
        Public Shared Async Function ResolveKnowledgeAsync(
                request As KnowledgeRequest,
                context As ISharedContext) As Task(Of (Content As String, StatusMessage As String))

            If request Is Nothing Then Return ("", "")

            If Not KnowledgeStoreCatalog.IsConfigured(context) Then
                Return ("", "Knowledge Store is not configured.")
            End If

            Dim stores = KnowledgeStoreCatalog.GetActiveStores(context)
            If stores.Count = 0 Then
                Return ("", "No active Knowledge Stores found.")
            End If

            ' Resolve store name from request (explicit or by matching CSV values against store names)
            Dim targetStores As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) = Nothing
            Dim resolvedStoreFromTag As Boolean = False

            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                Dim matches = KnowledgeStoreCatalog.GetStoresByName(request.StoreName, context)
                If matches.Count > 0 Then
                    targetStores = matches
                Else
                    Return ("", $"Knowledge Store '{request.StoreName}' not found.")
                End If
            ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length = 1 AndAlso Not request.LoadAll Then
                Dim singleVal = request.Tags(0)
                Dim storeMatches = KnowledgeStoreCatalog.GetStoresByName(singleVal, context)

                If storeMatches.Count > 0 Then
                    targetStores = storeMatches
                    resolvedStoreFromTag = True
                    request = New KnowledgeRequest() With {
            .LoadAll = True,
            .StoreName = singleVal,
            .RawTrigger = request.RawTrigger
        }
                End If
            End If

            If targetStores Is Nothing Then
                targetStores = stores
            End If

            ' =================================================================
            ' 1. NEW SEMANTIC VECTOR SEARCH BRANCH
            ' If the user provided a search term (rather than just loading all docs or explicit tags),
            ' bypass legacy metadata scanning and route directly to the Embedding engine!
            ' =================================================================
            If ShouldUseSemanticFirst(request) Then
                Dim queryText = request.SearchQuery.Trim()

                If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                    queryText = $"store:{request.StoreName} {queryText}"
                End If

                Dim semanticContext = Await KnowledgeQueryService.ResolveAndBuildAsync(
                    query:=queryText,
                    context:=context,
                    maxResults:=MaxDocuments,
                    maxTotalChars:=MaxTotalChars).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(semanticContext) Then
                    Return ("", $"No semantic knowledge found matching '{request.SearchQuery}'.")
                End If

                Dim semanticMsg = $"Loaded semantic knowledge for '{request.SearchQuery}'"
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                    semanticMsg &= $" in store '{request.StoreName}'"
                End If

                Return (semanticContext, semanticMsg & ".")
            End If


            ' =================================================================
            ' 2. LEGACY LOGIC / METADATA SEARCH (Fallback or LoadAll logic)
            ' =================================================================
            Dim allMatches As New List(Of Tuple(Of KnowledgeStoreManager.KnowledgeEntry, String, String))() ' (entry, storeName, wikiPath)

            For Each store In targetStores
                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim wikiDir = KnowledgeStoreCatalog.GetWikiPath(store, createIfMissing:=False)

                For Each entry In manifest.Entries
                    If String.IsNullOrWhiteSpace(entry.FilePath) Then Continue For

                    Dim include As Boolean = False

                    If request.LoadAll Then
                        include = True
                    ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 Then
                        ' Try tag match
                        If entry.Tags IsNot Nothing Then
                            For Each reqTag In request.Tags
                                If entry.Tags.Any(Function(t) t.Equals(reqTag, StringComparison.OrdinalIgnoreCase)) Then
                                    include = True
                                    Exit For
                                End If
                            Next
                        End If

                        ' If no tag match, try keyword/title search
                        If Not include AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                            Dim sq = request.SearchQuery.ToLowerInvariant()
                            If If(entry.Title, "").ToLowerInvariant().Contains(sq) Then include = True
                            If Not include AndAlso entry.Keywords IsNot Nothing Then
                                If entry.Keywords.Any(Function(k) k.IndexOf(sq, StringComparison.OrdinalIgnoreCase) >= 0) Then
                                    include = True
                                End If
                            End If
                            If Not include AndAlso If(entry.Summary, "").ToLowerInvariant().Contains(sq) Then
                                include = True
                            End If
                        End If
                    Else
                        include = True
                    End If

                    If include Then
                        Dim wikiPath As String = FindWikiPagePathForEntry(store, entry)
                        allMatches.Add(Tuple.Create(entry, store.Name, wikiPath))
                    End If
                Next
            Next

            If allMatches.Count = 0 Then
                If Not request.LoadAll AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                    Dim queryText = request.SearchQuery.Trim()

                    If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                        queryText = $"store:{request.StoreName} {queryText}"
                    End If

                    Dim semanticContext = Await KnowledgeQueryService.ResolveAndBuildAsync(
                        query:=queryText,
                        context:=context,
                        maxResults:=MaxDocuments,
                        maxTotalChars:=MaxTotalChars).ConfigureAwait(False)

                    If Not String.IsNullOrWhiteSpace(semanticContext) Then
                        Dim semanticMsg = $"Loaded semantic knowledge for '{request.SearchQuery}'"
                        If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                            semanticMsg &= $" in store '{request.StoreName}'"
                        End If

                        Return (semanticContext, semanticMsg & ".")
                    End If
                End If

                Dim filterDesc = ""
                If request.Tags IsNot Nothing Then filterDesc = String.Join(", ", request.Tags)
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then filterDesc = $"store '{request.StoreName}'"
                If String.IsNullOrWhiteSpace(filterDesc) AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                    filterDesc = request.SearchQuery
                End If

                Return ("", $"No Knowledge Store documents found matching '{filterDesc}'.")
            End If

            ' Limit and assemble
            Dim entriesToLoad = allMatches.Take(MaxDocuments).ToList()
            Dim richMatches As New List(Of KnowledgeQueryService.KnowledgeMatch)()

            For Each item In entriesToLoad
                Dim entry = item.Item1
                Dim sName = item.Item2
                Dim wikiPath = item.Item3

                richMatches.Add(New KnowledgeQueryService.KnowledgeMatch With {
                    .Entry = entry,
                    .StoreName = sName,
                    .WikiPagePath = wikiPath,
                    .Title = If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "Unknown"))),
                    .Summary = If(entry.Summary, ""),
                    .SourcePath = ResolveSourcePath(entry.FilePath, stores.FirstOrDefault(Function(s) s.Name.Equals(sName, StringComparison.OrdinalIgnoreCase))),
                    .MatchReason = "kb-trigger"
                })
            Next

            Dim richContext = KnowledgeQueryService.BuildKnowledgeContext(richMatches, MaxTotalChars)

            Dim statusMsg = $"Loaded {entriesToLoad.Count} document(s)"
            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                statusMsg &= $" from store '{request.StoreName}'"
            End If
            If request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 AndAlso Not resolvedStoreFromTag Then
                statusMsg &= $" for '{String.Join(", ", request.Tags)}'"
            End If
            statusMsg &= "."

            Return (richContext, statusMsg)
        End Function

#End Region

#Region "Helpers"

        ''' <summary>
        ''' Resolves a potentially relative source file path against the store's root directory.
        ''' Returns the original path if it is already absolute or the store is unavailable.
        ''' </summary>
        Private Shared Function ResolveSourcePath(filePath As String, store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""
            Try
                If Path.IsPathRooted(filePath) AndAlso File.Exists(filePath) Then
                    Return filePath
                End If
                If store IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                    Dim combined = Path.GetFullPath(Path.Combine(store.ResolvedSourcePath, filePath))
                    If File.Exists(combined) Then Return combined
                End If
                Return filePath
            Catch
                Return filePath
            End Try
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(name.Where(Function(c) Not invalid.Contains(c)).ToArray())
            If cleaned.Length > 100 Then cleaned = cleaned.Substring(0, 100)
            Return If(String.IsNullOrWhiteSpace(cleaned), "unnamed", cleaned.Trim())
        End Function

#End Region

    End Class

End Namespace