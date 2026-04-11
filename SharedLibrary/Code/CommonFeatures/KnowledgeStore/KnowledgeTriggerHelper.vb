' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeTriggerHelper.vb
' Purpose: Parses (kb:...) trigger patterns from user prompts, resolves matching
'          Knowledge Store entries, and assembles document content for LLM injection.
'
' Trigger Syntax:
'  - (kb)                     → Load ALL documents from all active stores
'  - (kb:tag)                 → Load documents matching the specified tag (or store name)
'  - (kb:tag1,tag2)           → Load documents matching ANY of the specified tags/store names
'  - (kb:store:Contracts)     → Explicit store filter
'  - (kb:tag:employment)      → Explicit tag filter across all stores
'  - (kb:store:X tag:Y)       → Store + tag combined
'  - (kb:search term)         → Fallback: search by title/keyword/summary
'
' Resolution order for (kb:xxx):
'  1. Check if "xxx" is a store name → scope to that store
'  2. Check if "xxx" matches tags → filter by tag
'  3. Fall back to keyword search across all stores
'
' Architecture / How it works:
'  - TryParseKnowledgeTrigger() extracts the trigger from the prompt.
'  - StripKnowledgeTrigger() removes the trigger, returning clean instruction.
'  - ResolveKnowledge() loads manifests from matching stores, filters, and
'    assembles content from wiki pages (preferred) or raw documents (fallback).
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
        Private Const DocOpenTag As String = "<KSDOCUMENT title=""{0}"" store=""{1}"">"
        Private Const DocCloseTag As String = "</KSDOCUMENT>"

#End Region

#Region "Data Model"

        ''' <summary>
        ''' Describes a parsed Knowledge Store request extracted from a user prompt.
        ''' </summary>
        Public Class KnowledgeRequest
            ''' <summary>If True, load all documents (unfiltered).</summary>
            Public Property LoadAll As Boolean

            ''' <summary>Tags to filter by. Nothing if no tags specified.</summary>
            Public Property Tags As String()

            ''' <summary>Store name to scope to. Nothing if not specified.</summary>
            Public Property StoreName As String

            ''' <summary>Free-text search query (fallback). Nothing if tags or store name are specified.</summary>
            Public Property SearchQuery As String

            ''' <summary>The raw trigger text as it appeared in the prompt (for removal).</summary>
            Public Property RawTrigger As String
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
                            .RawTrigger = rawTrigger
                        }
                    End If

                    Return ParseParameter(parameter, rawTrigger)
                End If
            End If

            ' Check for simple trigger: (kb)
            Dim simpleIdx = prompt.IndexOf(KbTrigger, StringComparison.OrdinalIgnoreCase)
            If simpleIdx >= 0 Then
                Return New KnowledgeRequest() With {
                    .LoadAll = True,
                    .RawTrigger = KbTrigger
                }
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' Parses the parameter string inside (kb:...).
        ''' Supports: "store:X", "tag:Y", "store:X tag:Y", plain CSV tags/store names.
        ''' </summary>
        Private Shared Function ParseParameter(parameter As String, rawTrigger As String) As KnowledgeRequest
            Dim req As New KnowledgeRequest() With {.RawTrigger = rawTrigger}

            ' Check for explicit prefixes
            Dim hasStorePrefix = parameter.IndexOf(StorePrefix, StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasTagPrefix = parameter.IndexOf(TagPrefix, StringComparison.OrdinalIgnoreCase) >= 0

            If hasStorePrefix OrElse hasTagPrefix Then
                ' Parse explicit prefixes (e.g., "store:Contracts tag:NDA")
                Dim parts = parameter.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                For Each part In parts
                    If part.StartsWith(StorePrefix, StringComparison.OrdinalIgnoreCase) Then
                        req.StoreName = part.Substring(StorePrefix.Length).Trim()
                    ElseIf part.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase) Then
                        Dim tagVal = part.Substring(TagPrefix.Length).Trim()
                        Dim tagParts = tagVal.Split(","c).
                            Select(Function(t) t.Trim()).
                            Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                            ToArray()
                        req.Tags = tagParts
                    Else
                        ' Treat as search query
                        req.SearchQuery = If(String.IsNullOrWhiteSpace(req.SearchQuery), part, req.SearchQuery & " " & part)
                    End If
                Next
                Return req
            End If

            ' No explicit prefixes — comma-separated values that could be store names or tags
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

        ''' <summary>
        ''' Resolves a KnowledgeRequest to assembled document content ready for LLM injection.
        ''' Uses the multi-store catalog and manifest infrastructure.
        ''' </summary>
        Public Shared Function ResolveKnowledge(
                request As KnowledgeRequest,
                context As ISharedContext) As (Content As String, StatusMessage As String)

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
                ' Explicit store: prefix
                Dim match = stores.FirstOrDefault(Function(s) s.Name.Equals(request.StoreName, StringComparison.OrdinalIgnoreCase))
                If match IsNot Nothing Then
                    targetStores = New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) From {match}
                Else
                    Return ("", $"Knowledge Store '{request.StoreName}' not found.")
                End If
            ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length = 1 AndAlso Not request.LoadAll Then
                ' Single value in (kb:xxx) — check if it's a store name first
                Dim singleVal = request.Tags(0)
                Dim storeMatch = stores.FirstOrDefault(Function(s) s.Name.Equals(singleVal, StringComparison.OrdinalIgnoreCase))
                If storeMatch IsNot Nothing Then
                    targetStores = New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) From {storeMatch}
                    resolvedStoreFromTag = True
                    ' Clear tags since it was actually a store name
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

            ' Collect matching entries across target stores
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
                        Dim wikiPath As String = ""
                        If Not String.IsNullOrWhiteSpace(wikiDir) Then
                            Dim safeName = SanitizeFileName(If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "unknown"))))
                            Dim candidate = Path.Combine(wikiDir, safeName & "-summary.md")
                            If File.Exists(candidate) Then wikiPath = candidate
                        End If
                        allMatches.Add(Tuple.Create(entry, store.Name, wikiPath))
                    End If
                Next
            Next

            If allMatches.Count = 0 Then
                Dim filterDesc = ""
                If request.Tags IsNot Nothing Then filterDesc = String.Join(", ", request.Tags)
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then filterDesc = $"store '{request.StoreName}'"
                Return ("", $"No Knowledge Store documents found matching '{filterDesc}'.")
            End If

            ' Limit and assemble
            Dim entriesToLoad = allMatches.Take(MaxDocuments).ToList()
            Dim sb As New StringBuilder()
            sb.AppendLine(KnowledgeOpenTag)

            Dim totalChars As Integer = 0
            Dim loadedCount As Integer = 0

            For Each item In entriesToLoad
                Dim entry = item.Item1
                Dim sName = item.Item2
                Dim wikiPath = item.Item3

                ' Prefer wiki summary if available
                Dim content As String = ""
                If Not String.IsNullOrWhiteSpace(wikiPath) Then
                    Try
                        content = File.ReadAllText(wikiPath, System.Text.Encoding.UTF8)
                    Catch
                    End Try
                End If

                ' Fall back to raw document
                If String.IsNullOrWhiteSpace(content) Then
                    content = KnowledgeStoreManager.GetContent(entry)
                End If

                If String.IsNullOrWhiteSpace(content) Then Continue For

                If totalChars + content.Length > MaxTotalChars Then
                    Dim remaining = MaxTotalChars - totalChars
                    If remaining > 500 Then
                        content = content.Substring(0, remaining) & vbCrLf & "... [truncated]"
                    Else
                        Exit For
                    End If
                End If

                Dim title = If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "Unknown")))
                sb.AppendLine(String.Format(DocOpenTag, SecurityElement.Escape(title), SecurityElement.Escape(sName)))
                sb.AppendLine(content)
                sb.AppendLine(DocCloseTag)

                totalChars += content.Length
                loadedCount += 1
            Next

            sb.AppendLine(KnowledgeCloseTag)

            If loadedCount = 0 Then
                Return ("", "Knowledge Store documents could not be read.")
            End If

            Dim statusMsg = $"Loaded {loadedCount} document(s)"
            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                statusMsg &= $" from store '{request.StoreName}'"
            End If
            If request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 AndAlso Not resolvedStoreFromTag Then
                statusMsg &= $" for '{String.Join(", ", request.Tags)}'"
            End If
            statusMsg &= "."

            Return (sb.ToString(), statusMsg)
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
                ' Explicit store: prefix
                Dim match = stores.FirstOrDefault(Function(s) s.Name.Equals(request.StoreName, StringComparison.OrdinalIgnoreCase))
                If match IsNot Nothing Then
                    targetStores = New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) From {match}
                Else
                    Return ("", $"Knowledge Store '{request.StoreName}' not found.")
                End If
            ElseIf request.Tags IsNot Nothing AndAlso request.Tags.Length = 1 AndAlso Not request.LoadAll Then
                ' Single value in (kb:xxx) — check if it's a store name first
                Dim singleVal = request.Tags(0)
                Dim storeMatch = stores.FirstOrDefault(Function(s) s.Name.Equals(singleVal, StringComparison.OrdinalIgnoreCase))
                If storeMatch IsNot Nothing Then
                    targetStores = New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition) From {storeMatch}
                    resolvedStoreFromTag = True
                    ' Clear tags since it was actually a store name
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
            If Not request.LoadAll AndAlso request.Tags Is Nothing AndAlso Not String.IsNullOrWhiteSpace(request.SearchQuery) Then
                Dim searchResults As New List(Of KnowledgeSearchResult)()

                For Each store In targetStores
                    Dim storeResults = Await KnowledgeEmbeddingService.SearchAsync(store.ResolvedSourcePath, request.SearchQuery, MaxDocuments, context)
                    searchResults.AddRange(storeResults)
                Next

                If searchResults.Count = 0 Then
                    Return ("", $"No semantic knowledge found matching '{request.SearchQuery}'.")
                End If

                ' Order across all stores by score
                searchResults = searchResults.OrderByDescending(Function(x) x.Score).Take(MaxDocuments).ToList()

                Dim semanticSb As New StringBuilder()
                semanticSb.AppendLine(KnowledgeOpenTag)
                Dim tChars As Integer = 0

                For Each res In searchResults
                    If String.IsNullOrWhiteSpace(res.TextChunk) Then Continue For
                    If tChars + res.TextChunk.Length > MaxTotalChars Then Exit For

                    Dim title = If(String.IsNullOrWhiteSpace(res.FilePath), "Unknown", Path.GetFileNameWithoutExtension(res.FilePath))

                    ' Include the search ranking score as metadata context instead of the store name
                    semanticSb.AppendLine(String.Format(DocOpenTag, SecurityElement.Escape(title), SecurityElement.Escape($"RankScore: {res.Score:F2}")))
                    semanticSb.AppendLine(res.TextChunk)
                    semanticSb.AppendLine(DocCloseTag)
                    tChars += res.TextChunk.Length
                Next
                semanticSb.AppendLine(KnowledgeCloseTag)

                Dim semanticMsg = $"Loaded top {searchResults.Count} matches via Semantic Search"
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then semanticMsg &= $" in store '{request.StoreName}'"
                Return (semanticSb.ToString(), semanticMsg & ".")
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
                        Dim wikiPath As String = ""
                        If Not String.IsNullOrWhiteSpace(wikiDir) Then
                            Dim safeName = SanitizeFileName(If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "unknown"))))
                            Dim candidate = Path.Combine(wikiDir, safeName & "-summary.md")
                            If File.Exists(candidate) Then wikiPath = candidate
                        End If
                        allMatches.Add(Tuple.Create(entry, store.Name, wikiPath))
                    End If
                Next
            Next

            If allMatches.Count = 0 Then
                Dim filterDesc = ""
                If request.Tags IsNot Nothing Then filterDesc = String.Join(", ", request.Tags)
                If Not String.IsNullOrWhiteSpace(request.StoreName) Then filterDesc = $"store '{request.StoreName}'"
                Return ("", $"No Knowledge Store documents found matching '{filterDesc}'.")
            End If

            ' Limit and assemble
            Dim entriesToLoad = allMatches.Take(MaxDocuments).ToList()
            Dim sb As New StringBuilder()
            sb.AppendLine(KnowledgeOpenTag)

            Dim totalChars As Integer = 0
            Dim loadedCount As Integer = 0

            For Each item In entriesToLoad
                Dim entry = item.Item1
                Dim sName = item.Item2
                Dim wikiPath = item.Item3

                ' Prefer wiki summary if available
                Dim content As String = ""
                If Not String.IsNullOrWhiteSpace(wikiPath) Then
                    Try
                        content = File.ReadAllText(wikiPath, System.Text.Encoding.UTF8)
                    Catch
                    End Try
                End If

                ' Fall back to raw document
                If String.IsNullOrWhiteSpace(content) Then
                    content = KnowledgeStoreManager.GetContent(entry)
                End If

                If String.IsNullOrWhiteSpace(content) Then Continue For

                If totalChars + content.Length > MaxTotalChars Then
                    Dim remaining = MaxTotalChars - totalChars
                    If remaining > 500 Then
                        content = content.Substring(0, remaining) & vbCrLf & "... [truncated]"
                    Else
                        Exit For
                    End If
                End If

                Dim title = If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "Unknown")))
                sb.AppendLine(String.Format(DocOpenTag, SecurityElement.Escape(title), SecurityElement.Escape(sName)))
                sb.AppendLine(content)
                sb.AppendLine(DocCloseTag)

                totalChars += content.Length
                loadedCount += 1
            Next

            sb.AppendLine(KnowledgeCloseTag)

            If loadedCount = 0 Then
                Return ("", "Knowledge Store documents could not be read.")
            End If

            Dim statusMsg = $"Loaded {loadedCount} document(s)"
            If Not String.IsNullOrWhiteSpace(request.StoreName) Then
                statusMsg &= $" from store '{request.StoreName}'"
            End If
            If request.Tags IsNot Nothing AndAlso request.Tags.Length > 0 AndAlso Not resolvedStoreFromTag Then
                statusMsg &= $" for '{String.Join(", ", request.Tags)}'"
            End If
            statusMsg &= "."

            Return (sb.ToString(), statusMsg)
        End Function

#End Region

#Region "Helpers"

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(name.Where(Function(c) Not invalid.Contains(c)).ToArray())
            If cleaned.Length > 100 Then cleaned = cleaned.Substring(0, 100)
            Return If(String.IsNullOrWhiteSpace(cleaned), "unnamed", cleaned.Trim())
        End Function

#End Region

    End Class

End Namespace