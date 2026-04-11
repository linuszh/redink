' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeQueryService.vb
' Purpose: Resolves Knowledge Store queries by searching across all active stores'
'          manifests, wiki pages, and raw documents, then assembles context text
'          for injection into LLM prompts.
'
' Architecture / How it works:
'  - ResolveQuery scans all active stores' manifests for keyword/tag matches.
'  - BuildKnowledgeContext first tries to read wiki summary pages (pre-compiled
'    knowledge); falls back to raw document content via KnowledgeStoreManager.GetContent.
'  - Tag-based filtering is supported: queries prefixed with "tag:" restrict
'    results to entries matching specific tags.
'  - Store-scoped queries: prefix with "store:Name" to restrict to one KB.
'
' Future enhancements (planned for next iteration):
'  - Read index.md first to let the LLM navigate the wiki structure.
'  - Use embeddings for semantic search instead of keyword matching.
'  - Support LLM re-ranking of candidates before content retrieval.
'  - File query answers back into the wiki as new pages.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Resolves knowledge queries against all active Knowledge Stores and builds
    ''' context blocks for LLM prompt injection.
    ''' </summary>
    Public Class KnowledgeQueryService

        ''' <summary>
        ''' Represents a scored document match from a knowledge query.
        ''' </summary>
        Public Class KnowledgeMatch
            Public Property Entry As KnowledgeStoreManager.KnowledgeEntry
            Public Property Score As Double = 0.0
            Public Property StoreName As String = ""
            Public Property WikiSummaryPath As String = ""
        End Class

        ''' <summary>
        ''' Resolves a query against all active Knowledge Stores, returning scored matches
        ''' ordered by relevance (highest first).
        ''' </summary>
        ''' <param name="query">
        ''' The user query string. Supported prefixes:
        '''   "tag:tagname" — filter by tag
        '''   "store:storename" — restrict to a specific KB
        '''   "tag:tagname store:storename" — both
        ''' </param>
        ''' <param name="context">Shared context for loading catalogs and manifests.</param>
        ''' <param name="maxResults">Maximum number of results to return.</param>
        Public Shared Function ResolveQuery(query As String, context As ISharedContext,
                                            Optional maxResults As Integer = 5) As List(Of KnowledgeMatch)

            Dim results As New List(Of KnowledgeMatch)()
            If String.IsNullOrWhiteSpace(query) OrElse context Is Nothing Then Return results
            If Not KnowledgeStoreCatalog.IsConfigured(context) Then Return results

            ' Parse prefixes
            Dim tagFilter As String = ""
            Dim storeFilter As String = ""
            Dim searchTerms As String = query.Trim()

            ' Extract tag: prefix
            If searchTerms.StartsWith("tag:", StringComparison.OrdinalIgnoreCase) Then
                Dim parts = searchTerms.Substring(4).Trim().Split(New Char() {" "c}, 2, StringSplitOptions.RemoveEmptyEntries)
                If parts.Length > 0 Then
                    tagFilter = parts(0).Trim().ToLowerInvariant()
                    searchTerms = If(parts.Length > 1, parts(1).Trim(), "")
                End If
            End If

            ' Extract store: prefix
            If searchTerms.StartsWith("store:", StringComparison.OrdinalIgnoreCase) Then
                Dim parts = searchTerms.Substring(6).Trim().Split(New Char() {" "c}, 2, StringSplitOptions.RemoveEmptyEntries)
                If parts.Length > 0 Then
                    storeFilter = parts(0).Trim()
                    searchTerms = If(parts.Length > 1, parts(1).Trim(), "")
                End If
            End If

            ' Tokenize search terms
            Dim keywords As String() = If(String.IsNullOrWhiteSpace(searchTerms),
                                          Array.Empty(Of String)(),
                                          searchTerms.ToLowerInvariant().Split(New Char() {" "c, ","c, ";"c},
                                                                                StringSplitOptions.RemoveEmptyEntries))

            ' Iterate all active stores
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(context)

            For Each store In stores
                ' Apply store filter
                If Not String.IsNullOrWhiteSpace(storeFilter) AndAlso
                   Not store.Name.Equals(storeFilter, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim wikiDir = KnowledgeStoreCatalog.GetWikiPath(store, createIfMissing:=False)

                For Each entry In manifest.Entries
                    If String.IsNullOrWhiteSpace(entry.FilePath) Then Continue For

                    ' Tag filter
                    If Not String.IsNullOrWhiteSpace(tagFilter) Then
                        If entry.Tags Is Nothing OrElse
                           Not entry.Tags.Any(Function(t) t.Equals(tagFilter, StringComparison.OrdinalIgnoreCase)) Then
                            Continue For
                        End If
                    End If

                    ' Score the entry
                    Dim score As Double = 0.0

                    If keywords.Length = 0 AndAlso Not String.IsNullOrWhiteSpace(tagFilter) Then
                        score = 1.0 ' Tag-only query
                    Else
                        Dim titleLower = If(entry.Title, "").ToLowerInvariant()
                        Dim summaryLower = If(entry.Summary, "").ToLowerInvariant()

                        For Each kw In keywords
                            If titleLower.Contains(kw) Then score += 3.0
                            If entry.Tags IsNot Nothing Then
                                For Each t In entry.Tags
                                    If t.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                        score += 2.0 : Exit For
                                    End If
                                Next
                            End If
                            If entry.Keywords IsNot Nothing Then
                                For Each k In entry.Keywords
                                    If k.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                        score += 1.5 : Exit For
                                    End If
                                Next
                            End If
                            If summaryLower.Contains(kw) Then score += 1.0
                        Next
                    End If

                    If score > 0 Then
                        ' Check for wiki summary page
                        Dim wikiPath As String = ""
                        If Not String.IsNullOrWhiteSpace(wikiDir) Then
                            Dim safeName = SanitizeFileName(If(entry.Title, Path.GetFileNameWithoutExtension(If(entry.FilePath, "unknown"))))
                            Dim candidate = Path.Combine(wikiDir, safeName & "-summary.md")
                            If File.Exists(candidate) Then wikiPath = candidate
                        End If

                        results.Add(New KnowledgeMatch With {
                            .Entry = entry,
                            .Score = score,
                            .StoreName = store.Name,
                            .WikiSummaryPath = wikiPath
                        })
                    End If
                Next
            Next

            ' Sort by descending score, take top N
            results.Sort(Function(a, b) b.Score.CompareTo(a.Score))
            If results.Count > maxResults Then
                results = results.GetRange(0, maxResults)
            End If

            Return results
        End Function


        ''' <summary>
        ''' Builds a knowledge context block from the given matches for LLM prompt injection.
        ''' Prefers wiki summary pages when available; falls back to raw document content.
        ''' </summary>
        Public Shared Function BuildKnowledgeContext(matches As List(Of KnowledgeMatch),
                                                     Optional maxTotalChars As Integer = 80000) As String
            If matches Is Nothing OrElse matches.Count = 0 Then Return ""

            Dim sb As New StringBuilder()
            sb.AppendLine("<KNOWLEDGESTORE>")
            sb.AppendLine("The following knowledge documents have been provided for context.")
            sb.AppendLine("IMPORTANT: Whenever you use or reference information from these documents, you MUST include a clickable source link in your answer using the precise markdown format: [Text](sourcePath).")
            sb.AppendLine()

            Dim totalChars As Integer = 0
            Dim docIndex As Integer = 1

            For Each m In matches
                If m.Entry Is Nothing Then Continue For

                ' Try wiki summary first (pre-compiled, richer)
                Dim content As String = ""
                If Not String.IsNullOrWhiteSpace(m.WikiSummaryPath) AndAlso File.Exists(m.WikiSummaryPath) Then
                    Try
                        content = File.ReadAllText(m.WikiSummaryPath, System.Text.Encoding.UTF8)
                    Catch
                    End Try
                End If

                ' Fall back to raw document content
                If String.IsNullOrWhiteSpace(content) Then
                    content = KnowledgeStoreManager.GetContent(m.Entry)
                End If

                If String.IsNullOrWhiteSpace(content) Then Continue For

                ' Budget check
                Dim remaining = maxTotalChars - totalChars
                If remaining <= 0 Then Exit For
                If content.Length > remaining Then
                    content = content.Substring(0, remaining) & vbCrLf & "[... truncated ...]"
                End If

                Dim title = If(m.Entry.Title, Path.GetFileNameWithoutExtension(If(m.Entry.FilePath, "Unknown")))
                Dim storeLabel = If(Not String.IsNullOrWhiteSpace(m.StoreName), m.StoreName, "")

                ' Resolve true clickable local file URI
                Dim safeURI = ""
                If Not String.IsNullOrWhiteSpace(m.Entry.FilePath) Then
                    safeURI = "file:///" & m.Entry.FilePath.Replace("\", "/")
                End If

                sb.AppendLine($"<KSDOCUMENT title=""{System.Security.SecurityElement.Escape(title)}"" store=""{System.Security.SecurityElement.Escape(storeLabel)}"" sourcePath=""{safeURI}"">")
                sb.AppendLine(content)
                sb.AppendLine("</KSDOCUMENT>")
                sb.AppendLine()

                totalChars += content.Length
                docIndex += 1
            Next

            sb.AppendLine("</KNOWLEDGESTORE>")
            Return sb.ToString()
        End Function


        ''' <summary>
        ''' Convenience method: resolves a query and builds the context block in one call.
        ''' </summary>
        Public Shared Function ResolveAndBuild(query As String, context As ISharedContext,
                                               Optional maxResults As Integer = 5,
                                               Optional maxTotalChars As Integer = 80000) As String
            Dim matches = ResolveQuery(query, context, maxResults)
            Return BuildKnowledgeContext(matches, maxTotalChars)
        End Function

        ''' <summary>Sanitizes a string for use as a filename.</summary>
        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(name.Where(Function(c) Not invalid.Contains(c)).ToArray())
            If cleaned.Length > 100 Then cleaned = cleaned.Substring(0, 100)
            Return If(String.IsNullOrWhiteSpace(cleaned), "unnamed", cleaned.Trim())
        End Function

    End Class

End Namespace