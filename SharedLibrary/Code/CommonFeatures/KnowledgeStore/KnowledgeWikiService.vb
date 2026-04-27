' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeWikiService.vb
' Purpose:
'   Central orchestration service for the Knowledge Store wiki layer.
'
' Responsibilities:
'   - Initialize and maintain the physical wiki structure under `.redink\Wiki`.
'   - Create, merge, normalize, and persist wiki pages with YAML front matter.
'   - Build schema-aware prompts for ingestion, supplemental pages, repair,
'     and health-analysis workflows.
'   - Run deterministic wiki maintenance such as section normalization, link
'     repair, review-queue refresh, and index rebuilding.
'   - Run gated LLM-assisted health analysis and structural repair, coordinated
'     through `KnowledgeStoreHostGate` so Knowledge Store AI work yields to
'     higher-priority host activity.
'   - Refresh embeddings and write operational diagnostics and repair logs.
'
' Key Artifacts:
'   - `index.md`
'   - `log.md`
'   - `health_report.md`
'   - `review_queue.md`
'
' Notes:
'   - This service is the main wiki-side engine of the Knowledge Store subsystem.
'   - It depends on `KnowledgeStoreCatalog`, `KnowledgeStoreSchema`,
'     `KnowledgeEmbeddingService`, `KnowledgeIndexer`, and `SharedMethods.LLM`.
'   - The implementation favors source-backed, durable pages and explicit
'     preservation of uncertainty, contradictions, and review signals.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext
Imports System.Runtime.ExceptionServices

Namespace SharedLibrary
    Public Class KnowledgeWikiService

        Public Const MaxChunkChars As Integer = 500000

        Public Const IngestionLlmThrottleForegroundMs As Integer = 0
        Public Const IngestionLlmThrottleBackgroundMs As Integer = 1500
        Public Const IngestionLlmEmptyRetryCount As Integer = 2
        Public Const IngestionLlmEmptyRetryDelayMs As Integer = 2000

        Private Class ParsedAgentResponse
            Public Property Title As String = "Untitled Concept"
            Public Property Summary As String = "Auto-generated wiki page."
            Public Property Kind As String = "Concept"
            Public Property Body As String = ""
        End Class

        Private Class WikiPageInfo
            Public Property FilePath As String = ""
            Public Property FileName As String = ""
            Public Property RelativePath As String = ""
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property Kind As String = "Concept"
            Public Property Status As String = "stable"
            Public Property ReviewNeeded As Boolean = False
            Public Property ContradictionCount As Integer = 0
            Public Property SourceCount As Integer = 0
            Public Property UpdatedUtc As DateTime? = Nothing
            Public Property Content As String = ""
            Public Property SourcePaths As New List(Of String)()
        End Class

        Private Class WikiComputedState
            Public Property Status As String = "stable"
            Public Property ReviewNeeded As Boolean = False
            Public Property ContradictionCount As Integer = 0
            Public Property SourceCount As Integer = 0
        End Class

        Private Class WikiClaimInfo
            Public Property ClaimText As String = ""
            Public Property EvidenceItems As New List(Of String)()
            Public Property SourcePaths As New List(Of String)()
            Public Property IsContradicted As Boolean = False
            Public Property NeedsReview As Boolean = False
        End Class

        Private Class PotentialContradictionPair
            Public Property PageA As WikiPageInfo = Nothing
            Public Property PageB As WikiPageInfo = Nothing
            Public Property ClaimA As String = ""
            Public Property ClaimB As String = ""
            Public Property Reason As String = ""
        End Class



        Private Class SupplementalWikiPageDraft
            Public Property Title As String = ""
            Public Property Summary As String = ""
            Public Property Kind As String = "Concept"
            Public Property Body As String = ""
        End Class

        Private Class WikiAutoFixResult
            Public Property UpdatedPages As Integer = 0
            Public Property AddedSourcesSections As Integer = 0
            Public Property AddedRelatedSections As Integer = 0
            Public Property RepairedLinks As Integer = 0
            Public Property RemovedBrokenLinks As Integer = 0
            Public Property LlmRepairedPages As Integer = 0
        End Class

#Region "Schema-driven helpers"

        Private Shared ReadOnly BuiltInPageKinds As String() = {"Source", "Entity", "Concept", "Analysis"}
        Private Shared ReadOnly DefaultAnalyticalSections As String() = {"## Key Claims", "## Evidence", "## Contradictions", "## Open Questions"}

        Private Shared Function GetEffectivePageKinds(schema As KnowledgeStoreSchema) As List(Of String)
            Dim list As New List(Of String)()
            If schema IsNot Nothing AndAlso schema.PreferredPageKinds IsNot Nothing Then
                For Each k In schema.PreferredPageKinds
                    Dim v = If(k, "").Trim()
                    If v.Length > 0 AndAlso Not list.Any(Function(x) x.Equals(v, StringComparison.OrdinalIgnoreCase)) Then
                        list.Add(v)
                    End If
                Next
            End If
            For Each b In BuiltInPageKinds
                If Not list.Any(Function(x) x.Equals(b, StringComparison.OrdinalIgnoreCase)) Then
                    list.Add(b)
                End If
            Next
            Return list
        End Function

        Private Shared Function GetEffectiveRequiredSections(schema As KnowledgeStoreSchema) As List(Of String)
            Dim list As New List(Of String)()
            Dim hasSchemaSections As Boolean = schema IsNot Nothing AndAlso
                                               schema.RequiredSections IsNot Nothing AndAlso
                                               schema.RequiredSections.Length > 0

            If hasSchemaSections Then
                For Each s In schema.RequiredSections
                    Dim v = If(s, "").Trim()
                    If v.Length = 0 Then Continue For
                    If v.Equals("Summary", StringComparison.OrdinalIgnoreCase) Then Continue For
                    If v.Equals("Title", StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim heading = If(v.StartsWith("#"), v, "## " & v)
                    If Not list.Any(Function(x) x.Equals(heading, StringComparison.OrdinalIgnoreCase)) Then
                        list.Add(heading)
                    End If
                Next
            Else
                list.AddRange(DefaultAnalyticalSections)
            End If

            ' Helper: did the user already cover this analytical role, either explicitly
            ' (via AnalyticalSections.<Role>) or implicitly (via a RequiredSections entry
            ' whose text matches one of the English search tokens)?
            Dim isAnalyticalRoleCovered As Func(Of String, String(), Boolean) =
                Function(explicitOverride As String, searchTerms As String()) As Boolean
                    Dim explicitClean = If(explicitOverride, "").Trim()
                    If explicitClean.Length > 0 Then
                        Dim explicitHeading = If(explicitClean.StartsWith("#"), explicitClean, "## " & explicitClean)
                        Return list.Any(Function(x) x.Equals(explicitHeading, StringComparison.OrdinalIgnoreCase))
                    End If
                    For Each existing In list
                        Dim clean = Regex.Replace(existing, "^\s*#+\s*", "").Trim().ToLowerInvariant()
                        For Each term In searchTerms
                            If clean.Contains(term) Then Return True
                        Next
                    Next
                    Return False
                End Function

            ' Resolve each analytical role against the schema overrides first, then the
            ' English fallback. Only append the English default if the role is genuinely
            ' missing from the user's section list AND no override is configured.
            Dim analyticalRoles As New List(Of (Override As String, SearchTerms As String(), DefaultHeading As String))()
            Dim ovr = If(schema, New KnowledgeStoreSchema()).AnalyticalSections

            analyticalRoles.Add((If(ovr, New KnowledgeStoreSchema.AnalyticalSectionsOptions()).Claims,
                                 New String() {"key claims", "key provisions", "key points", "claims", "provisions", "obligations", "requirements", "controls"},
                                 "## Key Claims"))
            analyticalRoles.Add((If(ovr, New KnowledgeStoreSchema.AnalyticalSectionsOptions()).Evidence,
                                 New String() {"evidence", "support", "supporting evidence", "basis", "justification", "authority", "supporting material"},
                                 "## Evidence"))
            analyticalRoles.Add((If(ovr, New KnowledgeStoreSchema.AnalyticalSectionsOptions()).OpenQuestions,
                                 New String() {"open questions", "questions", "gaps", "unknowns", "follow-up"},
                                 "## Open Questions"))
            If schema Is Nothing OrElse schema.DetectContradictions Then
                analyticalRoles.Add((If(ovr, New KnowledgeStoreSchema.AnalyticalSectionsOptions()).Contradictions,
                                     New String() {"contradictions", "conflicts", "disputes", "exceptions"},
                                     "## Contradictions"))
            End If

            For Each role In analyticalRoles
                If isAnalyticalRoleCovered(role.Override, role.SearchTerms) Then Continue For

                Dim toAppend As String
                Dim explicitClean = If(role.Override, "").Trim()
                If explicitClean.Length > 0 Then
                    toAppend = If(explicitClean.StartsWith("#"), explicitClean, "## " & explicitClean)
                Else
                    toAppend = role.DefaultHeading
                End If

                If Not list.Any(Function(x) x.Equals(toAppend, StringComparison.OrdinalIgnoreCase)) Then
                    list.Add(toAppend)
                End If
            Next

            ' Sources and Related Pages: honor the configured Labels override so the
            ' English literal is never appended alongside a localized variant.
            Dim sourcesHeading As String = "## Sources"
            Dim relatedHeading As String = "## Related Pages"
            If schema IsNot Nothing Then
                sourcesHeading = schema.GetLabel("SourcesHeading", sourcesHeading)
                relatedHeading = schema.GetLabel("RelatedPagesHeading", relatedHeading)
            End If

            If schema Is Nothing OrElse schema.AlwaysAddSourceLinks Then
                Dim hasSources = list.Any(Function(x) x.Equals(sourcesHeading, StringComparison.OrdinalIgnoreCase) OrElse
                                                       x.Equals("## Sources", StringComparison.OrdinalIgnoreCase))
                If Not hasSources Then list.Add(sourcesHeading)
            End If

            If schema Is Nothing OrElse schema.AlwaysCreateCrossLinks Then
                Dim hasRelated = list.Any(Function(x) x.Equals(relatedHeading, StringComparison.OrdinalIgnoreCase) OrElse
                                                       x.Equals("## Related Pages", StringComparison.OrdinalIgnoreCase))
                If Not hasRelated Then list.Add(relatedHeading)
            End If

            Return list
        End Function

        Private Shared Function PluralizeKindName(kind As String) As String
            Dim v = If(kind, "").Trim()
            If v.Length = 0 Then Return "Concepts"
            If v.Equals("Analysis", StringComparison.OrdinalIgnoreCase) Then Return "Analyses"
            If v.Equals("Entity", StringComparison.OrdinalIgnoreCase) Then Return "Entities"
            If v.Equals("Policy", StringComparison.OrdinalIgnoreCase) Then Return "Policies"
            If v.EndsWith("y", StringComparison.OrdinalIgnoreCase) AndAlso v.Length > 1 Then
                Dim beforeY = v.Chars(v.Length - 2)
                If "aeiouAEIOU".IndexOf(beforeY) < 0 Then Return v.Substring(0, v.Length - 1) & "ies"
            End If
            If v.EndsWith("s", StringComparison.OrdinalIgnoreCase) OrElse
               v.EndsWith("x", StringComparison.OrdinalIgnoreCase) OrElse
               v.EndsWith("sh", StringComparison.OrdinalIgnoreCase) OrElse
               v.EndsWith("ch", StringComparison.OrdinalIgnoreCase) Then
                Return v & "es"
            End If
            Return v & "s"
        End Function

        Private Shared Function FormatKindList(kinds As IEnumerable(Of String)) As String
            If kinds Is Nothing Then Return "Source|Entity|Concept|Analysis"
            Dim joined = String.Join("|", kinds.Where(Function(k) Not String.IsNullOrWhiteSpace(k)))
            If String.IsNullOrWhiteSpace(joined) Then Return "Source|Entity|Concept|Analysis"
            Return joined
        End Function

        Private Shared Function BuildSchemaSectionsGuidance(schema As KnowledgeStoreSchema) As String
            Dim sections = GetEffectiveRequiredSections(schema)
            Dim sb As New StringBuilder()
            sb.AppendLine("REQUIRED BODY SECTIONS (use these exact ## headings, in this order):")
            For Each s In sections
                sb.AppendLine("- " & s)
            Next
            Return sb.ToString().TrimEnd()
        End Function

        Private Shared Function ResolvePreferredAnalyticalHeading(schema As KnowledgeStoreSchema,
                                                                  defaultHeading As String,
                                                                  ParamArray searchTerms As String()) As String
            Dim sections = GetEffectiveRequiredSections(schema)

            For Each heading In sections
                Dim cleanHeading = Regex.Replace(heading, "^\s*#+\s*", "").Trim().ToLowerInvariant()

                For Each term In searchTerms
                    Dim cleanTerm = If(term, "").Trim().ToLowerInvariant()
                    If cleanTerm.Length > 0 AndAlso cleanHeading.Contains(cleanTerm) Then
                        Return heading
                    End If
                Next
            Next

            Return defaultHeading
        End Function

        Private Shared Function ApplyExplicitAnalyticalOverride(rawValue As String) As String
            Dim v = If(rawValue, "").Trim()
            If v.Length = 0 Then Return ""
            If v.StartsWith("#") Then Return v
            Return "## " & v
        End Function

        Private Shared Function GetClaimsHeading(schema As KnowledgeStoreSchema) As String
            If schema IsNot Nothing AndAlso schema.AnalyticalSections IsNot Nothing Then
                Dim explicitName = ApplyExplicitAnalyticalOverride(schema.AnalyticalSections.Claims)
                If explicitName.Length > 0 Then Return explicitName
            End If

            Return ResolvePreferredAnalyticalHeading(
                schema,
                "## Key Claims",
                "key claims", "key provisions", "key points",
                "claims", "provisions", "obligations", "requirements", "controls")
        End Function

        Private Shared Function GetEvidenceHeading(schema As KnowledgeStoreSchema) As String
            If schema IsNot Nothing AndAlso schema.AnalyticalSections IsNot Nothing Then
                Dim explicitName = ApplyExplicitAnalyticalOverride(schema.AnalyticalSections.Evidence)
                If explicitName.Length > 0 Then Return explicitName
            End If

            Return ResolvePreferredAnalyticalHeading(
                schema,
                "## Evidence",
                "evidence", "support", "supporting evidence",
                "basis", "justification", "authority", "supporting material")
        End Function

        Private Shared Function GetContradictionsHeading(schema As KnowledgeStoreSchema) As String
            If schema IsNot Nothing AndAlso schema.AnalyticalSections IsNot Nothing Then
                Dim explicitName = ApplyExplicitAnalyticalOverride(schema.AnalyticalSections.Contradictions)
                If explicitName.Length > 0 Then Return explicitName
            End If

            Return ResolvePreferredAnalyticalHeading(
                schema,
                "## Contradictions",
                "contradictions", "conflicts", "disputes", "exceptions")
        End Function

        Private Shared Function GetOpenQuestionsHeading(schema As KnowledgeStoreSchema) As String
            If schema IsNot Nothing AndAlso schema.AnalyticalSections IsNot Nothing Then
                Dim explicitName = ApplyExplicitAnalyticalOverride(schema.AnalyticalSections.OpenQuestions)
                If explicitName.Length > 0 Then Return explicitName
            End If

            Return ResolvePreferredAnalyticalHeading(
                schema,
                "## Open Questions",
                "open questions", "questions", "gaps", "unknowns", "follow-up")
        End Function

        Private Shared Function GetDefaultQueryPageKind(schema As KnowledgeStoreSchema) As String
            Dim kinds = GetEffectivePageKinds(schema)

            Dim analysisKind = kinds.FirstOrDefault(
                Function(k) k.Equals("Analysis", StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrWhiteSpace(analysisKind) Then
                Return analysisKind
            End If

            Dim firstNonSource = kinds.FirstOrDefault(
                Function(k) Not k.Equals("Source", StringComparison.OrdinalIgnoreCase))
            If Not String.IsNullOrWhiteSpace(firstNonSource) Then
                Return firstNonSource
            End If

            Return "Analysis"
        End Function

        Private Shared Function GetDefaultSectionBullet(heading As String,
                                                        schema As KnowledgeStoreSchema) As String
            If heading.Equals(GetClaimsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                Return "- No durable claims extracted yet."
            End If

            If heading.Equals(GetEvidenceHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                Return "- None recorded yet."
            End If

            If heading.Equals(GetContradictionsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                Return "- None noted."
            End If

            If heading.Equals(GetOpenQuestionsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                Return "- None noted."
            End If

            If heading.Equals("## Sources", StringComparison.OrdinalIgnoreCase) Then
                Return "- None recorded yet."
            End If

            If heading.Equals("## Related Pages", StringComparison.OrdinalIgnoreCase) Then
                Return "- None recorded yet."
            End If

            Return "- _Not yet provided._"
        End Function

#End Region

#Region "Embedding batching"

        ''' <summary>
        ''' Per async-flow embedding batch. While active, page writes only mark
        ''' files as "dirty"; embeddings are computed once per unique file when
        ''' the outermost batch closes. Re-entrant via Depth.
        ''' </summary>
        Private Class EmbeddingBatch
            Public Property KbRootPath As String = ""
            Public Property Depth As Integer = 0
            Public Property DirtyPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        End Class

        ' AsyncLocal so the batch flows naturally across Await boundaries
        ' without bleeding into unrelated tasks.
        Private Shared ReadOnly _currentBatch As New Threading.AsyncLocal(Of EmbeddingBatch)()

        ''' <summary>
        ''' Begins (or re-enters) an embedding batch scoped to <paramref name="kbRootPath"/>.
        ''' During an active batch, page rewrites are coalesced and a single embedding
        ''' update per unique file is performed when the outermost batch ends.
        ''' Always pair with <see cref="EndEmbeddingBatchAsync"/> in a Finally block.
        ''' </summary>
        Friend Shared Sub BeginEmbeddingBatch(kbRootPath As String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return

            Dim batch = _currentBatch.Value
            If batch Is Nothing Then
                batch = New EmbeddingBatch With {.KbRootPath = kbRootPath, .Depth = 1}
                _currentBatch.Value = batch
                Return
            End If

            ' Same store: nest. Different store: do not silently merge — leave the
            ' caller in immediate-write mode for the new path by ignoring nesting.
            If String.Equals(batch.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                batch.Depth += 1
            End If
        End Sub

        ''' <summary>
        ''' Closes a batch level. On the outermost exit, flushes one embedding
        ''' update per unique dirty page, reading the latest content from disk.
        ''' </summary>
        Friend Shared Async Function EndEmbeddingBatchAsync(context As ISharedContext) As Task
            Dim batch = _currentBatch.Value
            If batch Is Nothing Then Return

            batch.Depth -= 1
            If batch.Depth > 0 Then Return

            Dim kbRootPath = batch.KbRootPath
            Dim dirty = batch.DirtyPaths.ToArray()

            ' Clear before awaiting so re-entrant calls during the flush do not
            ' accidentally rejoin a half-closed batch.
            _currentBatch.Value = Nothing

            For Each filePath In dirty
                Try
                    If Not File.Exists(filePath) Then Continue For
                    Dim content = File.ReadAllText(filePath, Encoding.UTF8)
                    If String.IsNullOrWhiteSpace(content) Then Continue For

                    Await KnowledgeEmbeddingService.UpdateFileEmbeddingsAsync(
                        kbRootPath:=kbRootPath,
                        filePath:=filePath,
                        text:=content,
                        context:=context).ConfigureAwait(False)
                Catch ex As Exception
                    LogWikiError(kbRootPath, filePath, $"Deferred embedding update failed: {ex.Message}")
                End Try
            Next
        End Function

        ''' <summary>
        ''' If an embedding batch is active for <paramref name="kbRootPath"/>, queues
        ''' the file for a deferred single embedding update. Otherwise updates
        ''' immediately, preserving the legacy synchronous behavior.
        ''' </summary>
        Private Shared Async Function QueueOrUpdateEmbeddingAsync(kbRootPath As String,
                                                                  filePath As String,
                                                                  content As String,
                                                                  context As ISharedContext) As Task
            Dim batch = _currentBatch.Value
            If batch IsNot Nothing AndAlso
               String.Equals(batch.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                batch.DirtyPaths.Add(filePath)
                Return
            End If

            Await KnowledgeEmbeddingService.UpdateFileEmbeddingsAsync(
                kbRootPath:=kbRootPath,
                filePath:=filePath,
                text:=content,
                context:=context).ConfigureAwait(False)
        End Function

#End Region

#Region "Ingest scope (per-ingest snapshot cache & deferred maintenance)"

        ''' <summary>
        ''' Per async-flow ingest scope. While active:
        '''   - GetAllWikiPages returns a cached snapshot (invalidated on every save),
        '''     eliminating the repeated full-tree disk scan + front-matter parse.
        '''   - KnowledgeStoreSchema is loaded once and reused.
        '''   - RebuildIndex, RefreshReviewArtifacts, and the per-save related-link
        '''     backfill are suppressed and run exactly once when the outermost
        '''     scope ends, against the union of all saved pages.
        ''' Re-entrant via Depth, scoped to a single kbRootPath.
        ''' </summary>
        Private Class IngestScope
            Public Property KbRootPath As String = ""
            Public Property Depth As Integer = 0
            Public Property CachedPages As List(Of WikiPageInfo) = Nothing
            Public Property CachedSchema As KnowledgeStoreSchema = Nothing
            Public Property SavedPages As New List(Of WikiPageInfo)()
        End Class

        Private Shared ReadOnly _currentIngestScope As New Threading.AsyncLocal(Of IngestScope)()

        Private Shared Sub BeginIngestScope(kbRootPath As String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return

            Dim scope = _currentIngestScope.Value
            If scope Is Nothing Then
                _currentIngestScope.Value = New IngestScope With {.KbRootPath = kbRootPath, .Depth = 1}
                Return
            End If

            If String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                scope.Depth += 1
            End If
        End Sub

        Private Shared Async Function EndIngestScopeAsync(context As ISharedContext) As Task
            Dim scope = _currentIngestScope.Value
            If scope Is Nothing Then Return

            scope.Depth -= 1
            If scope.Depth > 0 Then Return

            Dim kbRootPath = scope.KbRootPath
            Dim savedPages = scope.SavedPages.ToList()
            Dim cachedSchema = scope.CachedSchema

            ' Clear before awaiting so re-entrant calls during the flush do not
            ' accidentally rejoin a half-closed scope.
            _currentIngestScope.Value = Nothing

            ' Deduplicate saved pages by file path; only the latest content matters.
            Dim uniqueSaves = savedPages.
                GroupBy(Function(p) p.FilePath, StringComparer.OrdinalIgnoreCase).
                Select(Function(g) g.Last()).
                ToList()

            For Each saved In uniqueSaves
                Try
                    Await BackfillRelatedLinksForExistingPagesAsync(
                        kbRootPath:=kbRootPath,
                        savedPage:=saved,
                        context:=context,
                        schema:=cachedSchema).ConfigureAwait(False)
                Catch ex As Exception
                    LogWikiError(kbRootPath, saved.FilePath, $"Deferred related backfill failed: {ex.Message}")
                End Try
            Next

            Try
                RebuildIndex(kbRootPath)
            Catch ex As Exception
                LogWikiError(kbRootPath, "RebuildIndex", $"Deferred index rebuild failed: {ex.Message}")
            End Try

            Try
                RefreshReviewArtifacts(kbRootPath)
            Catch ex As Exception
                LogWikiError(kbRootPath, "RefreshReviewArtifacts", $"Deferred review refresh failed: {ex.Message}")
            End Try
        End Function

        Private Shared Function IsIngestScopeActive(kbRootPath As String) As Boolean
            Dim scope = _currentIngestScope.Value
            Return scope IsNot Nothing AndAlso
                   String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function GetCachedSchema(kbRootPath As String) As KnowledgeStoreSchema
            Dim scope = _currentIngestScope.Value
            If scope IsNot Nothing AndAlso
               String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                If scope.CachedSchema Is Nothing Then
                    scope.CachedSchema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
                End If
                Return scope.CachedSchema
            End If
            Return KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
        End Function

        Private Shared Sub InvalidateIngestPageCache(kbRootPath As String)
            Dim scope = _currentIngestScope.Value
            If scope IsNot Nothing AndAlso
               String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                scope.CachedPages = Nothing
            End If
        End Sub

        Private Shared Sub TrackIngestSavedPage(kbRootPath As String, page As WikiPageInfo)
            Dim scope = _currentIngestScope.Value
            If scope Is Nothing Then Return
            If Not String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then Return
            scope.SavedPages.Add(page)
        End Sub

#End Region

        Public Shared Sub InitializeWikiStructure(kbRootPath As String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return

            Dim wikiRoot = GetWikiRootPath(kbRootPath)
            Dim rawPath = Path.Combine(kbRootPath, ".redink", "Raw")

            If Not Directory.Exists(wikiRoot) Then Directory.CreateDirectory(wikiRoot)
            If Not Directory.Exists(rawPath) Then Directory.CreateDirectory(rawPath)

            ' Ensure folder for every schema-preferred kind + built-ins.
            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            For Each kind In GetEffectivePageKinds(schema)
                Dim folderPath = Path.Combine(wikiRoot, GetKindFolderName(kind))
                If Not Directory.Exists(folderPath) Then
                    Directory.CreateDirectory(folderPath)
                End If
            Next

            Dim filesFolderName As String =
                    If(schema IsNot Nothing AndAlso schema.SourceAccess IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(schema.SourceAccess.FilesFolderName),
                       schema.SourceAccess.FilesFolderName, "_files")
            Dim sourcesFolderName As String =
                    If(schema IsNot Nothing AndAlso schema.SourceRegistry IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(schema.SourceRegistry.FolderName),
                       schema.SourceRegistry.FolderName, "Sources")

            Dim filesDir = Path.Combine(wikiRoot, filesFolderName)
            If Not Directory.Exists(filesDir) Then Directory.CreateDirectory(filesDir)
            Dim sourcesDir = Path.Combine(wikiRoot, sourcesFolderName)
            If Not Directory.Exists(sourcesDir) Then Directory.CreateDirectory(sourcesDir)

            Dim indexPath = Path.Combine(wikiRoot, KnowledgeStoreCatalog.IndexFile)
            If Not File.Exists(indexPath) Then
                File.WriteAllText(indexPath,
                                  "# Knowledge Base Index" & vbCrLf & vbCrLf &
                                  "> Auto-maintained catalog of wiki pages." & vbCrLf,
                                  Encoding.UTF8)
            End If

            Dim logPath = Path.Combine(wikiRoot, KnowledgeStoreCatalog.LogFile)
            If Not File.Exists(logPath) Then
                File.WriteAllText(logPath,
                                  "# Wiki Audit Log" & vbCrLf & vbCrLf &
                                  "> Chronological record of ingestions, updates, queries, and lint operations." & vbCrLf & vbCrLf,
                                  Encoding.UTF8)
            End If
            GenerateMkDocsConfig(kbRootPath)
        End Sub


        ''' <summary>
        ''' Creates or updates .redink/mkdocs.yml so that mkdocs can serve the wiki.
        ''' </summary>
        Private Shared Sub GenerateMkDocsConfig(kbRootPath As String)
            Try
                If String.IsNullOrWhiteSpace(kbRootPath) Then Return

                Dim redinkDir = Path.Combine(kbRootPath, ".redink")
                If Not Directory.Exists(redinkDir) Then Directory.CreateDirectory(redinkDir)

                Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)

                ' Resolve site name: StoreName > Domain > directory name > fallback.
                Dim siteName = ""
                If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.StoreName) Then
                    siteName = schema.StoreName.Trim()
                ElseIf schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.Domain) Then
                    siteName = schema.Domain.Trim() & " Wiki"
                End If

                If String.IsNullOrWhiteSpace(siteName) Then
                    Dim dirName = New DirectoryInfo(kbRootPath).Name
                    siteName = If(String.IsNullOrWhiteSpace(dirName), "Knowledge Store", dirName) & " Wiki"
                End If

                Dim sb As New StringBuilder()
                sb.AppendLine($"site_name: {siteName}")
                sb.AppendLine()
                sb.AppendLine($"docs_dir: {KnowledgeStoreCatalog.WikiFolder}")
                sb.AppendLine()
                sb.AppendLine("not_in_nav: |")
                sb.AppendLine("  /_files/")
                sb.AppendLine()
                sb.AppendLine("# Allow non-md assets to be served verbatim.")
                sb.AppendLine("# (mkdocs copies any non-md file under docs_dir into the build output.)")
                sb.AppendLine("theme:")
                sb.AppendLine("  name: material")
                sb.AppendLine()
                sb.AppendLine("plugins:")
                sb.AppendLine("  - search")
                sb.AppendLine("  - roamlinks")

                Dim mkdocsPath = Path.Combine(redinkDir, "mkdocs.yml")
                File.WriteAllText(mkdocsPath, sb.ToString().Trim() & vbLf, Encoding.UTF8)
            Catch
            End Try
        End Sub

        Private Shared Sub AppendToLog(kbRootPath As String, action As String, details As String)
            Dim logPath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, KnowledgeStoreCatalog.LogFile)
            If Not File.Exists(logPath) Then InitializeWikiStructure(kbRootPath)

            Dim datePrefix = DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            Dim safeDetails = If(details, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            Dim logEntry = $"## [{datePrefix}] {action} | {safeDetails}" & vbCrLf
            File.AppendAllText(logPath, logEntry, Encoding.UTF8)
        End Sub

        Friend Shared Sub AppendOperationalLog(kbRootPath As String, action As String, details As String)
            AppendToLog(kbRootPath, action, details)
        End Sub


        Private Shared Function GetWikiRootPath(kbRootPath As String) As String
            Return Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
        End Function

        Private Shared Function GetKindFolderName(kind As String) As String
            Dim normalized = If(kind, "").Trim()

            If normalized.Length = 0 Then
                normalized = "Concept"
            End If

            Select Case normalized
                Case "Source"
                    Return "Sources"
                Case "Entity"
                    Return "Entities"
                Case "Analysis"
                    Return "Analyses"
                Case "Concept"
                    Return "Concepts"
                Case Else
                    Return PluralizeKindName(normalized)
            End Select
        End Function

        Private Shared Function NormalizePageKind(kind As String, fallbackKind As String) As String
            Return NormalizePageKind(kind, fallbackKind, Nothing)
        End Function

        Private Shared Function NormalizePageKind(kind As String, fallbackKind As String, schema As KnowledgeStoreSchema) As String
            Dim value = If(kind, "").Trim()
            If value.Length = 0 Then Return If(String.IsNullOrWhiteSpace(fallbackKind), "Concept", fallbackKind)

            Dim effective = GetEffectivePageKinds(schema)
            For Each allowed In effective
                If value.Equals(allowed, StringComparison.OrdinalIgnoreCase) Then Return allowed
            Next

            Return If(String.IsNullOrWhiteSpace(fallbackKind), "Concept", fallbackKind)
        End Function

        Private Shared Function GetPageDirectoryForKind(kbRootPath As String, kind As String) As String
            Return Path.Combine(GetWikiRootPath(kbRootPath), GetKindFolderName(kind))
        End Function

        Private Shared Function GetRelativeWikiPath(kbRootPath As String, fullPath As String) As String
            Try
                Dim wikiRoot = GetWikiRootPath(kbRootPath)
                Dim rel = GetRelativePathCompat(wikiRoot, fullPath)
                Return NormalizeWikiRelativePath(rel)
            Catch
                Return Path.GetFileName(fullPath)
            End Try
        End Function

        Private Shared Function NormalizeWikiRelativePath(pathValue As String) As String
            Return If(pathValue, "").Replace("\"c, "/"c).Trim()
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
                Return targetPath
            End Try
        End Function


        Private Shared Function ResolveWikiLinkTarget(currentPagePath As String,
                                                      target As String,
                                                      kbRootPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(target) Then Return ""

                If target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                   target.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
                   target.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                    Return ""
                End If

                Dim cleanTarget = target.Replace("/"c, Path.DirectorySeparatorChar)

                If Path.IsPathRooted(cleanTarget) Then
                    Return NormalizeWikiRelativePath(GetRelativePathCompat(GetWikiRootPath(kbRootPath), cleanTarget))
                End If

                Dim currentDir = Path.GetDirectoryName(currentPagePath)
                If String.IsNullOrWhiteSpace(currentDir) Then Return ""

                Dim absolute = Path.GetFullPath(Path.Combine(currentDir, cleanTarget))
                Return NormalizeWikiRelativePath(GetRelativePathCompat(GetWikiRootPath(kbRootPath), absolute))
            Catch
                Return ""
            End Try
        End Function


        Public Shared Async Function CreatePageFromClipboardAsync(kbRootPath As String,
                                                                  clipboardText As String,
                                                                  context As ISharedContext) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(kbRootPath) OrElse String.IsNullOrWhiteSpace(clipboardText) Then Return False

            InitializeWikiStructure(kbRootPath)

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim wikiIndex = LoadIndexText(kbRootPath)
            Dim candidatePages = FindCandidateWikiPages(kbRootPath, clipboardText, 5)

            Dim systemPrompt = BuildAgentPrompt(
                roleDescription:="You are a Wiki Assistant. Turn the user's raw text into a clean, durable wiki page.",
                schema:=schema,
                wikiIndex:=wikiIndex,
                candidatePages:=candidatePages,
                defaultKind:="Analysis")

            Dim responseString As String = Await ExecuteKnowledgeStoreScopedAsync(
                context,
                Async Function(useAlternateAPI)
                    If clipboardText.Length <= MaxChunkChars Then
                        Return Await SharedMethods.LLM(
                            context:=context,
                            promptSystem:=systemPrompt,
                            promptUser:=clipboardText,
                            UseSecondAPI:=useAlternateAPI,
                            Hidesplash:=True)
                    End If

                    Return Await ProcessLargeTextInChunksAsync(context, systemPrompt, clipboardText, useAlternateAPI, "Clipboard_Text")
                End Function)

            If String.IsNullOrWhiteSpace(responseString) Then
                LogWikiError(kbRootPath, "Clipboard", "LLM returned empty or failed to process.")
                Return False
            End If

            Return Await ParseAndSaveAgentResponseAsync(
                kbRootPath:=kbRootPath,
                agentResponse:=responseString,
                context:=context,
                sourceFilePath:="",
                defaultKind:="Analysis",
                relatedCandidates:=candidatePages,
                actionName:="clipboard", schema:=schema).ConfigureAwait(False)
        End Function

        Public Shared Async Function IngestSourceAsync(kbRootPath As String,
                                                       sourceFilePath As String,
                                                       context As ISharedContext,
                                                       Optional isBackground As Boolean = False) As Task(Of Boolean)

            If String.IsNullOrWhiteSpace(sourceFilePath) OrElse Not File.Exists(sourceFilePath) Then Return False

            InitializeWikiStructure(kbRootPath)

            BeginEmbeddingBatch(kbRootPath)
            BeginIngestScope(kbRootPath)

            Dim succeeded As Boolean = False
            Dim captured As ExceptionDispatchInfo = Nothing

            Try
                succeeded = Await IngestSourceCoreAsync(
                    kbRootPath:=kbRootPath,
                    sourceFilePath:=sourceFilePath,
                    context:=context,
                    isBackground:=isBackground).ConfigureAwait(False)
            Catch ex As Exception
                captured = ExceptionDispatchInfo.Capture(ex)
            End Try

            ' End ingest scope BEFORE the embedding batch:
            ' the scope's flush triggers backfill writes that must still land
            ' in the same embedding batch (one embedding update per file).
            Await EndIngestScopeAsync(context).ConfigureAwait(False)
            Await EndEmbeddingBatchAsync(context).ConfigureAwait(False)

            If captured IsNot Nothing Then captured.Throw()
            Return succeeded
        End Function

        Private Shared Async Function IngestSourceCoreAsync(kbRootPath As String,
                                                            sourceFilePath As String,
                                                            context As ISharedContext,
                                                            isBackground As Boolean) As Task(Of Boolean)

            Dim sourceText As String = ""
            Try
                sourceText = Await KnowledgeIndexer.ReadSourceTextAsync(sourceFilePath, context).ConfigureAwait(False)
            Catch ex As Exception
                LogWikiError(kbRootPath, sourceFilePath, $"File could not be read: {ex.Message}")
                Return False
            End Try

            If String.IsNullOrWhiteSpace(sourceText) Then
                LogWikiError(kbRootPath, sourceFilePath, "File produced empty extracted text.")
                Return False
            End If

            ' ── Re-ingest guard ──────────────────────────────────────────────
            ' If this exact source has already been ingested (i.e. at least one
            ' wiki page lists it under source_paths) AND the extracted text
            ' hash hasn't changed since last time, skip the entire LLM ingest.
            ' This is what stops background re-scans (mtime drift, AV touches,
            ' resync) from minting fresh supplemental pages every run.
            Dim currentHash As String = ComputeSourceTextHash(sourceText)
            Dim hashFilePath As String = GetSourceHashCachePath(kbRootPath, sourceFilePath)

            Dim alreadyIngested As Boolean =
                GetAllWikiPages(kbRootPath).
                    Any(Function(p) p.SourcePaths IsNot Nothing AndAlso
                                    p.SourcePaths.Any(Function(sp) String.Equals(sp, sourceFilePath, StringComparison.OrdinalIgnoreCase)))

            Dim contentUnchanged As Boolean = False
            Try
                If File.Exists(hashFilePath) Then
                    Dim previousHash = File.ReadAllText(hashFilePath, Encoding.UTF8).Trim()
                    contentUnchanged = String.Equals(previousHash, currentHash, StringComparison.OrdinalIgnoreCase)
                End If
            Catch
                contentUnchanged = False
            End Try

            If alreadyIngested AndAlso contentUnchanged Then
                AppendToLog(kbRootPath, "ingest-skip", $"{Path.GetFileName(sourceFilePath)} | unchanged content")
                Return True
            End If
            ' ─────────────────────────────────────────────────────────────────

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim wikiIndex = LoadIndexText(kbRootPath)
            Dim candidatePages = FindCandidateWikiPages(kbRootPath, Path.GetFileNameWithoutExtension(sourceFilePath) & " " & sourceText, 8)

            Dim systemPrompt = BuildAgentPrompt(
                roleDescription:="You are an autonomous Wiki Maintainer. Read the source and integrate it into the persistent wiki.",
                schema:=schema,
                wikiIndex:=wikiIndex,
                candidatePages:=candidatePages,
                defaultKind:="Source")

            Dim promptUser = $"Source File: {Path.GetFileName(sourceFilePath)}{vbCrLf}{vbCrLf}{sourceText}"

            Dim responseString As String

            If sourceText.Length <= MaxChunkChars Then
                responseString = Await ExecuteKnowledgeStoreLlmAsync(
                    context:=context,
                    promptSystem:=systemPrompt,
                    promptUser:=promptUser,
                    hideSplash:=True,
                    isBackground:=isBackground).ConfigureAwait(False)
            Else
                responseString = Await ExecuteKnowledgeStoreScopedAsync(
                    context,
                    Async Function(useAlternateAPI)
                        Return Await ProcessLargeTextInChunksAsync(
                            context:=context,
                            systemPrompt:=systemPrompt,
                            fullText:=sourceText,
                            useAlternateAPI:=useAlternateAPI,
                            fileName:=Path.GetFileName(sourceFilePath),
                            isBackground:=isBackground).ConfigureAwait(False)
                    End Function).ConfigureAwait(False)
            End If

            If String.IsNullOrWhiteSpace(responseString) Then
                LogWikiError(kbRootPath, sourceFilePath, "LLM returned empty or failed to process.")
                Return False
            End If

            Dim primarySaved = Await ParseAndSaveAgentResponseAsync(
                kbRootPath:=kbRootPath,
                agentResponse:=responseString,
                context:=context,
                sourceFilePath:=sourceFilePath,
                defaultKind:="Source",
                relatedCandidates:=candidatePages,
                actionName:="ingest", schema:=schema).ConfigureAwait(False)

            If Not primarySaved Then
                Return False
            End If

            ' Only run the expensive supplemental fan-out the FIRST time we see
            ' this source. On any later re-ingest the primary page is merged in
            ' place above and we skip the 8-15 extra LLM-generated pages that
            ' otherwise accumulate every background run.
            If Not alreadyIngested Then
                Try
                    Await GenerateAndApplySupplementalPagesAsync(
                        kbRootPath:=kbRootPath,
                        sourceFilePath:=sourceFilePath,
                        sourceText:=sourceText,
                        context:=context,
                        schema:=schema,
                        candidatePages:=candidatePages,
                        isBackground:=isBackground).ConfigureAwait(False)
                Catch ex As Exception
                    LogWikiError(kbRootPath, sourceFilePath, $"Supplemental page generation failed: {ex.Message}")
                End Try
            End If

            ' Record the content hash so the next run can short-circuit cleanly.
            Try
                Dim hashDir = Path.GetDirectoryName(hashFilePath)
                If Not String.IsNullOrWhiteSpace(hashDir) AndAlso Not Directory.Exists(hashDir) Then
                    Directory.CreateDirectory(hashDir)
                End If
                File.WriteAllText(hashFilePath, currentHash, Encoding.UTF8)
            Catch ex As Exception
                LogWikiError(kbRootPath, sourceFilePath, $"Hash cache write failed: {ex.Message}")
            End Try

            Return True
        End Function

        Private Shared Function ComputeSourceTextHash(text As String) As String
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(If(text, "")))
                Return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Function GetSourceHashCachePath(kbRootPath As String, sourceFilePath As String) As String
            Dim slug As String = SanitizeFileName(Path.GetFileNameWithoutExtension(If(sourceFilePath, "")))
            If String.IsNullOrWhiteSpace(slug) Then slug = "source"

            ' Disambiguate same-name files in different folders by appending a
            ' short hash of the full path. Keeps the cache filename stable
            ' across runs while avoiding collisions.
            Dim pathKey As String = ""
            Try
                Using sha = System.Security.Cryptography.SHA256.Create()
                    Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(If(sourceFilePath, "").ToLowerInvariant()))
                    pathKey = BitConverter.ToString(bytes, 0, 6).Replace("-", "").ToLowerInvariant()
                End Using
            Catch
                pathKey = "x"
            End Try

            Dim dir = Path.Combine(kbRootPath, ".redink", "ingest-cache")
            Return Path.Combine(dir, $"{slug}.{pathKey}.sha256")
        End Function

        Private Shared Async Function GenerateAndApplySupplementalPagesAsync(kbRootPath As String,
                                                                             sourceFilePath As String,
                                                                             sourceText As String,
                                                                             context As ISharedContext,
                                                                             schema As KnowledgeStoreSchema,
                                                                             candidatePages As List(Of WikiPageInfo),
                                                                             Optional isBackground As Boolean = False) As Task(Of Integer)
            If String.IsNullOrWhiteSpace(kbRootPath) OrElse
               String.IsNullOrWhiteSpace(sourceFilePath) OrElse
               String.IsNullOrWhiteSpace(sourceText) Then
                Return 0
            End If

            ' Schema-driven hard cap. If supplemental generation is disabled for
            ' this store, skip the LLM round-trip entirely.
            Dim maxSupplemental As Integer = 15
            If schema IsNot Nothing Then
                maxSupplemental = Math.Max(0, schema.MaxSupplementalPagesPerSource)
            End If

            If maxSupplemental <= 0 Then
                Return 0
            End If

            Dim wikiIndex = LoadIndexText(kbRootPath)
            Dim promptSystem = BuildSupplementalPagesPrompt(schema, wikiIndex, candidatePages)
            Dim promptUser As New StringBuilder()

            promptUser.AppendLine($"Source File: {Path.GetFileName(sourceFilePath)}")
            promptUser.AppendLine()
            promptUser.AppendLine("SOURCE CONTENT:")
            promptUser.AppendLine(sourceText)

            Dim response As String

            If sourceText.Length <= MaxChunkChars Then
                response = Await ExecuteKnowledgeStoreLlmAsync(
                    context:=context,
                    promptSystem:=promptSystem,
                    promptUser:=promptUser.ToString(),
                    hideSplash:=True,
                    isBackground:=isBackground).ConfigureAwait(False)
            Else
                response = Await ExecuteKnowledgeStoreScopedAsync(
                    context,
                    Async Function(useAlternateAPI)
                        Return Await ProcessLargeTextInChunksAsync(
                            context:=context,
                            systemPrompt:=promptSystem,
                            fullText:=sourceText,
                            useAlternateAPI:=useAlternateAPI,
                            fileName:=Path.GetFileName(sourceFilePath),
                            isBackground:=isBackground).ConfigureAwait(False)
                    End Function).ConfigureAwait(False)
            End If

            Dim drafts = ParseSupplementalPageDrafts(response)
            If drafts.Count = 0 Then Return 0

            ' Defensive cap: even if the model ignored the prompt's upper bound,
            ' never apply more than the configured maximum to the wiki.
            If drafts.Count > maxSupplemental Then
                drafts = drafts.Take(maxSupplemental).ToList()
            End If

            Dim savedCount As Integer = 0
            Dim seenTitles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each draft In drafts
                If String.IsNullOrWhiteSpace(draft.Title) OrElse String.IsNullOrWhiteSpace(draft.Body) Then
                    Continue For
                End If

                If seenTitles.Contains(draft.Title.Trim()) Then
                    Continue For
                End If
                seenTitles.Add(draft.Title.Trim())

                Dim syntheticAgentResponse = ComposeSyntheticAgentResponse(
                            title:=draft.Title,
                            summary:=draft.Summary,
                            kind:=draft.Kind,
                            body:=draft.Body,
                            schema:=schema)

                Dim saved = Await ParseAndSaveAgentResponseAsync(
                    kbRootPath:=kbRootPath,
                    agentResponse:=syntheticAgentResponse,
                    context:=context,
                    sourceFilePath:=sourceFilePath,
                    defaultKind:=NormalizePageKind(draft.Kind, "Concept", schema),
                    relatedCandidates:=candidatePages,
                    actionName:="supplemental-update", schema:=schema).ConfigureAwait(False)

                If saved Then
                    savedCount += 1
                End If
            Next

            If savedCount > 0 Then
                AppendToLog(kbRootPath, "supplemental", $"{Path.GetFileName(sourceFilePath)} | Pages={savedCount}; Cap={maxSupplemental}")
            End If

            Return savedCount
        End Function

        Private Shared Function BuildSupplementalPagesPrompt(schema As KnowledgeStoreSchema,
                                                             wikiIndex As String,
                                                             candidatePages As List(Of WikiPageInfo)) As String
            Dim sb As New StringBuilder()
            Dim kindsList = FormatKindList(GetEffectivePageKinds(schema).
                                           Where(Function(k) Not k.Equals("Source", StringComparison.OrdinalIgnoreCase) AndAlso
                                                             Not k.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase)))

            If schema IsNot Nothing Then sb.Append(schema.ResolveLanguageDirective())
            sb.AppendLine("You are an autonomous Wiki Maintainer.")
            sb.AppendLine("Your job is to identify additional durable wiki pages that should be created or updated based on the source.")
            sb.AppendLine("Do NOT restate the main source page. Only produce pages that add long-term structure to the wiki.")
            sb.AppendLine()
            sb.AppendLine("OUTPUT FORMAT:")
            sb.AppendLine("Return zero or more PAGE blocks in the following exact format:")
            sb.AppendLine("<PAGE>")
            sb.AppendLine("TITLE: ...")
            sb.AppendLine("SUMMARY: ...")
            sb.AppendLine($"KIND: {kindsList}")
            sb.AppendLine()
            sb.AppendLine("<markdown body>")
            sb.AppendLine("</PAGE>")
            sb.AppendLine()
            sb.AppendLine("If no supplemental pages are justified, return exactly: NONE")
            sb.AppendLine()
            sb.AppendLine("Rules:")
            sb.AppendLine("1. Prefer durable pages over transient notes.")
            sb.AppendLine("2. Add meaningful markdown cross-links to existing pages.")
            sb.AppendLine("3. Preserve contradictions and uncertainty.")
            sb.AppendLine("4. Do not invent sources.")
            sb.AppendLine("5. Only create pages that are likely to remain useful after this source.")
            sb.AppendLine("6. If the source conflicts with an existing page, update the relevant page rather than hiding the conflict.")
            sb.AppendLine("7. Only create clickable markdown page links to pages that already exist in the supplied wiki context.")
            sb.AppendLine("8. If a useful concept does not yet have a page, mention it as plain text or under ## Open Questions, but do not create a broken page link.")
            sb.AppendLine()
            Dim entityTypesList = ""
            If schema IsNot Nothing AndAlso schema.EntityTypes IsNot Nothing AndAlso schema.EntityTypes.Length > 0 Then
                entityTypesList = String.Join(", ", schema.EntityTypes.Where(Function(t) Not String.IsNullOrWhiteSpace(t)))
            End If

            Dim minSupplemental As Integer = 8
            Dim maxSupplemental As Integer = 15
            If schema IsNot Nothing Then
                minSupplemental = Math.Max(0, schema.MinSupplementalPagesPerSource)
                maxSupplemental = Math.Max(0, schema.MaxSupplementalPagesPerSource)
                If maxSupplemental > 0 AndAlso minSupplemental > maxSupplemental Then
                    minSupplemental = maxSupplemental
                End If
            End If

            sb.AppendLine("PAGE-KIND DIVERSITY:")
            sb.AppendLine($"- The available page kinds for this store are: {kindsList.Replace("|", ", ")}.")
            sb.AppendLine("- Strongly prefer a MIX of these kinds across the supplemental batch. Do not default every supplemental page to the same one or two kinds.")
            sb.AppendLine("- For each candidate page, pick the MOST SPECIFIC kind that fits its subject matter from the list above.")
            sb.AppendLine("- Whenever the source mentions a recurring concept, role, named organization, system, jurisdiction, defined term, broader theme, or interpretive analysis that does not yet have a wiki page, that is a strong candidate for its own page under the most fitting kind.")
            sb.AppendLine("- Do not use document-style kinds (such as ones representing codified directives, policies, regulations, or standalone source documents) for general concepts, roles, definitions, or themes; reserve those for pages that themselves represent a separate codified document.")
            If Not String.IsNullOrWhiteSpace(entityTypesList) Then
                sb.AppendLine($"- The store recognizes the following entity types: {entityTypesList}. Treat each as a candidate subject for its own page when applicable.")
            End If

            ' Schema-configurable fan-out guidance. The model is told the real
            ' min/max so we don't have to lean on a hard-coded "8-15" any more.
            If maxSupplemental <= 0 Then
                sb.AppendLine("- Supplemental page generation is disabled for this store; return exactly: NONE")
            ElseIf maxSupplemental = 1 Then
                sb.AppendLine("- Return AT MOST 1 supplemental page, and only if it is clearly durable. Otherwise return: NONE")
            Else
                Dim softFloor As Integer = Math.Max(1, Math.Min(minSupplemental, maxSupplemental))
                sb.AppendLine($"- Aim for roughly {softFloor}-{maxSupplemental} supplemental pages per non-trivial source, distributed across several kinds. Returning fewer is acceptable only when the source is genuinely thin. Do NOT exceed {maxSupplemental}.")
                sb.AppendLine("- Before deciding what to return, internally enumerate every distinct concept, role, organization, system, jurisdiction, defined term, broader theme, and entity-type instance mentioned or strongly implied by the source. Then create a page for each one that meets the durability bar.")
                sb.AppendLine("- Coverage check: if the source mentions an entity-type instance for which no wiki page exists yet, you SHOULD create one unless it is clearly trivial.")
                sb.AppendLine("- It is normal and desirable to return many pages of varied kinds per source. Do not stop at one or two.")
            End If

            sb.AppendLine("SOURCE CITATION RULES:")
            sb.AppendLine("- Refer to sources only by ID markers like [S1], [S2] embedded in Key Claims and Evidence bullets.")
            sb.AppendLine("- DO NOT author the ## Sources section; it is generated deterministically.")
            sb.AppendLine("- DO NOT write file paths, file:// URLs, or download links anywhere.")
            If schema?.PerClaimCitations IsNot Nothing AndAlso schema.PerClaimCitations.Enabled Then
                sb.AppendLine("- Every Key Claim and every Evidence bullet MUST end with at least one [S#] marker referring to the source it relies on.")
            End If
            sb.AppendLine()
            sb.AppendLine(BuildSchemaSectionsGuidance(schema))
            sb.AppendLine()

            If schema IsNot Nothing Then
                sb.AppendLine(schema.ToPromptBlock())
                sb.AppendLine()
            End If

            If Not String.IsNullOrWhiteSpace(wikiIndex) Then
                sb.AppendLine("CURRENT WIKI INDEX:")
                sb.AppendLine(wikiIndex)
                sb.AppendLine()
            End If

            If candidatePages IsNot Nothing AndAlso candidatePages.Count > 0 Then
                sb.AppendLine("MOST RELEVANT EXISTING PAGES:")
                For Each candidate In candidatePages
                    sb.AppendLine($"- Title: {candidate.Title}")
                    sb.AppendLine($"  Path: {candidate.FileName}")
                    sb.AppendLine($"  Kind: {candidate.Kind}")
                    sb.AppendLine($"  Status: {candidate.Status}")
                    If Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        sb.AppendLine($"  Summary: {candidate.Summary}")
                    End If
                Next
                sb.AppendLine()
            End If

            Return sb.ToString().Trim()
        End Function

        Public Shared Async Function FileQueryResultAsync(kbRootPath As String,
                                                          queryText As String,
                                                          answerMarkdown As String,
                                                          context As ISharedContext,
                                                          Optional preferredTitle As String = "",
                                                          Optional sourcePaths As IEnumerable(Of String) = Nothing) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(kbRootPath) OrElse
               String.IsNullOrWhiteSpace(queryText) OrElse
               String.IsNullOrWhiteSpace(answerMarkdown) Then
                Return False
            End If

            InitializeWikiStructure(kbRootPath)

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            If schema Is Nothing OrElse Not schema.QueryFilingEnabled Then
                Return False
            End If

            Dim sourceList As New List(Of String)()
            If sourcePaths IsNot Nothing Then
                sourceList.AddRange(
                    sourcePaths.
                        Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                        Select(Function(p) p.Trim()).
                        Distinct(StringComparer.OrdinalIgnoreCase))
            End If

            Dim candidatePages = FindCandidateWikiPages(kbRootPath, queryText & " " & answerMarkdown, 8)
            Dim wikiIndex = LoadIndexText(kbRootPath)
            Dim queryKind = GetDefaultQueryPageKind(schema)
            Dim kindsList = FormatKindList(GetEffectivePageKinds(schema))

            Dim systemPrompt As New StringBuilder()
            If schema IsNot Nothing Then systemPrompt.Append(schema.ResolveLanguageDirective())
            systemPrompt.AppendLine("You are an autonomous Wiki Maintainer.")
            systemPrompt.AppendLine("Convert the provided query and answer into a durable markdown wiki page.")
            systemPrompt.AppendLine("Preserve the substance of the answer, keep uncertainty where appropriate, add meaningful cross-links to existing wiki pages, and retain clickable markdown source links.")
            systemPrompt.AppendLine("This page represents a reusable analysis derived from a user query.")
            systemPrompt.AppendLine()
            systemPrompt.AppendLine("OUTPUT FORMAT REQUIREMENTS:")
            systemPrompt.AppendLine("1. First line: TITLE: <concise title>")
            systemPrompt.AppendLine("2. Second line: SUMMARY: <one-sentence summary>")
            systemPrompt.AppendLine($"3. Third line: KIND: <{kindsList}>")
            systemPrompt.AppendLine("4. Start the markdown body after a blank line.")
            systemPrompt.AppendLine("5. Use standard markdown links, not wiki syntax.")
            systemPrompt.AppendLine("6. Preserve source citations and keep them clickable.")
            systemPrompt.AppendLine("7. Only create clickable markdown page links to pages that already exist in the supplied wiki context.")
            systemPrompt.AppendLine("8. If a useful concept does not yet have a page, mention it as plain text or under ## Open Questions, but do not create a broken page link.")
            systemPrompt.AppendLine("9. Do not invent sources.")
            systemPrompt.AppendLine("SOURCE CITATION RULES:")
            systemPrompt.AppendLine("- Refer to sources only by ID markers like [S1], [S2] embedded in Key Claims and Evidence bullets.")
            systemPrompt.AppendLine("- DO NOT author the ## Sources section; it is generated deterministically.")
            systemPrompt.AppendLine("- DO NOT write file paths, file:// URLs, or download links anywhere.")
            If schema?.PerClaimCitations IsNot Nothing AndAlso schema.PerClaimCitations.Enabled Then
                systemPrompt.AppendLine("- Every Key Claim and every Evidence bullet MUST end with at least one [S#] marker referring to the source it relies on.")
            End If
            systemPrompt.AppendLine()
            systemPrompt.AppendLine(BuildSchemaSectionsGuidance(schema))
            systemPrompt.AppendLine()

            If schema IsNot Nothing AndAlso schema.DetectContradictions Then
                systemPrompt.AppendLine($"If the answer conflicts with existing pages, record the disagreement explicitly under {GetContradictionsHeading(schema)}.")
                systemPrompt.AppendLine()
            End If

            If schema IsNot Nothing Then
                systemPrompt.AppendLine(schema.ToPromptBlock())
            End If

            If Not String.IsNullOrWhiteSpace(wikiIndex) Then
                systemPrompt.AppendLine()
                systemPrompt.AppendLine("CURRENT WIKI INDEX:")
                systemPrompt.AppendLine(wikiIndex)
            End If

            If candidatePages IsNot Nothing AndAlso candidatePages.Count > 0 Then
                systemPrompt.AppendLine()
                systemPrompt.AppendLine("MOST RELEVANT EXISTING PAGES:")
                For Each candidate In candidatePages
                    systemPrompt.AppendLine($"- Title: {candidate.Title}")
                    systemPrompt.AppendLine($"  Path: {candidate.FileName}")
                    systemPrompt.AppendLine($"  Kind: {candidate.Kind}")
                    systemPrompt.AppendLine($"  Status: {candidate.Status}")
                    If Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        systemPrompt.AppendLine($"  Summary: {candidate.Summary}")
                    End If
                Next
            End If

            systemPrompt.AppendLine()
            systemPrompt.AppendLine($"Default page kind if unclear: {queryKind}")

            Dim userPrompt As New StringBuilder()
            userPrompt.AppendLine($"QUERY: {queryText}")
            userPrompt.AppendLine()
            userPrompt.AppendLine("ANSWER:")
            userPrompt.AppendLine(answerMarkdown.Trim())

            If sourceList.Count > 0 Then
                userPrompt.AppendLine()
                userPrompt.AppendLine("SOURCE PATHS:")
                For Each path In sourceList
                    userPrompt.AppendLine($"- {path}")
                Next
            End If

            Dim agentResponse = Await ExecuteKnowledgeStoreLlmAsync(
                context:=context,
                promptSystem:=systemPrompt.ToString().Trim(),
                promptUser:=userPrompt.ToString().Trim(),
                hideSplash:=True).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(agentResponse) Then
                agentResponse = BuildFallbackQueryAgentResponse(queryText, answerMarkdown, preferredTitle, schema)
            ElseIf Not String.IsNullOrWhiteSpace(preferredTitle) AndAlso
                   agentResponse.IndexOf("TITLE:", StringComparison.OrdinalIgnoreCase) >= 0 Then
                agentResponse = OverrideAgentTitle(agentResponse, preferredTitle)
            End If

            Dim saved = Await ParseAndSaveAgentResponseAsync(
                kbRootPath:=kbRootPath,
                agentResponse:=agentResponse,
                context:=context,
                sourceFilePath:="",
                defaultKind:=queryKind,
                relatedCandidates:=candidatePages,
                actionName:="query-file",
                additionalSourcePaths:=sourceList,
                schema:=schema).ConfigureAwait(False)

            If saved Then
                AppendToLog(kbRootPath, "query-file", queryText)
            End If

            Return saved
        End Function

        Private Shared Function BuildFallbackQueryAgentResponse(queryText As String,
                                                                answerMarkdown As String,
                                                                preferredTitle As String,
                                                                Optional schema As KnowledgeStoreSchema = Nothing) As String
            Dim title = preferredTitle.Trim()
            If String.IsNullOrWhiteSpace(title) Then
                title = BuildQueryTitle(queryText)
            End If

            Dim summary = BuildFallbackSummary(StripMarkdown(answerMarkdown))
            If String.IsNullOrWhiteSpace(summary) Then
                summary = "Filed analysis generated from a knowledge store query."
            End If

            Dim queryKind = GetDefaultQueryPageKind(schema)
            Dim sections = GetEffectiveRequiredSections(schema)

            Dim sb As New StringBuilder()
            sb.AppendLine($"TITLE: {title}")
            sb.AppendLine($"SUMMARY: {summary}")
            sb.AppendLine($"KIND: {queryKind}")
            sb.AppendLine()

            For Each heading In sections
                If heading.Equals("## Sources", StringComparison.OrdinalIgnoreCase) OrElse
                   heading.Equals("## Related Pages", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                sb.AppendLine(heading)
                sb.AppendLine()

                If heading.Equals(GetClaimsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("- Derived from the filed query and answer below.")
                ElseIf heading.Equals(GetEvidenceHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("- Evidence currently comes from the answer text and linked source paths.")
                ElseIf heading.Equals(GetContradictionsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("- None noted.")
                ElseIf heading.Equals(GetOpenQuestionsHeading(schema), StringComparison.OrdinalIgnoreCase) Then
                    sb.AppendLine("- Should this analysis be merged into an existing page?")
                Else
                    sb.AppendLine(GetDefaultSectionBullet(heading, schema))
                End If

                sb.AppendLine()
            Next

            sb.AppendLine("## Query")
            sb.AppendLine()
            sb.AppendLine(queryText.Trim())
            sb.AppendLine()
            sb.AppendLine("## Answer")
            sb.AppendLine()
            sb.AppendLine(answerMarkdown.Trim())

            Return sb.ToString().Trim()
        End Function

        Private Shared Function OverrideAgentTitle(agentResponse As String, preferredTitle As String) As String
            Dim lines = agentResponse.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList()

            For i As Integer = 0 To lines.Count - 1
                If lines(i).Trim().StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                    lines(i) = "TITLE: " & preferredTitle.Trim()
                    Return String.Join(vbCrLf, lines)
                End If
            Next

            lines.Insert(0, "TITLE: " & preferredTitle.Trim())
            Return String.Join(vbCrLf, lines)
        End Function

        Private Shared Function BuildQueryTitle(queryText As String) As String
            Dim normalized = Regex.Replace(If(queryText, "").Trim(), "\s+", " ")
            If normalized.Length > 80 Then
                normalized = normalized.Substring(0, 80).Trim()
            End If

            If String.IsNullOrWhiteSpace(normalized) Then
                normalized = "Query Analysis"
            End If

            Return "Query - " & normalized
        End Function

        Private Shared Function StripMarkdown(markdown As String) As String
            Dim text = If(markdown, "")

            text = Regex.Replace(text, "\[([^\]]+)\]\([^)]+\)", "$1")
            text = Regex.Replace(text, "[#>*_`~-]+", " ")
            text = Regex.Replace(text, "\s+", " ")

            Return text.Trim()
        End Function


        Private Shared Async Function ExecuteKnowledgeStoreLlmWithRetryAsync(context As ISharedContext,
                                                                             promptSystem As String,
                                                                             promptUser As String,
                                                                             useAlternateAPI As Boolean,
                                                                             hideSplash As Boolean,
                                                                             isBackground As Boolean,
                                                                             Optional progressCallback As Action(Of String) = Nothing,
                                                                             Optional operationName As String = "Knowledge Store AI step") As Task(Of String)
            Dim throttleMs As Integer = If(isBackground,
                                           IngestionLlmThrottleBackgroundMs,
                                           IngestionLlmThrottleForegroundMs)

            Dim maxAttempts As Integer = Math.Max(1, IngestionLlmEmptyRetryCount + 1)
            Dim effectiveOperationName = If(String.IsNullOrWhiteSpace(operationName), "Knowledge Store AI step", operationName.Trim())

            If throttleMs > 0 Then
                ReportMaintenanceProgress(progressCallback, $"{effectiveOperationName}: pacing request.")
                Await Task.Delay(throttleMs).ConfigureAwait(False)
            End If

            For attempt As Integer = 1 To maxAttempts
                If maxAttempts > 1 Then
                    ReportMaintenanceProgress(progressCallback, $"{effectiveOperationName}: AI request attempt {attempt} of {maxAttempts}.")
                End If

                Dim heartbeatCts As Threading.CancellationTokenSource = Nothing
                Dim heartbeatTask As Task = Nothing
                Dim callStartedUtc = DateTime.UtcNow
                Dim response As String = ""

                Try
                    heartbeatCts = New Threading.CancellationTokenSource()

                    heartbeatTask = Task.Run(
                        Async Function()
                            While Not heartbeatCts.Token.IsCancellationRequested
                                Await Task.Delay(5000, heartbeatCts.Token).ConfigureAwait(False)

                                If heartbeatCts.Token.IsCancellationRequested Then
                                    Exit While
                                End If

                                Dim runningSeconds = CInt(Math.Max(1, (DateTime.UtcNow - callStartedUtc).TotalSeconds))
                                ReportMaintenanceProgress(
                                    progressCallback,
                                    $"{effectiveOperationName}: AI still running ({runningSeconds}s).")
                            End While
                        End Function,
                        heartbeatCts.Token)

                    response = Await SharedMethods.LLM(
                        context:=context,
                        promptSystem:=promptSystem,
                        promptUser:=promptUser,
                        UseSecondAPI:=useAlternateAPI,
                        Hidesplash:=hideSplash).ConfigureAwait(False)

                Finally
                    If heartbeatCts IsNot Nothing Then
                        heartbeatCts.Cancel()
                    End If
                End Try

                If heartbeatTask IsNot Nothing Then
                    Try
                        Await heartbeatTask.ConfigureAwait(False)
                    Catch ex As OperationCanceledException
                    Catch ex As AggregateException When ex.InnerExceptions.All(Function(inner) TypeOf inner Is OperationCanceledException)
                    End Try
                End If

                If heartbeatCts IsNot Nothing Then
                    heartbeatCts.Dispose()
                End If

                Dim elapsedSeconds = CInt(Math.Max(1, (DateTime.UtcNow - callStartedUtc).TotalSeconds))
                ReportMaintenanceProgress(
                    progressCallback,
                    $"{effectiveOperationName}: AI step finished ({elapsedSeconds}s).")

                If Not String.IsNullOrWhiteSpace(response) Then
                    Return response.Trim()
                End If

                If attempt < maxAttempts AndAlso IngestionLlmEmptyRetryDelayMs > 0 Then
                    ReportMaintenanceProgress(progressCallback, $"{effectiveOperationName}: empty AI response, retrying.")
                    Await Task.Delay(IngestionLlmEmptyRetryDelayMs).ConfigureAwait(False)
                End If
            Next

            ReportMaintenanceProgress(progressCallback, $"{effectiveOperationName}: no usable AI response returned.")
            Return ""
        End Function


        Public Shared Async Function ProcessLargeTextInChunksAsync(context As ISharedContext,
                                                                   systemPrompt As String,
                                                                   fullText As String,
                                                                   useAlternateAPI As Boolean,
                                                                   Optional fileName As String = "Document",
                                                                   Optional isBackground As Boolean = False) As Task(Of String)
            Dim chunks As New List(Of String)()
            Dim currentIdx As Integer = 0

            While currentIdx < fullText.Length
                Dim len As Integer = Math.Min(MaxChunkChars, fullText.Length - currentIdx)
                chunks.Add(fullText.Substring(currentIdx, len))
                currentIdx += len
            End While

            Dim rollingSummary As String = ""

            For i As Integer = 0 To chunks.Count - 1
                Dim chunkPrompt = $"Source File: {fileName} (Part {i + 1} of {chunks.Count}){vbCrLf}"
                If Not String.IsNullOrWhiteSpace(rollingSummary) Then
                    chunkPrompt &= $"Previous Context/Summary so far:{vbCrLf}{rollingSummary}{vbCrLf}{vbCrLf}"
                End If
                chunkPrompt &= $"Current Chunk Content:{vbCrLf}{chunks(i)}"

                Dim iterPrompt = If(
                    i = chunks.Count - 1,
                    systemPrompt,
                    "You are a rolling document summarizer. Summarize the following chunk, combining it with previous context. Output ONLY the updated running summary.")

                Dim chunkResult = Await ExecuteKnowledgeStoreLlmWithRetryAsync(
                    context:=context,
                    promptSystem:=iterPrompt,
                    promptUser:=chunkPrompt,
                    useAlternateAPI:=useAlternateAPI,
                    hideSplash:=True,
                    isBackground:=isBackground).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(chunkResult) Then
                    Return ""
                End If

                rollingSummary = chunkResult
            Next

            Return rollingSummary
        End Function

        Public Shared Sub LogWikiError(kbRootPath As String, source As String, errorMessage As String)
            Try
                Dim errorLogPath = Path.Combine(kbRootPath, ".redink", "knowledge_errors.log")
                Dim msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Error in {Path.GetFileName(source)}: {errorMessage}{vbCrLf}"
                File.AppendAllText(errorLogPath, msg, Encoding.UTF8)
            Catch
            End Try
        End Sub

        Private Shared Async Function ExecuteKnowledgeStoreScopedAsync(Of T)(context As ISharedContext,
                                                                             operation As Func(Of Boolean, Task(Of T))) As Task(Of T)
            Dim backupConfig As ModelConfig = Nothing
            Dim useAlternateAPI As Boolean = False

            Try
                If context IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                    backupConfig = SharedMethods.GetCurrentConfig(context)
                    If SharedMethods.GetSpecialTaskModel(context, context.INI_AlternateModelPath, "KnowledgeStore") Then
                        useAlternateAPI = True
                    Else
                        backupConfig = Nothing
                    End If
                End If

                Return Await operation(useAlternateAPI).ConfigureAwait(False)
            Finally
                If backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(context, backupConfig)
                End If
            End Try
        End Function

        Private Shared Async Function ExecuteKnowledgeStoreLlmAsync(context As ISharedContext,
                                                                    promptSystem As String,
                                                                    promptUser As String,
                                                                    Optional hideSplash As Boolean = True,
                                                                    Optional isBackground As Boolean = False,
                                                                    Optional progressCallback As Action(Of String) = Nothing,
                                                                    Optional operationName As String = "Knowledge Store AI step") As Task(Of String)
            Dim gateLease As IDisposable = Nothing

            Try
                gateLease = Await KnowledgeStoreHostGate.EnterLlmExecutionAsync(
                    operationName:=operationName,
                    statusCallback:=progressCallback).ConfigureAwait(False)

                Return Await ExecuteKnowledgeStoreScopedAsync(
                    context,
                    Async Function(useAlternateAPI)
                        Return Await ExecuteKnowledgeStoreLlmWithRetryAsync(
                            context:=context,
                            promptSystem:=promptSystem,
                            promptUser:=promptUser,
                            useAlternateAPI:=useAlternateAPI,
                            hideSplash:=hideSplash,
                            isBackground:=isBackground,
                            progressCallback:=progressCallback,
                            operationName:=operationName).ConfigureAwait(False)
                    End Function).ConfigureAwait(False)
            Finally
                If gateLease IsNot Nothing Then
                    gateLease.Dispose()
                End If
            End Try
        End Function

        Private Shared Function BuildAgentPrompt(roleDescription As String,
                                                 schema As KnowledgeStoreSchema,
                                                 wikiIndex As String,
                                                 candidatePages As List(Of WikiPageInfo),
                                                 defaultKind As String) As String
            Dim sb As New StringBuilder()
            Dim kindsList = FormatKindList(GetEffectivePageKinds(schema))

            If schema IsNot Nothing Then sb.Append(schema.ResolveLanguageDirective())
            sb.AppendLine(roleDescription)
            sb.AppendLine()
            sb.AppendLine("OUTPUT FORMAT REQUIREMENTS:")
            sb.AppendLine("1. First line: TITLE: <concise title>")
            sb.AppendLine("2. Second line: SUMMARY: <one-sentence summary>")
            sb.AppendLine($"3. Third line: KIND: <{kindsList}>")
            sb.AppendLine("4. Start the markdown body after a blank line.")
            sb.AppendLine("5. Use standard markdown links, not wiki syntax.")
            sb.AppendLine("6. Preserve nuance, contradictions, and explicit uncertainty.")
            sb.AppendLine("7. Create meaningful cross-links to relevant existing pages whenever possible.")
            sb.AppendLine("8. Only create clickable markdown page links to pages that already exist in the supplied wiki context.")
            sb.AppendLine("9. If a useful concept does not yet have a page, mention it as plain text or place it under ## Open Questions, but do not create a broken page link.")
            sb.AppendLine("10. Do not invent sources.")
            sb.AppendLine("SOURCE CITATION RULES:")
            sb.AppendLine("- Refer to sources only by ID markers like [S1], [S2] embedded in Key Claims and Evidence bullets.")
            sb.AppendLine("- DO NOT author the ## Sources section; it is generated deterministically.")
            sb.AppendLine("- DO NOT write file paths, file:// URLs, or download links anywhere.")
            If schema?.PerClaimCitations IsNot Nothing AndAlso schema.PerClaimCitations.Enabled Then
                sb.AppendLine("- Every Key Claim and every Evidence bullet MUST end with at least one [S#] marker referring to the source it relies on.")
            End If
            sb.AppendLine()
            sb.AppendLine(BuildSchemaSectionsGuidance(schema))
            sb.AppendLine()

            If schema IsNot Nothing AndAlso schema.DetectContradictions Then
                sb.AppendLine("If evidence conflicts with existing pages, keep both positions and record the conflict explicitly.")
                sb.AppendLine()
            End If

            If schema IsNot Nothing Then
                sb.AppendLine(schema.ToPromptBlock())
            End If

            If Not String.IsNullOrWhiteSpace(wikiIndex) Then
                sb.AppendLine()
                sb.AppendLine("CURRENT WIKI INDEX:")
                sb.AppendLine(wikiIndex)
            End If

            If candidatePages IsNot Nothing AndAlso candidatePages.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("MOST RELEVANT EXISTING PAGES:")
                For Each candidate In candidatePages
                    sb.AppendLine($"- Title: {candidate.Title}")
                    sb.AppendLine($"  Path: {candidate.FileName}")
                    sb.AppendLine($"  Kind: {candidate.Kind}")
                    sb.AppendLine($"  Status: {candidate.Status}")
                    If Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        sb.AppendLine($"  Summary: {candidate.Summary}")
                    End If
                Next
            End If

            sb.AppendLine()
            sb.AppendLine($"Default page kind if unclear: {NormalizePageKind(defaultKind, "Concept", schema)}")

            Return sb.ToString().Trim()
        End Function

        Private Shared Function LoadIndexText(kbRootPath As String) As String
            Try
                Dim indexPath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, KnowledgeStoreCatalog.IndexFile)
                If Not File.Exists(indexPath) Then Return ""
                Dim text = File.ReadAllText(indexPath, Encoding.UTF8)
                If text.Length > 20000 Then
                    Return text.Substring(0, 20000) & vbCrLf & "... [truncated]"
                End If
                Return text
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function FindCandidateWikiPages(kbRootPath As String,
                                               seedText As String,
                                               maxCandidates As Integer) As List(Of WikiPageInfo)
            Dim pages = GetAllWikiPages(kbRootPath).
                Where(Function(p) Not p.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase)).
                ToList()

            If String.IsNullOrWhiteSpace(seedText) OrElse pages.Count = 0 Then Return New List(Of WikiPageInfo)()

            Dim seedTokens = Tokenize(seedText)

            Dim scored = pages.Select(
                Function(p)
                    Dim weightedText =
                        p.Title & " " &
                        p.Title & " " &
                        p.Summary & " " &
                        p.Kind & " " &
                        p.Status & " " &
                        RemoveFrontMatter(p.Content)

                    Dim score = ScoreTokens(seedTokens, Tokenize(weightedText))
                    Return Tuple.Create(p, score)
                End Function).
                Where(Function(x) x.Item2 > 0).
                OrderByDescending(Function(x) x.Item2).
                ThenBy(Function(x) x.Item1.Title, StringComparer.OrdinalIgnoreCase).
                Take(maxCandidates).
                Select(Function(x) x.Item1).
                ToList()

            Return scored
        End Function

        Private Shared Async Function ParseAndSaveAgentResponseAsync(kbRootPath As String,
                                                                     agentResponse As String,
                                                                     context As ISharedContext,
                                                                     sourceFilePath As String,
                                                                     defaultKind As String,
                                                                     relatedCandidates As List(Of WikiPageInfo),
                                                                     actionName As String,
                                                                     Optional additionalSourcePaths As IEnumerable(Of String) = Nothing,
                                                                     Optional schema As KnowledgeStoreSchema = Nothing) As Task(Of Boolean)

            If schema Is Nothing Then schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)

            Dim parsed = ParseAgentResponse(agentResponse, defaultKind)
            If String.IsNullOrWhiteSpace(parsed.Body) Then Return False

            Dim finalTitle = parsed.Title
            Dim finalSummary = parsed.Summary
            Dim finalKind = NormalizePageKind(parsed.Kind, defaultKind, schema)
            Dim finalBody = parsed.Body

            Dim existingPagePath = FindExistingWikiPagePath(kbRootPath, finalTitle)
            Dim isUpdate As Boolean = Not String.IsNullOrWhiteSpace(existingPagePath)
            Dim finalPath As String = existingPagePath
            Dim existingPageInfo As WikiPageInfo = Nothing

            If isUpdate Then
                Try
                    existingPageInfo = ReadWikiPageInfo(kbRootPath, existingPagePath)
                Catch
                    existingPageInfo = Nothing
                End Try
            End If

            If String.IsNullOrWhiteSpace(finalPath) Then
                Dim safeName = SanitizeFileName(finalTitle)
                If String.IsNullOrWhiteSpace(safeName) Then
                    safeName = "WikiPage_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
                End If

                Dim targetDir = GetPageDirectoryForKind(kbRootPath, finalKind)
                If Not Directory.Exists(targetDir) Then
                    Directory.CreateDirectory(targetDir)
                End If

                finalPath = Path.Combine(targetDir, safeName & ".md")
            End If

            Dim mergedSourcePaths = MergeSourcePaths(
                sourceFilePath,
                additionalSourcePaths,
                If(existingPageInfo Is Nothing, Nothing, existingPageInfo.SourcePaths))

            If isUpdate Then
                Dim existingContent = If(
                    existingPageInfo IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(existingPageInfo.Content),
                    existingPageInfo.Content,
                    File.ReadAllText(finalPath, Encoding.UTF8))

                Dim mergedResponse = Await ExecuteKnowledgeStoreLlmAsync(
                    context:=context,
                    promptSystem:=BuildMergePrompt(schema),
                    promptUser:=BuildMergeUserPrompt(
                        existingContent:=existingContent,
                        newContent:=finalBody,
                        relatedCandidates:=relatedCandidates,
                        kbRootPath:=kbRootPath,
                        currentPagePath:=finalPath,
                        schema:=schema),
                    hideSplash:=True).ConfigureAwait(False)

                If Not String.IsNullOrWhiteSpace(mergedResponse) Then
                    Dim merged = ParseAgentResponse(mergedResponse, finalKind)
                    finalTitle = merged.Title
                    finalSummary = merged.Summary
                    finalKind = NormalizePageKind(merged.Kind, finalKind, schema)
                    finalBody = merged.Body
                End If
            End If

            finalBody = NormalizePageBody(
                body:=finalBody,
                currentPagePath:=finalPath,
                sourcePaths:=mergedSourcePaths,
                relatedCandidates:=relatedCandidates,
                schema:=schema)

            Dim repairedLinksCount As Integer = 0
            finalBody = RepairResolvableMarkdownLinks(finalBody, finalPath, kbRootPath, repairedLinksCount)

            Dim removedBrokenLinksCount As Integer = 0
            finalBody = StripUnresolvableInternalLinks(finalBody, finalPath, kbRootPath, removedBrokenLinksCount)

            Dim pageState = ComputePageComputedState(finalBody, mergedSourcePaths, schema)

            If isUpdate AndAlso
               existingPageInfo IsNot Nothing AndAlso
               If(actionName, "").StartsWith("repair", StringComparison.OrdinalIgnoreCase) AndAlso
               IsRiskyAutoApplyRewrite(existingPageInfo, pageState) Then

                AppendToLog(
                    kbRootPath,
                    "repair-skip",
                    $"{finalTitle} | Risky LLM rewrite skipped | OldStatus={existingPageInfo.Status}; NewStatus={pageState.Status}; OldSources={existingPageInfo.SourceCount}; NewSources={pageState.SourceCount}; OldContradictions={existingPageInfo.ContradictionCount}; NewContradictions={pageState.ContradictionCount}")

                Return False
            End If

            ' Create dossiers and deterministically rewrite the Sources section
            ' BEFORE composing/saving, so the file on disk reflects the rewrite.
            Dim dossierRefs = EnsureDossiersForPage(kbRootPath, finalPath, finalTitle, mergedSourcePaths, schema)

            If schema Is Nothing OrElse schema.AlwaysAddSourceLinks Then
                finalBody = RebuildSourcesSectionDeterministically(finalBody, finalPath, kbRootPath, dossierRefs)
            End If

            ' Recompute computed state on the rewritten body so front-matter
            ' source_count and status reflect the deterministic Sources block.
            pageState = ComputePageComputedState(finalBody, mergedSourcePaths, schema)

            Dim fullDocument = ComposePageDocument(
                title:=finalTitle,
                summary:=finalSummary,
                kind:=finalKind,
                body:=finalBody,
                sourceFilePath:="",
                additionalSourcePaths:=mergedSourcePaths,
                status:=pageState.Status,
                reviewNeeded:=pageState.ReviewNeeded,
                contradictionCount:=pageState.ContradictionCount,
                sourceCount:=pageState.SourceCount)

            Try
                File.WriteAllText(finalPath, fullDocument, Encoding.UTF8)
                Await QueueOrUpdateEmbeddingAsync(kbRootPath, finalPath, fullDocument, context).ConfigureAwait(False)

                Dim savedPageInfo As New WikiPageInfo With {
                    .FilePath = finalPath,
                    .FileName = Path.GetFileName(finalPath),
                    .RelativePath = GetRelativeWikiPath(kbRootPath, finalPath),
                    .Title = finalTitle,
                    .Summary = finalSummary,
                    .Kind = finalKind,
                    .Status = pageState.Status,
                    .ReviewNeeded = pageState.ReviewNeeded,
                    .ContradictionCount = pageState.ContradictionCount,
                    .SourceCount = pageState.SourceCount,
                    .Content = fullDocument,
                    .SourcePaths = mergedSourcePaths
                }

                ' The wiki on disk has changed; force the next snapshot read to refresh.
                InvalidateIngestPageCache(kbRootPath)

                Dim backfilledRelatedPages As Integer = 0
                Dim deferredMaintenance As Boolean = IsIngestScopeActive(kbRootPath)

                If deferredMaintenance Then
                    ' Inside an ingest scope: collapse N per-save maintenance passes
                    ' into one outermost flush in EndIngestScopeAsync.
                    TrackIngestSavedPage(kbRootPath, savedPageInfo)
                Else
                    backfilledRelatedPages = Await BackfillRelatedLinksForExistingPagesAsync(
                        kbRootPath:=kbRootPath,
                        savedPage:=savedPageInfo,
                        context:=context,
                        schema:=schema).ConfigureAwait(False)

                    RebuildIndex(kbRootPath)
                    RefreshReviewArtifacts(kbRootPath)
                End If

                Dim logAction = If(isUpdate, "update", actionName)
                AppendToLog(
                    kbRootPath,
                    logAction,
                    $"{finalTitle}{If(String.IsNullOrWhiteSpace(sourceFilePath), "", " | " & Path.GetFileName(sourceFilePath))} | Kind={finalKind}; Status={pageState.Status}; Review={pageState.ReviewNeeded}; Contradictions={pageState.ContradictionCount}; LinksRepaired={repairedLinksCount}; BrokenLinksRemoved={removedBrokenLinksCount}; BackfilledRelatedPages={backfilledRelatedPages}{If(deferredMaintenance, "; Deferred=true", "")}")

                Return True
            Catch ex As Exception
                LogWikiError(kbRootPath, sourceFilePath, ex.Message)
                Return False
            End Try
        End Function



        Private Shared Function RebuildSourcesSectionDeterministically(body As String,
                                                               currentPagePath As String,
                                                               kbRootPath As String,
                                                               dossiers As List(Of KnowledgeSourceDossierService.DossierRef)) As String
            If dossiers Is Nothing OrElse dossiers.Count = 0 Then
                ' Nothing to render; if a stub Sources section exists, leave it alone.
                Return body
            End If

            Dim wikiRoot = GetWikiRootPath(kbRootPath)
            Dim currentDir = Path.GetDirectoryName(currentPagePath)

            Dim sb As New StringBuilder()
            Dim sourcesHeading As String = "## Sources"
            Dim wikiSchema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            If wikiSchema IsNot Nothing Then sourcesHeading = wikiSchema.GetLabel("SourcesHeading", sourcesHeading)
            sb.AppendLine(sourcesHeading)
            sb.AppendLine()

            Dim idx As Integer = 1
            For Each d In dossiers
                Dim dossierAbs = Path.Combine(wikiRoot, d.DossierRelPath.Replace("/"c, Path.DirectorySeparatorChar))
                Dim relFromPage = NormalizeWikiRelativePath(GetRelativePathCompat(currentDir, dossierAbs))
                Dim blobAbs = Path.Combine(wikiRoot, d.BlobRelPath.Replace("/"c, Path.DirectorySeparatorChar))
                Dim blobRelFromPage = NormalizeWikiRelativePath(GetRelativePathCompat(currentDir, blobAbs))

                ' Format: [S1] [Title](dossier.md) — [download](../_files/hash.ext)
                sb.AppendLine($"- [S{idx}] [{d.Title}]({relFromPage}) — [download]({blobRelFromPage})")
                idx += 1
            Next

            Return ReplaceMarkdownSection(body, sourcesHeading, sb.ToString().Trim())
        End Function


        Private Shared Function ParseAgentResponse(agentResponse As String, defaultKind As String) As ParsedAgentResponse
            Dim result As New ParsedAgentResponse()
            Dim lines = agentResponse.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None)
            Dim contentStartIdx As Integer = -1

            result.Kind = defaultKind

            For i As Integer = 0 To Math.Min(lines.Length - 1, 8)
                Dim line = lines(i).Trim()

                If line.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                    result.Title = line.Substring(6).Trim()
                ElseIf line.StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) Then
                    result.Summary = line.Substring(8).Trim()
                ElseIf line.StartsWith("KIND:", StringComparison.OrdinalIgnoreCase) Then
                    result.Kind = line.Substring(5).Trim()
                ElseIf contentStartIdx = -1 AndAlso Not String.IsNullOrWhiteSpace(line) Then
                    contentStartIdx = i
                End If
            Next

            If contentStartIdx < 0 Then
                contentStartIdx = 0
                While contentStartIdx < lines.Length AndAlso
                      (lines(contentStartIdx).Trim().StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) OrElse
                       lines(contentStartIdx).Trim().StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) OrElse
                       lines(contentStartIdx).Trim().StartsWith("KIND:", StringComparison.OrdinalIgnoreCase) OrElse
                       String.IsNullOrWhiteSpace(lines(contentStartIdx)))
                    contentStartIdx += 1
                End While
            End If

            If contentStartIdx < lines.Length Then
                result.Body = String.Join(vbCrLf, lines.Skip(contentStartIdx)).Trim()
            End If

            If String.IsNullOrWhiteSpace(result.Title) Then
                result.Title = "Untitled Concept"
            End If

            If String.IsNullOrWhiteSpace(result.Summary) Then
                result.Summary = "Auto-generated wiki page."
            End If

            If String.IsNullOrWhiteSpace(result.Kind) Then
                result.Kind = defaultKind
            End If

            Return result
        End Function


        Private Shared Function MergeSourcePaths(sourceFilePath As String,
                                                 additionalSourcePaths As IEnumerable(Of String),
                                                 Optional existingSourcePaths As IEnumerable(Of String) = Nothing) As List(Of String)
            Dim allSourcePaths As New List(Of String)()

            If existingSourcePaths IsNot Nothing Then
                allSourcePaths.AddRange(
                    existingSourcePaths.
                        Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                        Select(Function(p) p.Trim()))
            End If

            If Not String.IsNullOrWhiteSpace(sourceFilePath) Then
                allSourcePaths.Add(sourceFilePath.Trim())
            End If

            If additionalSourcePaths IsNot Nothing Then
                allSourcePaths.AddRange(
                    additionalSourcePaths.
                        Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                        Select(Function(p) p.Trim()))
            End If

            Return allSourcePaths.
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function NormalizePageBody(body As String,
                                                  currentPagePath As String,
                                                  sourcePaths As IEnumerable(Of String),
                                                  relatedCandidates As List(Of WikiPageInfo)) As String
            Return NormalizePageBody(body, currentPagePath, sourcePaths, relatedCandidates, Nothing)
        End Function

        Private Shared Function NormalizePageBody(body As String,
                                                  currentPagePath As String,
                                                  sourcePaths As IEnumerable(Of String),
                                                  relatedCandidates As List(Of WikiPageInfo),
                                                  schema As KnowledgeStoreSchema) As String
            Dim normalized = If(body, "").Trim()
            Dim sections = GetEffectiveRequiredSections(schema)

            ' Headings that are managed by the deterministic Sources / Related Pages
            ' rewriters must NOT be added as placeholders here, regardless of which
            ' label (English default or schema-configured) the user picked.
            Dim sourcesHeadingForSkip As String = "## Sources"
            Dim relatedHeadingForSkip As String = "## Related Pages"
            If schema IsNot Nothing Then
                sourcesHeadingForSkip = schema.GetLabel("SourcesHeading", sourcesHeadingForSkip)
                relatedHeadingForSkip = schema.GetLabel("RelatedPagesHeading", relatedHeadingForSkip)
            End If

            Dim placeholderBullet As String = "- _Not yet provided._"
            If schema IsNot Nothing Then
                placeholderBullet = "- " & schema.GetLabel("PlaceholderNotProvided", "_Not yet provided._")
            End If

            ' Ensure each schema-required section exists (with a neutral placeholder bullet),
            ' except those owned by the deterministic rewriters.
            For Each heading In sections
                If heading.Equals("## Sources", StringComparison.OrdinalIgnoreCase) OrElse
                   heading.Equals(sourcesHeadingForSkip, StringComparison.OrdinalIgnoreCase) OrElse
                   heading.Equals("## Related Pages", StringComparison.OrdinalIgnoreCase) OrElse
                   heading.Equals(relatedHeadingForSkip, StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
                normalized = EnsureSection(normalized, heading, placeholderBullet)
            Next

            Dim analyticalHeadings As New List(Of String) From {
                GetClaimsHeading(schema),
                GetEvidenceHeading(schema),
                GetOpenQuestionsHeading(schema)
            }

            If schema Is Nothing OrElse schema.DetectContradictions Then
                analyticalHeadings.Add(GetContradictionsHeading(schema))
            End If

            analyticalHeadings = analyticalHeadings.
                Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            For Each analytical In analyticalHeadings
                If sections.Any(Function(x) x.Equals(analytical, StringComparison.OrdinalIgnoreCase)) Then
                    normalized = NormalizeMarkdownListSection(
                        normalized,
                        analytical,
                        GetDefaultSectionBullet(analytical, schema))
                End If
            Next

            Dim alwaysCrossLinks As Boolean = schema Is Nothing OrElse schema.AlwaysCreateCrossLinks

            ' NOTE: The Sources section is rendered exclusively by
            ' RebuildSourcesSectionDeterministically (called from ParseAndSaveAgentResponseAsync
            ' and from the lint migration block). The legacy EnsureSourcesSection is no longer
            ' called here, because it (a) hard-codes "## Sources" in English regardless of the
            ' configured SourcesHeading label and (b) writes non-clickable [name](name) links
            ' that point at nothing. Calling it here used to produce a duplicate, broken
            ' Sources block alongside the deterministic one in localized stores.

            If alwaysCrossLinks Then
                normalized = EnsureRelatedSection(normalized, currentPagePath, relatedCandidates, schema)
            End If

            Return normalized.Trim()
        End Function

        Private Shared Function EnsureSection(body As String,
                                              heading As String,
                                              defaultBullet As String) As String
            If body.IndexOf(heading, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return body
            End If

            Dim sb As New StringBuilder()

            If Not String.IsNullOrWhiteSpace(body) Then
                sb.AppendLine(body.Trim())
                sb.AppendLine()
            End If

            sb.AppendLine(heading)
            sb.AppendLine()
            sb.AppendLine(defaultBullet)

            Return sb.ToString().Trim()
        End Function


        Private Shared Function NormalizeMarkdownListSection(markdown As String,
                                                             heading As String,
                                                             defaultBullet As String) As String
            Dim sectionItems = ExtractSectionBulletItems(markdown, heading)

            Dim normalizedItems = sectionItems.
                Select(Function(x) Regex.Replace(If(x, "").Trim(), "\s+", " ").Trim()).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            normalizedItems = DeduplicatePreservingOrder(normalizedItems)

            If normalizedItems.Count = 0 Then
                normalizedItems.Add(defaultBullet.TrimStart("-"c, " "c))
            End If

            Dim rebuiltSection As New StringBuilder()
            rebuiltSection.AppendLine(heading)
            rebuiltSection.AppendLine()
            For Each item In normalizedItems
                rebuiltSection.AppendLine("- " & item)
            Next

            Return ReplaceMarkdownSection(markdown, heading, rebuiltSection.ToString().Trim())
        End Function

        Private Shared Function ReplaceMarkdownSection(markdown As String,
                                                       heading As String,
                                                       replacementSection As String) As String
            Dim normalized = If(markdown, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim pattern = "(?ims)^" & Regex.Escape(heading.Trim()) & "\s*$.*?(?=^##\s+|\z)"

            If Regex.IsMatch(normalized, pattern) Then
                Dim safeReplacement = vbLf & vbLf & replacementSection.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf) & vbLf & vbLf
                normalized = Regex.Replace(
                    normalized,
                    pattern,
                    safeReplacement)
            Else
                If Not String.IsNullOrWhiteSpace(normalized) Then
                    normalized &= vbLf & vbLf
                End If
                normalized &= replacementSection.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            End If

            normalized = Regex.Replace(normalized, "\n{3,}", vbLf & vbLf)
            Return normalized.Replace(vbLf, vbCrLf).Trim()
        End Function

        Private Shared Function DeduplicatePreservingOrder(items As IEnumerable(Of String)) As List(Of String)
            Dim result As New List(Of String)()

            If items Is Nothing Then Return result

            For Each item In items
                Dim normalized = If(item, "").Trim()
                If String.IsNullOrWhiteSpace(normalized) Then
                    Continue For
                End If

                Dim isDuplicate = result.Any(
                    Function(existing)
                        Return AreSectionItemsNearDuplicates(existing, normalized)
                    End Function)

                If Not isDuplicate Then
                    result.Add(normalized)
                End If
            Next

            Return result
        End Function

        Private Shared Function AreSectionItemsNearDuplicates(a As String, b As String) As Boolean
            Dim valueA = If(a, "").Trim()
            Dim valueB = If(b, "").Trim()

            If String.IsNullOrWhiteSpace(valueA) OrElse String.IsNullOrWhiteSpace(valueB) Then
                Return False
            End If

            If valueA.Equals(valueB, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim keyA = NormalizeWikiKey(valueA)
            Dim keyB = NormalizeWikiKey(valueB)

            If Not String.IsNullOrWhiteSpace(keyA) AndAlso keyA.Equals(keyB, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            Dim tokensA = Tokenize(valueA)
            Dim tokensB = Tokenize(valueB)

            If tokensA.Count = 0 OrElse tokensB.Count = 0 Then
                Return False
            End If

            Dim similarity = Jaccard(tokensA, tokensB)
            If similarity >= 0.8 Then
                Return True
            End If

            Dim overlap = tokensA.Intersect(tokensB, StringComparer.OrdinalIgnoreCase).Count()
            Dim minCount = Math.Min(tokensA.Count, tokensB.Count)

            If minCount >= 3 AndAlso overlap >= minCount Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function GetMarkdownSectionContent(markdown As String, heading As String) As String
            Dim normalized = If(markdown, "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If String.IsNullOrWhiteSpace(normalized) Then Return ""

            Dim headingMatch = Regex.Match(
                normalized,
                "(?im)^" & Regex.Escape(heading.Trim()) & "\s*$")

            If Not headingMatch.Success Then Return ""

            Dim startIdx = headingMatch.Index + headingMatch.Length

            While startIdx < normalized.Length AndAlso
                  (normalized.Chars(startIdx) = ControlChars.Lf OrElse normalized.Chars(startIdx) = ControlChars.Cr)
                startIdx += 1
            End While

            Dim remaining = normalized.Substring(startIdx)
            Dim nextHeading = Regex.Match(remaining, "(?im)^##\s+")

            If nextHeading.Success Then
                remaining = remaining.Substring(0, nextHeading.Index)
            End If

            Return remaining.Trim().Replace(vbLf, vbCrLf)
        End Function

        Private Shared Function ExtractSectionBulletItems(markdown As String, heading As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim sectionText = GetMarkdownSectionContent(markdown, heading)

            If String.IsNullOrWhiteSpace(sectionText) Then
                Return result
            End If

            For Each rawLine In sectionText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim line = rawLine.Trim()
                If line.StartsWith("- ", StringComparison.Ordinal) Then
                    result.Add(line.Substring(2).Trim())
                ElseIf line.StartsWith("* ", StringComparison.Ordinal) Then
                    result.Add(line.Substring(2).Trim())
                End If
            Next

            Return result
        End Function

        Private Shared Function IsMeaningfulSectionItem(value As String) As Boolean
            Dim normalized = If(value, "").Trim()

            If String.IsNullOrWhiteSpace(normalized) Then Return False
            If normalized.Equals("None noted.", StringComparison.OrdinalIgnoreCase) Then Return False
            If normalized.Equals("None recorded yet.", StringComparison.OrdinalIgnoreCase) Then Return False
            If normalized.Equals("No durable claims extracted yet.", StringComparison.OrdinalIgnoreCase) Then Return False
            If normalized.Equals("None recorded.", StringComparison.OrdinalIgnoreCase) Then Return False

            Return True
        End Function

        Private Shared Function ComputePageComputedState(body As String,
                                                         sourcePaths As IEnumerable(Of String)) As WikiComputedState
            Return ComputePageComputedState(body, sourcePaths, Nothing)
        End Function

        Private Shared Function ComputePageComputedState(body As String,
                                                         sourcePaths As IEnumerable(Of String),
                                                         schema As KnowledgeStoreSchema) As WikiComputedState
            Dim result As New WikiComputedState()
            Dim mergedSourcePaths = MergeSourcePaths("", sourcePaths)

            Dim claimsHeading = GetClaimsHeading(schema)
            Dim evidenceHeading = GetEvidenceHeading(schema)
            Dim contradictionsHeading = GetContradictionsHeading(schema)

            Dim meaningfulEvidence = ExtractSectionBulletItems(body, evidenceHeading).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            Dim meaningfulContradictions As New List(Of String)()
            If schema Is Nothing OrElse schema.DetectContradictions Then
                meaningfulContradictions = ExtractSectionBulletItems(body, contradictionsHeading).
                    Where(Function(x) IsMeaningfulSectionItem(x)).
                    ToList()
            End If

            Dim meaningfulClaims = ExtractSectionBulletItems(body, claimsHeading).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            result.SourceCount = mergedSourcePaths.Count
            result.ContradictionCount = meaningfulContradictions.Count

            Dim hasSupport As Boolean = meaningfulEvidence.Count > 0 OrElse mergedSourcePaths.Count > 0

            If body.IndexOf("superseded", StringComparison.OrdinalIgnoreCase) >= 0 Then
                result.Status = "superseded"
            ElseIf meaningfulContradictions.Count > 0 Then
                result.Status = "disputed"
            ElseIf meaningfulClaims.Count = 0 OrElse Not hasSupport OrElse mergedSourcePaths.Count = 0 Then
                result.Status = "tentative"
            Else
                result.Status = "stable"
            End If

            result.ReviewNeeded =
                meaningfulContradictions.Count > 0 OrElse
                body.IndexOf("review needed", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                result.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase)

            Return result
        End Function

        Private Shared Function ExtractClaimsFromPage(page As WikiPageInfo) As List(Of WikiClaimInfo)
            Return ExtractClaimsFromPage(page, Nothing)
        End Function

        Private Shared Function ExtractClaimsFromPage(page As WikiPageInfo,
                                                      schema As KnowledgeStoreSchema) As List(Of WikiClaimInfo)
            Dim result As New List(Of WikiClaimInfo)()
            If page Is Nothing Then Return result

            Dim claimsHeading = GetClaimsHeading(schema)
            Dim evidenceHeading = GetEvidenceHeading(schema)
            Dim contradictionsHeading = GetContradictionsHeading(schema)

            Dim claimItems = ExtractSectionBulletItems(page.Content, claimsHeading).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            Dim evidenceItems = ExtractSectionBulletItems(page.Content, evidenceHeading).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            Dim contradictionItems As New List(Of String)()
            If schema Is Nothing OrElse schema.DetectContradictions Then
                contradictionItems = ExtractSectionBulletItems(page.Content, contradictionsHeading).
                    Where(Function(x) IsMeaningfulSectionItem(x)).
                    ToList()
            End If

            For Each claimItem In claimItems
                Dim matchedEvidence = GetEvidenceItemsSupportingClaim(claimItem, evidenceItems)

                Dim claim As New WikiClaimInfo() With {
                    .ClaimText = claimItem,
                    .EvidenceItems = matchedEvidence,
                    .SourcePaths = page.SourcePaths.ToList(),
                    .IsContradicted = contradictionItems.Any(
                        Function(c) AreClaimsPotentiallyContradictory(claimItem, c)),
                    .NeedsReview = page.ReviewNeeded OrElse
                                   page.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) OrElse
                                   (matchedEvidence.Count = 0 AndAlso page.SourcePaths.Count = 0)
                }

                result.Add(claim)
            Next

            Return result
        End Function

        Private Shared Function CountClaimsWithoutEvidence(page As WikiPageInfo) As Integer
            Return CountClaimsWithoutEvidence(page, Nothing)
        End Function

        Private Shared Function CountClaimsWithoutEvidence(page As WikiPageInfo,
                                                           schema As KnowledgeStoreSchema) As Integer
            If page Is Nothing Then Return 0

            Return ExtractClaimsFromPage(page, schema).
                Where(Function(c) c.EvidenceItems Is Nothing OrElse c.EvidenceItems.Count = 0).
                Count()
        End Function

        Private Shared Function CountContradictedClaims(page As WikiPageInfo) As Integer
            Return CountContradictedClaims(page, Nothing)
        End Function

        Private Shared Function CountContradictedClaims(page As WikiPageInfo,
                                                        schema As KnowledgeStoreSchema) As Integer
            If page Is Nothing Then Return 0

            Return ExtractClaimsFromPage(page, schema).
                Where(Function(c) c.IsContradicted).
                Count()
        End Function

        Private Shared Function CountClaimsWithoutSources(page As WikiPageInfo) As Integer
            Return CountClaimsWithoutSources(page, Nothing)
        End Function

        Private Shared Function CountClaimsWithoutSources(page As WikiPageInfo,
                                                          schema As KnowledgeStoreSchema) As Integer
            If page Is Nothing Then Return 0

            Dim claims = ExtractClaimsFromPage(page, schema)
            If claims.Count = 0 Then Return 0
            If page.SourcePaths Is Nothing OrElse page.SourcePaths.Count = 0 Then Return claims.Count

            Return 0
        End Function

        Private Shared Function NormalizePageStatus(status As String) As String
            Dim value = If(status, "").Trim()

            If value.Equals("stable", StringComparison.OrdinalIgnoreCase) Then Return "stable"
            If value.Equals("tentative", StringComparison.OrdinalIgnoreCase) Then Return "tentative"
            If value.Equals("disputed", StringComparison.OrdinalIgnoreCase) Then Return "disputed"
            If value.Equals("superseded", StringComparison.OrdinalIgnoreCase) Then Return "superseded"

            Return "stable"
        End Function


        Private Shared Function GetEvidenceItemsSupportingClaim(claimText As String,
                                                                evidenceItems As List(Of String)) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(claimText) OrElse evidenceItems Is Nothing OrElse evidenceItems.Count = 0 Then
                Return result
            End If

            Dim claimTokens = Tokenize(claimText)

            Dim scoredEvidence = evidenceItems.
                Select(
                    Function(evidenceItem)
                        Dim evidenceTokens = Tokenize(evidenceItem)
                        Dim overlap = claimTokens.Intersect(evidenceTokens, StringComparer.OrdinalIgnoreCase).Count()

                        If Regex.IsMatch(
                            evidenceItem,
                            "\b(source|report|document|states|reported|according to|file|memo|email|analysis)\b",
                            RegexOptions.IgnoreCase) Then
                            overlap += 1
                        End If

                        Return Tuple.Create(evidenceItem, overlap)
                    End Function).
                OrderByDescending(Function(x) x.Item2).
                ThenBy(Function(x) x.Item1, StringComparer.OrdinalIgnoreCase).
                ToList()

            For Each item In scoredEvidence
                If item.Item2 >= 2 Then
                    result.Add(item.Item1)
                End If
            Next

            result = DeduplicatePreservingOrder(result)

            If result.Count = 0 AndAlso evidenceItems.Count = 1 Then
                result.Add(evidenceItems(0))
            ElseIf result.Count = 0 AndAlso scoredEvidence.Count > 0 AndAlso scoredEvidence(0).Item2 > 0 Then
                result.Add(scoredEvidence(0).Item1)
            End If

            Return result.Take(3).ToList()
        End Function


        Private Shared Function FindPotentialSupersededPages(pages As List(Of WikiPageInfo)) As List(Of String)
            Return FindPotentialSupersededPages(pages, Nothing)
        End Function

        Private Shared Function FindPotentialSupersededPages(pages As List(Of WikiPageInfo),
                                                             schema As KnowledgeStoreSchema) As List(Of String)
            Dim result As New List(Of String)()
            If pages Is Nothing OrElse pages.Count = 0 Then Return result

            Dim seenKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Per-page ranked neighbor list. Score >= 3 mirrors the historical
            ' threshold used by the legacy LINQ path.
            Dim byPage As New Dictionary(Of String, List(Of Tuple(Of WikiPageInfo, Double)))(StringComparer.OrdinalIgnoreCase)

            For Each pair As CandidatePagePair In EnumerateCandidatePairs(pages, schema)
                Dim pageA As WikiPageInfo = pair.PageA
                Dim pageB As WikiPageInfo = pair.PageB

                Dim seedA = pageA.Title & " " & pageA.Summary & " " & RemoveFrontMatter(pageA.Content)
                Dim seedB = pageB.Title & " " & pageB.Summary & " " & RemoveFrontMatter(pageB.Content)
                Dim score = ScoreTokens(Tokenize(seedA), Tokenize(seedB))
                If score < 3 Then Continue For

                AddRankedCandidate(byPage, pageA, pageB, score)
                AddRankedCandidate(byPage, pageB, pageA, score)
            Next

            For Each page In pages
                Dim ranked As List(Of Tuple(Of WikiPageInfo, Double)) = Nothing
                If Not byPage.TryGetValue(page.FilePath, ranked) Then Continue For

                Dim candidates = ranked.
                    OrderByDescending(Function(x) x.Item2).
                    Take(5).
                    Select(Function(x) x.Item1).
                    ToList()

                For Each candidate In candidates
                    Dim pairKey = String.Join(
                        "|",
                        New String() {page.FilePath, candidate.FilePath}.
                            OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase))

                    If seenKeys.Contains(pairKey) Then
                        Continue For
                    End If

                    Dim candidateIsStronger =
                        candidate.SourceCount > page.SourceCount OrElse
                        (candidate.SourceCount = page.SourceCount AndAlso
                         candidate.Status.Equals("stable", StringComparison.OrdinalIgnoreCase) AndAlso
                         Not page.Status.Equals("stable", StringComparison.OrdinalIgnoreCase))

                    Dim candidateIsNewer =
                        candidate.UpdatedUtc.HasValue AndAlso
                        page.UpdatedUtc.HasValue AndAlso
                        candidate.UpdatedUtc.Value > page.UpdatedUtc.Value

                    Dim currentIsWeaker =
                        page.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase) OrElse
                        page.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase)

                    If candidateIsStronger AndAlso (candidateIsNewer OrElse currentIsWeaker) Then
                        result.Add(
                            $"{page.RelativePath} may be superseded by {candidate.RelativePath} " &
                            $"[CurrentStatus={page.Status}; CandidateStatus={candidate.Status}; CurrentSources={page.SourceCount}; CandidateSources={candidate.SourceCount}]")
                        seenKeys.Add(pairKey)
                    End If
                Next
            Next

            Return result.
                Distinct(StringComparer.OrdinalIgnoreCase).
                Take(20).
                ToList()
        End Function



        Private Shared Function GetPotentialSupersededPagePaths(pages As List(Of WikiPageInfo)) As HashSet(Of String)
            Return GetPotentialSupersededPagePaths(pages, Nothing)
        End Function

        Private Shared Function GetPotentialSupersededPagePaths(pages As List(Of WikiPageInfo),
                                                                schema As KnowledgeStoreSchema) As HashSet(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each item In FindPotentialSupersededPages(pages, schema)
                Dim marker = " may be superseded by "
                Dim idx = item.IndexOf(marker, StringComparison.OrdinalIgnoreCase)

                If idx > 0 Then
                    Dim relativePath = item.Substring(0, idx).Trim()
                    If Not String.IsNullOrWhiteSpace(relativePath) Then
                        result.Add(relativePath)
                    End If
                End If
            Next

            Return result
        End Function


        Private Shared Function FindStrongestContradictionForPage(page As WikiPageInfo,
                                                                  contradictionPairs As List(Of PotentialContradictionPair)) As PotentialContradictionPair
            If page Is Nothing OrElse contradictionPairs Is Nothing OrElse contradictionPairs.Count = 0 Then
                Return Nothing
            End If

            Return contradictionPairs.FirstOrDefault(
                Function(p)
                    Return p.PageA.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase) OrElse
                           p.PageB.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase)
                End Function)
        End Function

        Private Shared Function BuildRepairPromptUser(page As WikiPageInfo,
                                                      candidates As List(Of WikiPageInfo),
                                                      contradictionPairs As List(Of PotentialContradictionPair),
                                                      kbRootPath As String,
                                                      Optional schema As KnowledgeStoreSchema = Nothing) As String
            Dim promptUser As New StringBuilder()
            Dim contradictionsHeading = GetContradictionsHeading(schema)

            promptUser.AppendLine("CURRENT PAGE:")
            promptUser.AppendLine(page.Content)
            promptUser.AppendLine()
            promptUser.AppendLine("CURRENT INDEX:")
            promptUser.AppendLine(LoadIndexText(kbRootPath))
            promptUser.AppendLine()

            If candidates IsNot Nothing AndAlso candidates.Count > 0 Then
                promptUser.AppendLine("RELATED CANDIDATES:")
                For Each candidate In candidates
                    promptUser.AppendLine($"- {candidate.Title} | {candidate.RelativePath} | {candidate.Summary} | Status={candidate.Status}")
                Next
                promptUser.AppendLine()
            End If

            Dim contradiction = FindStrongestContradictionForPage(page, contradictionPairs)
            If contradiction IsNot Nothing Then
                Dim otherPage = If(
                    contradiction.PageA.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase),
                    contradiction.PageB,
                    contradiction.PageA)

                promptUser.AppendLine("STRONGEST CONTRADICTION CONTEXT:")
                promptUser.AppendLine($"- Current page: {page.RelativePath}")
                promptUser.AppendLine($"- Conflicting page: {otherPage.RelativePath}")
                promptUser.AppendLine($"- Current-page claim: {If(contradiction.PageA.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase), contradiction.ClaimA, contradiction.ClaimB)}")
                promptUser.AppendLine($"- Conflicting-page claim: {If(contradiction.PageA.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase), contradiction.ClaimB, contradiction.ClaimA)}")
                promptUser.AppendLine("CONFLICTING PAGE EXCERPT:")
                promptUser.AppendLine(TruncateForPrompt(RemoveFrontMatter(otherPage.Content), 1600))
                promptUser.AppendLine()
                promptUser.AppendLine($"Repair goal: preserve the conflict explicitly in {contradictionsHeading} if unresolved; do not smooth it away.")
                promptUser.AppendLine()
            End If

            Return promptUser.ToString().Trim()
        End Function

        Private Shared Function BuildReviewQueueDocument(pages As List(Of WikiPageInfo),
                                                         contradictionPairs As List(Of PotentialContradictionPair),
                                                         claimsWithoutEvidence As List(Of String),
                                                         claimsWithoutSources As List(Of String),
                                                         supersededCandidates As List(Of String),
                                                         Optional skippedRepairItems As List(Of String) = Nothing) As String
            Dim sb As New StringBuilder()
            Dim potentialSupersededPagePaths = supersededCandidates.
                Select(
                    Function(item)
                        Dim marker = " may be superseded by "
                        Dim idx = item.IndexOf(marker, StringComparison.OrdinalIgnoreCase)
                        If idx <= 0 Then Return ""
                        Return item.Substring(0, idx).Trim()
                    End Function).
                Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                ToHashSet(StringComparer.OrdinalIgnoreCase)

            sb.AppendLine("# Review Queue")
            sb.AppendLine()
            sb.AppendLine($"> Auto-generated on {DateTime.Now:yyyy-MM-dd HH:mm}.")
            sb.AppendLine()

            sb.AppendLine("## Pages Requiring Review")
            sb.AppendLine()

            Dim reviewPages = pages.
                Where(
                    Function(p)
                        Return p.ReviewNeeded OrElse
                               p.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) OrElse
                               p.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase) OrElse
                               potentialSupersededPagePaths.Contains(p.RelativePath)
                    End Function).
                OrderBy(Function(p) p.Title, StringComparer.OrdinalIgnoreCase).
                ToList()

            If reviewPages.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each page In reviewPages
                    Dim updatedText = If(page.UpdatedUtc.HasValue,
                                         page.UpdatedUtc.Value.ToString("yyyy-MM-dd"),
                                         "unknown")

                    sb.AppendLine(
                        $"- [{page.Title}]({page.RelativePath}) — Status={page.Status}; Contradictions={page.ContradictionCount}; Sources={page.SourceCount}; Updated={updatedText}{If(potentialSupersededPagePaths.Contains(page.RelativePath), "; PossibleSuperseded=true", "")}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("## Claims Without Evidence")
            sb.AppendLine()

            If claimsWithoutEvidence.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each item In claimsWithoutEvidence.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    sb.AppendLine($"- {item}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("## Claims Without Sources")
            sb.AppendLine()

            If claimsWithoutSources.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each item In claimsWithoutSources.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    sb.AppendLine($"- {item}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("## Potential Contradictions")
            sb.AppendLine()

            If contradictionPairs.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each pair In contradictionPairs
                    sb.AppendLine($"- {pair.PageA.RelativePath} <-> {pair.PageB.RelativePath}")
                    sb.AppendLine($"  - Claim A: {pair.ClaimA}")
                    sb.AppendLine($"  - Claim B: {pair.ClaimB}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("## Possible Superseded Pages")
            sb.AppendLine()

            If supersededCandidates.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each item In supersededCandidates.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    sb.AppendLine($"- {item}")
                Next
            End If

            sb.AppendLine()
            sb.AppendLine("## Skipped Risky Repairs")
            sb.AppendLine()

            If skippedRepairItems Is Nothing OrElse skippedRepairItems.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each item In skippedRepairItems.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    sb.AppendLine($"- {item}")
                Next
            End If

            Return sb.ToString().Trim() & vbCrLf
        End Function

        Private Shared Function LoadRecentRiskyRepairSkips(kbRootPath As String) As List(Of String)
            Dim result As New List(Of String)()

            Try
                Dim logPath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, KnowledgeStoreCatalog.LogFile)
                If Not File.Exists(logPath) Then
                    Return result
                End If

                Dim lines = File.ReadAllLines(logPath, Encoding.UTF8)

                result = lines.
                    Where(Function(line) line.IndexOf("repair-skip", StringComparison.OrdinalIgnoreCase) >= 0).
                    Select(Function(line) line.Trim()).
                    Take(20).
                    ToList()
            Catch ex As Exception
                LogWikiError(kbRootPath, "repair-skip-log", ex.Message)
            End Try

            Return result
        End Function


        Private Shared Sub RefreshReviewArtifacts(kbRootPath As String)
            Try
                If String.IsNullOrWhiteSpace(kbRootPath) Then Return

                Dim pages = GetAllWikiPages(kbRootPath)
                If pages.Count = 0 Then Return

                Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
                Dim contradictionPairs = FindPotentialContradictionPairs(pages, schema)
                Dim supersededCandidates = FindPotentialSupersededPages(pages, schema)

                Dim claimsWithoutEvidence = pages.
                    Select(
                        Function(p)
                            Dim unsupportedCount = CountClaimsWithoutEvidence(p, schema)
                            Return Tuple.Create(p.RelativePath, unsupportedCount)
                        End Function).
                    Where(Function(x) x.Item2 > 0).
                    Select(Function(x) $"{x.Item1} — Unsupported claims={x.Item2}").
                    ToList()

                Dim claimsWithoutSources = pages.
                    Select(
                        Function(p)
                            Dim unsourcedCount = CountClaimsWithoutSources(p, schema)
                            Return Tuple.Create(p.RelativePath, unsourcedCount)
                        End Function).
                    Where(Function(x) x.Item2 > 0).
                    Select(Function(x) $"{x.Item1} — Unsourced claims={x.Item2}").
                    ToList()

                Dim skippedRepairItems = LoadRecentRiskyRepairSkips(kbRootPath)

                Dim reviewQueueDocument = BuildReviewQueueDocument(
                    pages:=pages,
                    contradictionPairs:=contradictionPairs,
                    claimsWithoutEvidence:=claimsWithoutEvidence,
                    claimsWithoutSources:=claimsWithoutSources,
                    supersededCandidates:=supersededCandidates,
                    skippedRepairItems:=skippedRepairItems)

                Dim reviewQueuePath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, "review_queue.md")
                File.WriteAllText(reviewQueuePath, reviewQueueDocument, Encoding.UTF8)
            Catch ex As Exception
                LogWikiError(kbRootPath, "review_queue", ex.Message)
            End Try
        End Sub

        Private Shared Function IsRiskyAutoApplyRewrite(page As WikiPageInfo,
                                                        newState As WikiComputedState) As Boolean
            If page Is Nothing OrElse newState Is Nothing Then Return False

            If page.Status.Equals("stable", StringComparison.OrdinalIgnoreCase) AndAlso
               newState.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            If newState.SourceCount < page.SourceCount Then
                Return True
            End If

            If newState.ContradictionCount > page.ContradictionCount Then
                Return True
            End If

            Return False
        End Function

        Private Shared Function BuildMergePrompt() As String
            Return BuildMergePrompt(Nothing)
        End Function

        Private Shared Function BuildMergePrompt(schema As KnowledgeStoreSchema) As String
            Dim sb As New StringBuilder()
            Dim kindsList = FormatKindList(GetEffectivePageKinds(schema))

            If schema IsNot Nothing Then sb.Append(schema.ResolveLanguageDirective())
            sb.AppendLine("You are an autonomous Wiki Maintainer.")
            sb.AppendLine("Merge the existing page and the new extracted content into one improved page.")
            sb.AppendLine("Preserve valuable existing structure, integrate new findings, keep or improve cross-links, and explicitly retain contradictions when evidence conflicts.")
            sb.AppendLine("Never silently replace an older claim with a newer one when they materially conflict.")
            sb.AppendLine()
            sb.AppendLine("SOURCE CITATION RULES:")
            sb.AppendLine("- Refer to sources only by ID markers like [S1], [S2] embedded in Key Claims and Evidence bullets.")
            sb.AppendLine("- DO NOT author the ## Sources section; it is generated deterministically.")
            sb.AppendLine("- DO NOT write file paths, file:// URLs, or download links anywhere.")
            If schema?.PerClaimCitations IsNot Nothing AndAlso schema.PerClaimCitations.Enabled Then
                sb.AppendLine("- Every Key Claim and every Evidence bullet MUST end with at least one [S#] marker referring to the source it relies on.")
            End If
            sb.AppendLine()
            sb.AppendLine("Return exactly:")
            sb.AppendLine("TITLE: ...")
            sb.AppendLine("SUMMARY: ...")
            sb.AppendLine($"KIND: {kindsList}")
            sb.AppendLine()
            sb.AppendLine("<markdown body>")
            sb.AppendLine()
            sb.AppendLine(BuildSchemaSectionsGuidance(schema))

            Return sb.ToString().Trim()
        End Function

        Private Shared Function BuildMergeUserPrompt(existingContent As String,
                                                     newContent As String,
                                                     relatedCandidates As List(Of WikiPageInfo),
                                                     kbRootPath As String,
                                                     currentPagePath As String,
                                                     Optional schema As KnowledgeStoreSchema = Nothing) As String
            Dim sb As New StringBuilder()

            sb.AppendLine("=== EXISTING PAGE ===")
            sb.AppendLine(If(existingContent, "").Trim())
            sb.AppendLine()
            sb.AppendLine("=== NEW CONTENT ===")
            sb.AppendLine(If(newContent, "").Trim())
            sb.AppendLine()

            Dim claimsHeading = GetClaimsHeading(schema)
            Dim contradictionsHeading = GetContradictionsHeading(schema)

            Dim strongestCandidate As WikiPageInfo = Nothing
            Dim strongestExistingClaim As String = ""
            Dim strongestCandidateClaim As String = ""
            Dim strongestScore As Integer = -1

            Dim existingClaims = ExtractSectionBulletItems(existingContent, claimsHeading).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            If existingClaims.Count = 0 Then
                existingClaims.AddRange(
                    ExtractSectionBulletItems(newContent, claimsHeading).
                        Where(Function(x) IsMeaningfulSectionItem(x)).
                        Take(5))
            End If

            If relatedCandidates IsNot Nothing AndAlso relatedCandidates.Count > 0 Then
                sb.AppendLine("=== RELATED EXISTING PAGES ===")
                For Each candidate In relatedCandidates.Take(6)
                    sb.AppendLine($"TITLE: {candidate.Title}")
                    sb.AppendLine($"PATH: {candidate.RelativePath}")
                    sb.AppendLine($"KIND: {candidate.Kind}")
                    sb.AppendLine($"STATUS: {candidate.Status}")
                    If Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        sb.AppendLine($"SUMMARY: {candidate.Summary}")
                    End If
                    sb.AppendLine("EXCERPT:")
                    sb.AppendLine(TruncateForPrompt(RemoveFrontMatter(candidate.Content), 1200))
                    sb.AppendLine()

                    Dim candidateClaims = ExtractSectionBulletItems(candidate.Content, claimsHeading).
                        Where(Function(x) IsMeaningfulSectionItem(x)).
                        ToList()

                    If candidateClaims.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        candidateClaims.Add(candidate.Summary)
                    End If

                    For Each existingClaim In existingClaims.Take(5)
                        For Each candidateClaim In candidateClaims.Take(5)
                            If AreClaimsPotentiallyContradictory(existingClaim, candidateClaim) Then
                                Dim score = claimTokens(existingClaim:=existingClaim, candidateClaim:=candidateClaim)
                                If score > strongestScore Then
                                    strongestScore = score
                                    strongestCandidate = candidate
                                    strongestExistingClaim = existingClaim
                                    strongestCandidateClaim = candidateClaim
                                End If
                            End If
                        Next
                    Next
                Next
            End If

            If strongestCandidate IsNot Nothing Then
                sb.AppendLine("=== POTENTIAL CONTRADICTION CONTEXT ===")
                sb.AppendLine($"CURRENT PAGE: {GetRelativeWikiPath(kbRootPath, currentPagePath)}")
                sb.AppendLine($"CONFLICTING PAGE: {strongestCandidate.RelativePath}")
                sb.AppendLine($"CURRENT CLAIM: {strongestExistingClaim}")
                sb.AppendLine($"CONFLICTING CLAIM: {strongestCandidateClaim}")
                sb.AppendLine("CONFLICTING PAGE EXCERPT:")
                sb.AppendLine(TruncateForPrompt(RemoveFrontMatter(strongestCandidate.Content), 1600))
                sb.AppendLine()
                sb.AppendLine($"Instruction: if the conflict is real and unresolved, preserve both positions explicitly under {contradictionsHeading}.")
            End If

            Return sb.ToString().Trim()
        End Function


        Private Shared Function claimTokens(existingClaim As String, candidateClaim As String) As Integer
            Return Tokenize(existingClaim).
                Intersect(Tokenize(candidateClaim), StringComparer.OrdinalIgnoreCase).
                Count()
        End Function

        Private Shared Function TruncateForPrompt(text As String, maxChars As Integer) As String
            Dim value = If(text, "").Trim()
            If value.Length <= maxChars Then Return value
            Return value.Substring(0, Math.Max(0, maxChars)).Trim() & vbCrLf & "... [truncated]"
        End Function

        Private Shared Function ContainsAny(text As String, values As IEnumerable(Of String)) As Boolean
            Dim haystack = If(text, "")

            For Each value In values
                If haystack.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next

            Return False
        End Function

        Private Shared Function AreClaimsPotentiallyContradictory(claimA As String, claimB As String) As Boolean
            If String.IsNullOrWhiteSpace(claimA) OrElse String.IsNullOrWhiteSpace(claimB) Then
                Return False
            End If

            Dim tokensA = Tokenize(claimA)
            Dim tokensB = Tokenize(claimB)
            Dim sharedCount = tokensA.Intersect(tokensB, StringComparer.OrdinalIgnoreCase).Count()

            If sharedCount < 2 Then
                Return False
            End If

            Dim hasNegationA = Regex.IsMatch(
                claimA,
                "\b(no|not|never|without|none|lack|lacks|lacked|failed|fails|fail|cannot|can't|isn't|aren't|wasn't|weren't)\b",
                RegexOptions.IgnoreCase)

            Dim hasNegationB = Regex.IsMatch(
                claimB,
                "\b(no|not|never|without|none|lack|lacks|lacked|failed|fails|fail|cannot|can't|isn't|aren't|wasn't|weren't)\b",
                RegexOptions.IgnoreCase)

            If hasNegationA <> hasNegationB Then
                Return True
            End If

            Dim positiveMarkers = New String() {
                "increase", "increased", "higher", "growth", "improved", "supports", "supported", "effective", "confirmed", "beneficial"
            }

            Dim negativeMarkers = New String() {
                "decrease", "decreased", "lower", "decline", "worse", "refutes", "refuted", "ineffective", "denied", "harmful"
            }

            If (ContainsAny(claimA, positiveMarkers) AndAlso ContainsAny(claimB, negativeMarkers)) OrElse
               (ContainsAny(claimA, negativeMarkers) AndAlso ContainsAny(claimB, positiveMarkers)) Then
                Return True
            End If

            Return False
        End Function


        Private Class CandidatePagePair
            Public Property PageA As WikiPageInfo = Nothing
            Public Property PageB As WikiPageInfo = Nothing
        End Class

        ''' <summary>
        ''' Generates the candidate (A, B) page pairs that the contradiction and
        ''' superseded-page scanners should consider. When the schema enables
        ''' <see cref="KnowledgeStoreSchema.UseInvertedIndexPairScan"/>, an
        ''' inverted-index pre-filter prunes pairs whose shared-token count is
        ''' below the configured minimum; otherwise the legacy O(N²) enumeration
        ''' is used so behavior is identical to the historical code path.
        ''' Each unordered pair is yielded at most once.
        ''' </summary>
        Private Shared Iterator Function EnumerateCandidatePairs(
                pages As List(Of WikiPageInfo),
                schema As KnowledgeStoreSchema) As IEnumerable(Of CandidatePagePair)

            If pages Is Nothing OrElse pages.Count < 2 Then Return

            If schema Is Nothing OrElse Not schema.UseInvertedIndexPairScan Then
                For i As Integer = 0 To pages.Count - 2
                    For j As Integer = i + 1 To pages.Count - 1
                        Yield New CandidatePagePair With {.PageA = pages(i), .PageB = pages(j)}
                    Next
                Next
                Return
            End If

            Dim minShared As Integer = Math.Max(1, schema.PairScanMinSharedTokens)
            Dim postings As New Dictionary(Of String, List(Of Integer))(StringComparer.OrdinalIgnoreCase)
            Dim pageTokens(pages.Count - 1) As HashSet(Of String)

            For i As Integer = 0 To pages.Count - 1
                Dim toks = Tokenize(
                    pages(i).Title & " " &
                    pages(i).Summary & " " &
                    RemoveFrontMatter(pages(i).Content))
                pageTokens(i) = toks
                For Each t In toks
                    Dim list As List(Of Integer) = Nothing
                    If Not postings.TryGetValue(t, list) Then
                        list = New List(Of Integer)()
                        postings(t) = list
                    End If
                    list.Add(i)
                Next
            Next

            ' Skip extremely common tokens to avoid quadratic blowups in the
            ' candidate count. The cap scales with the wiki size.
            Dim postingCap As Integer = Math.Max(50, pages.Count \ 4)

            Dim emitted As New HashSet(Of Long)()
            For i As Integer = 0 To pages.Count - 1
                Dim counts As New Dictionary(Of Integer, Integer)()

                For Each t In pageTokens(i)
                    Dim list As List(Of Integer) = Nothing
                    If Not postings.TryGetValue(t, list) Then Continue For
                    If list.Count > postingCap Then Continue For

                    For Each j In list
                        If j <= i Then Continue For
                        Dim c As Integer = 0
                        counts.TryGetValue(j, c)
                        counts(j) = c + 1
                    Next
                Next

                For Each kv In counts
                    If kv.Value < minShared Then Continue For
                    Dim key As Long = (CLng(i) << 32) Or CLng(kv.Key)
                    If emitted.Add(key) Then
                        Yield New CandidatePagePair With {.PageA = pages(i), .PageB = pages(kv.Key)}
                    End If
                Next
            Next
        End Function

        Private Shared Sub AddRankedCandidate(
                byPage As Dictionary(Of String, List(Of Tuple(Of WikiPageInfo, Double))),
                fromPage As WikiPageInfo,
                toPage As WikiPageInfo,
                score As Double)
            Dim list As List(Of Tuple(Of WikiPageInfo, Double)) = Nothing
            If Not byPage.TryGetValue(fromPage.FilePath, list) Then
                list = New List(Of Tuple(Of WikiPageInfo, Double))()
                byPage(fromPage.FilePath) = list
            End If
            list.Add(Tuple.Create(toPage, score))
        End Sub


        Private Shared Function FindPotentialContradictionPairs(pages As List(Of WikiPageInfo)) As List(Of PotentialContradictionPair)
            Return FindPotentialContradictionPairs(pages, Nothing)
        End Function

        Private Shared Function FindPotentialContradictionPairs(pages As List(Of WikiPageInfo),
                                                                schema As KnowledgeStoreSchema) As List(Of PotentialContradictionPair)
            Dim result As New List(Of PotentialContradictionPair)()
            If pages Is Nothing OrElse pages.Count = 0 Then Return result
            If schema IsNot Nothing AndAlso Not schema.DetectContradictions Then Return result

            Dim claimsHeading = GetClaimsHeading(schema)
            Dim seenKeys As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' Pre-compute candidate neighbors per page using the (optionally
            ' inverted-index) pair generator. Behavior is identical to the
            ' legacy code path when UseInvertedIndexPairScan is False.
            Dim neighborRanking As New Dictionary(Of String, List(Of Tuple(Of WikiPageInfo, Double)))(StringComparer.OrdinalIgnoreCase)

            For Each pair As CandidatePagePair In EnumerateCandidatePairs(pages, schema)
                Dim pageA As WikiPageInfo = pair.PageA
                Dim pageB As WikiPageInfo = pair.PageB

                Dim seedA = pageA.Title & " " & pageA.Summary & " " & RemoveFrontMatter(pageA.Content)
                Dim seedB = pageB.Title & " " & pageB.Summary & " " & RemoveFrontMatter(pageB.Content)
                Dim score = ScoreTokens(Tokenize(seedA), Tokenize(seedB))
                If score <= 0 Then Continue For

                AddRankedCandidate(neighborRanking, pageA, pageB, score)
                AddRankedCandidate(neighborRanking, pageB, pageA, score)
            Next

            For Each page In pages
                Dim pageClaims = ExtractSectionBulletItems(page.Content, claimsHeading).
                    Where(Function(x) IsMeaningfulSectionItem(x)).
                    ToList()

                If pageClaims.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(page.Summary) Then
                    pageClaims.Add(page.Summary)
                End If

                If pageClaims.Count = 0 Then
                    Continue For
                End If

                Dim ranked As List(Of Tuple(Of WikiPageInfo, Double)) = Nothing
                Dim candidates As List(Of WikiPageInfo)

                If neighborRanking.TryGetValue(page.FilePath, ranked) Then
                    candidates = ranked.
                        OrderByDescending(Function(x) x.Item2).
                        Take(6).
                        Select(Function(x) x.Item1).
                        ToList()
                Else
                    candidates = New List(Of WikiPageInfo)()
                End If

                For Each candidate In candidates
                    Dim pairKey = String.Join(
                        "|",
                        New String() {page.FilePath, candidate.FilePath}.
                            OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase))

                    If seenKeys.Contains(pairKey) Then
                        Continue For
                    End If

                    Dim candidateClaims = ExtractSectionBulletItems(candidate.Content, claimsHeading).
                        Where(Function(x) IsMeaningfulSectionItem(x)).
                        ToList()

                    If candidateClaims.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(candidate.Summary) Then
                        candidateClaims.Add(candidate.Summary)
                    End If

                    Dim foundMatch As Boolean = False

                    For Each claimA In pageClaims.Take(5)
                        For Each claimB In candidateClaims.Take(5)
                            If AreClaimsPotentiallyContradictory(claimA, claimB) Then
                                result.Add(New PotentialContradictionPair With {
                                    .PageA = page,
                                    .PageB = candidate,
                                    .ClaimA = claimA,
                                    .ClaimB = claimB,
                                    .Reason = "Overlapping claim language with opposing polarity."
                                })
                                seenKeys.Add(pairKey)
                                foundMatch = True
                                Exit For
                            End If
                        Next

                        If foundMatch Then
                            Exit For
                        End If
                    Next
                Next
            Next

            Return result.Take(20).ToList()
        End Function

        Private Shared Function GetFrontMatterBoolean(frontMatter As String,
                                                      key As String,
                                                      Optional defaultValue As Boolean = False) As Boolean
            Dim value = GetFrontMatterScalar(frontMatter, key)
            If String.IsNullOrWhiteSpace(value) Then Return defaultValue

            If value.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return True
            If value.Equals("yes", StringComparison.OrdinalIgnoreCase) Then Return True
            If value.Equals("1", StringComparison.OrdinalIgnoreCase) Then Return True
            If value.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return False
            If value.Equals("no", StringComparison.OrdinalIgnoreCase) Then Return False
            If value.Equals("0", StringComparison.OrdinalIgnoreCase) Then Return False

            Return defaultValue
        End Function

        Private Shared Function GetFrontMatterInteger(frontMatter As String,
                                                      key As String,
                                                      Optional defaultValue As Integer = 0) As Integer
            Dim value = GetFrontMatterScalar(frontMatter, key)
            Dim parsed As Integer

            If Integer.TryParse(value, parsed) Then
                Return parsed
            End If

            Return defaultValue
        End Function

        Private Shared Function GetFrontMatterDateTime(frontMatter As String,
                                                       key As String) As DateTime?
            Dim value = GetFrontMatterScalar(frontMatter, key)
            If String.IsNullOrWhiteSpace(value) Then Return Nothing

            Dim parsed As DateTime
            If DateTime.TryParse(
                value,
                Globalization.CultureInfo.InvariantCulture,
                Globalization.DateTimeStyles.RoundtripKind Or Globalization.DateTimeStyles.AllowWhiteSpaces,
                parsed) Then
                Return parsed
            End If

            If DateTime.TryParse(value, parsed) Then
                Return parsed
            End If

            Return Nothing
        End Function

        Private Shared Function ComposePageDocument(title As String,
                                                    summary As String,
                                                    kind As String,
                                                    body As String,
                                                    sourceFilePath As String,
                                                    Optional additionalSourcePaths As IEnumerable(Of String) = Nothing,
                                                    Optional status As String = "stable",
                                                    Optional reviewNeeded As Boolean = False,
                                                    Optional contradictionCount As Integer = 0,
                                                    Optional sourceCount As Integer = 0) As String
            Dim sb As New StringBuilder()
            Dim allSourcePaths = MergeSourcePaths(sourceFilePath, additionalSourcePaths)

            sb.AppendLine("---")
            sb.AppendLine($"title: ""{EscapeYaml(title)}""")
            sb.AppendLine($"summary: ""{EscapeYaml(summary)}""")
            sb.AppendLine($"kind: ""{EscapeYaml(kind)}""")
            sb.AppendLine($"status: ""{EscapeYaml(NormalizePageStatus(status))}""")
            sb.AppendLine($"review_needed: {If(reviewNeeded, "true", "false")}")
            sb.AppendLine($"contradiction_count: {Math.Max(0, contradictionCount)}")
            sb.AppendLine($"source_count: {Math.Max(allSourcePaths.Count, sourceCount)}")
            sb.AppendLine($"updated_utc: ""{DateTime.UtcNow.ToString("o")}""")
            sb.AppendLine("source_paths:")
            For Each path In allSourcePaths
                sb.AppendLine($"  - '{EscapeYamlSingleQuote(path)}'")
            Next

            sb.AppendLine("source_refs:")
            For Each path In allSourcePaths
                Dim slug = SanitizeFileName(System.IO.Path.GetFileNameWithoutExtension(path))
                sb.AppendLine($"  - '{EscapeYamlSingleQuote(slug)}'")
            Next

            sb.AppendLine("---")
            sb.AppendLine()
            sb.AppendLine(If(body, "").Trim())

            Return sb.ToString().Trim() & vbCrLf
        End Function

        Private Shared Function EscapeYamlSingleQuote(value As String) As String
            ' In YAML single-quoted strings, the only escape is '' for a literal single quote.
            Return If(value, "").Replace("'", "''")
        End Function


        Private Shared Function EnsureSourcesSection(body As String,
                                                     sourceFilePath As String,
                                                     Optional additionalSourcePaths As IEnumerable(Of String) = Nothing) As String
            Dim allSourcePaths As New List(Of String)()

            If Not String.IsNullOrWhiteSpace(sourceFilePath) Then
                allSourcePaths.Add(sourceFilePath)
            End If

            If additionalSourcePaths IsNot Nothing Then
                allSourcePaths.AddRange(
                    additionalSourcePaths.
                        Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                        Select(Function(p) p.Trim()))
            End If

            allSourcePaths = allSourcePaths.
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()

            If allSourcePaths.Count = 0 Then Return body
            If body.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) >= 0 Then Return body

            Dim sb As New StringBuilder()
            sb.AppendLine(body.Trim())
            sb.AppendLine()
            sb.AppendLine("## Sources")
            sb.AppendLine()

            For Each path In allSourcePaths
                Dim fileName = System.IO.Path.GetFileName(path)
                If Not String.IsNullOrWhiteSpace(fileName) Then
                    sb.AppendLine($"- [{fileName}]({fileName})")
                End If
            Next

            Return sb.ToString().Trim()
        End Function


        Private Shared Function EnsureDossiersForPage(kbRootPath As String,
                                              pageFullPath As String,
                                              pageTitle As String,
                                              sourcePaths As IEnumerable(Of String),
                                              schema As KnowledgeStoreSchema) As List(Of KnowledgeSourceDossierService.DossierRef)
            Dim refs As New List(Of KnowledgeSourceDossierService.DossierRef)()
            If sourcePaths Is Nothing Then Return refs

            Dim wikiRoot = GetWikiRootPath(kbRootPath)
            Dim pageRelFromWiki = NormalizeWikiRelativePath(GetRelativePathCompat(wikiRoot, pageFullPath))

            For Each sp In sourcePaths.
                    Where(Function(p) Not String.IsNullOrWhiteSpace(p)).
                    Distinct(StringComparer.OrdinalIgnoreCase)
                Dim ref = KnowledgeSourceDossierService.EnsureDossier(kbRootPath, sp, schema)
                If ref Is Nothing Then Continue For
                refs.Add(ref)
                Try
                    KnowledgeSourceDossierService.AddCitedBy(kbRootPath, ref, pageTitle, pageRelFromWiki)
                Catch ex As Exception
                    LogWikiError(kbRootPath, sp, "AddCitedBy failed: " & ex.Message)
                End Try
            Next
            Return refs
        End Function

        ''' <summary>
        ''' Removes blobs from Wiki/_files/ that no .md page references, plus any
        ''' legacy non-md files that linger from the old per-page-copy layout.
        ''' </summary>
        Private Shared Sub CleanupUnusedSourceCopies(kbRootPath As String)
            Try
                Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)

                Dim filesFolderName As String =
            If(schema IsNot Nothing AndAlso schema.SourceAccess IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(schema.SourceAccess.FilesFolderName),
               schema.SourceAccess.FilesFolderName, "_files")

                Dim sourcesFolderName As String =
            If(schema IsNot Nothing AndAlso schema.SourceRegistry IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(schema.SourceRegistry.FolderName),
               schema.SourceRegistry.FolderName, "Sources")

                Dim wikiRoot = GetWikiRootPath(kbRootPath)
                If String.IsNullOrWhiteSpace(wikiRoot) OrElse Not Directory.Exists(wikiRoot) Then Return

                Dim filesRoot = Path.Combine(wikiRoot, filesFolderName)
                Dim sourcesRoot = Path.Combine(wikiRoot, sourcesFolderName)

                ' 1) Collect blob filenames referenced by ANY markdown link or front-matter
                '    pointing into the configured files folder. The regex tolerates "../"
                '    prefixes from dossier-relative links.
                Dim referenced As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim escapedFolder = Regex.Escape(filesFolderName)
                Dim blobPattern = "(?:^|[\s(""'/\\])" & escapedFolder & "/([^\s)""'\r\n]+)"

                For Each md In Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories)
                    Dim text As String
                    Try
                        text = File.ReadAllText(md, Encoding.UTF8)
                    Catch
                        Continue For
                    End Try

                    For Each m As Match In Regex.Matches(text, blobPattern)
                        Dim name = m.Groups(1).Value.Trim()
                        ' Strip any URL fragment / query suffix.
                        Dim cut = name.IndexOfAny(New Char() {"#"c, "?"c})
                        If cut >= 0 Then name = name.Substring(0, cut)
                        If Not String.IsNullOrWhiteSpace(name) Then referenced.Add(name)
                    Next
                Next

                ' 2) Delete blobs nobody references.
                If Directory.Exists(filesRoot) Then
                    For Each blob In Directory.GetFiles(filesRoot)
                        If Not referenced.Contains(Path.GetFileName(blob)) Then
                            Try : File.Delete(blob) : Catch : End Try
                        End If
                    Next
                End If

                ' 3) Delete legacy per-page source copies: any non-.md file in the wiki
                '    tree that does NOT live under the shared files folder. Skip the
                '    Sources dossier folder (only .md belongs there anyway).
                For Each f In Directory.GetFiles(wikiRoot, "*.*", SearchOption.AllDirectories)
                    If Not String.IsNullOrWhiteSpace(filesRoot) AndAlso
               f.StartsWith(filesRoot, StringComparison.OrdinalIgnoreCase) Then Continue For

                    Dim ext = Path.GetExtension(f)
                    If ext.Equals(".md", StringComparison.OrdinalIgnoreCase) Then Continue For

                    Try : File.Delete(f) : Catch : End Try
                Next
            Catch ex As Exception
                LogWikiError(kbRootPath, "CleanupUnusedSourceCopies", ex.Message)
            End Try
        End Sub

        Private Shared Function EnsureRelatedSection(body As String,
                                                     currentPagePath As String,
                                                     relatedCandidates As List(Of WikiPageInfo),
                                                     Optional schema As KnowledgeStoreSchema = Nothing) As String
            If relatedCandidates Is Nothing OrElse relatedCandidates.Count = 0 Then Return body

            Dim relatedPagesHeading As String = "## Related Pages"
            If schema IsNot Nothing Then
                relatedPagesHeading = schema.GetLabel("RelatedPagesHeading", relatedPagesHeading)
            End If

            Dim mergedItems As New List(Of String)()

            For Each candidate In relatedCandidates
                Dim relatedItem = BuildRelatedPageSectionItem(currentPagePath, candidate)
                If Not String.IsNullOrWhiteSpace(relatedItem) Then
                    mergedItems.Add(relatedItem)
                End If
            Next

            Dim existingItems = ExtractSectionBulletItems(body, relatedPagesHeading).
                Select(Function(x) Regex.Replace(If(x, "").Trim(), "\s+", " ").Trim()).
                Where(Function(x) IsMeaningfulSectionItem(x)).
                ToList()

            mergedItems.AddRange(existingItems)

            mergedItems = mergedItems.
                Distinct(StringComparer.OrdinalIgnoreCase).
                Take(6).
                ToList()

            If mergedItems.Count = 0 Then Return body

            Dim sb As New StringBuilder()
            sb.AppendLine(relatedPagesHeading)
            sb.AppendLine()

            For Each item In mergedItems
                sb.AppendLine("- " & item)
            Next

            Return ReplaceMarkdownSection(body, relatedPagesHeading, sb.ToString().Trim())
        End Function


        Private Shared Function BuildRelatedPageSectionItem(currentPagePath As String,
                                                            candidate As WikiPageInfo) As String
            If candidate Is Nothing Then Return ""
            If String.IsNullOrWhiteSpace(candidate.FilePath) Then Return ""
            If candidate.FilePath.Equals(currentPagePath, StringComparison.OrdinalIgnoreCase) Then Return ""

            Dim currentDir = Path.GetDirectoryName(currentPagePath)
            If String.IsNullOrWhiteSpace(currentDir) Then Return ""

            Dim relativeLink = NormalizeWikiRelativePath(
                GetRelativePathCompat(currentDir, candidate.FilePath))

            If String.IsNullOrWhiteSpace(relativeLink) Then Return ""
            If String.IsNullOrWhiteSpace(candidate.Title) Then Return ""

            Return $"[{candidate.Title}]({relativeLink})"
        End Function


        Private Shared Function FindExistingWikiPagePath(kbRootPath As String, title As String) As String
            Dim normalizedTitle = NormalizeWikiKey(title)

            ' Dossier pages are managed by KnowledgeSourceDossierService and must
            ' never be treated as the existing page for a Concept/Analysis/etc. ingest,
            ' or freshly-proposed pages would be merged on top of the dossier file.
            Dim pages = GetAllWikiPages(kbRootPath).
                Where(Function(p) Not p.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase)).
                ToList()

            Dim exact = pages.FirstOrDefault(
                Function(p) NormalizeWikiKey(p.Title) = normalizedTitle OrElse NormalizeWikiKey(Path.GetFileNameWithoutExtension(p.FileName)) = normalizedTitle)

            If exact IsNot Nothing Then Return exact.FilePath

            Dim titleTokens = Tokenize(title)
            Dim best = pages.Select(
                Function(p)
                    Dim score = Jaccard(titleTokens, Tokenize(p.Title))
                    Return Tuple.Create(p.FilePath, score)
                End Function).
                Where(Function(x) x.Item2 >= 0.8).
                OrderByDescending(Function(x) x.Item2).
                FirstOrDefault()

            If best IsNot Nothing Then Return best.Item1
            Return ""
        End Function

        Private Shared Function GetAllWikiPages(kbRootPath As String) As List(Of WikiPageInfo)
            Dim scope = _currentIngestScope.Value
            If scope IsNot Nothing AndAlso
               String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) AndAlso
               scope.CachedPages IsNot Nothing Then
                Return scope.CachedPages
            End If

            Dim pages As New List(Of WikiPageInfo)()
            Dim wikiRoot = GetWikiRootPath(kbRootPath)
            If Not Directory.Exists(wikiRoot) Then Return pages

            Dim schema = GetCachedSchema(kbRootPath)

            For Each file In Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories)
                Dim name = Path.GetFileName(file)
                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("review_queue.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    pages.Add(ReadWikiPageInfo(kbRootPath, file, schema))
                Catch
                End Try
            Next

            If scope IsNot Nothing AndAlso
               String.Equals(scope.KbRootPath, kbRootPath, StringComparison.OrdinalIgnoreCase) Then
                scope.CachedPages = pages
            End If

            Return pages
        End Function

        Private Shared Function ReadWikiPageInfo(kbRootPath As String, filePath As String,
                                                 Optional schema As KnowledgeStoreSchema = Nothing) As WikiPageInfo
            Dim content = File.ReadAllText(filePath, Encoding.UTF8)
            Dim info As New WikiPageInfo() With {
                .FilePath = filePath,
                .FileName = Path.GetFileName(filePath),
                .RelativePath = GetRelativeWikiPath(kbRootPath, filePath),
                .Content = content
            }

            Dim frontMatter = GetFrontMatter(content)
            Dim rawStatus = GetFrontMatterScalar(frontMatter, "status")

            If info.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase) Then
                info.Status = "stable"
                info.ReviewNeeded = False
                info.ContradictionCount = 0
            End If

            info.Title = GetFrontMatterScalar(frontMatter, "title")
            info.Summary = GetFrontMatterScalar(frontMatter, "summary")
            info.Kind = GetFrontMatterScalar(frontMatter, "kind")
            info.ReviewNeeded = GetFrontMatterBoolean(frontMatter, "review_needed", False)
            info.ContradictionCount = GetFrontMatterInteger(frontMatter, "contradiction_count", 0)
            info.SourceCount = GetFrontMatterInteger(frontMatter, "source_count", 0)
            info.UpdatedUtc = GetFrontMatterDateTime(frontMatter, "updated_utc")
            info.SourcePaths = GetFrontMatterList(frontMatter, "source_paths")

            If String.IsNullOrWhiteSpace(info.Title) Then
                info.Title = Path.GetFileNameWithoutExtension(filePath)
            End If

            If String.IsNullOrWhiteSpace(info.Summary) Then
                info.Summary = BuildFallbackSummary(RemoveFrontMatter(content))
            End If

            If String.IsNullOrWhiteSpace(info.Kind) Then
                info.Kind = "Concept"
            End If

            If schema Is Nothing Then schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim computed = ComputePageComputedState(RemoveFrontMatter(content), info.SourcePaths, schema)

            If String.IsNullOrWhiteSpace(rawStatus) Then
                info.Status = computed.Status
            Else
                info.Status = NormalizePageStatus(rawStatus)
            End If

            If info.ContradictionCount <= 0 Then
                info.ContradictionCount = computed.ContradictionCount
            End If

            If info.SourceCount <= 0 Then
                info.SourceCount = computed.SourceCount
            End If

            If Not info.ReviewNeeded AndAlso computed.ReviewNeeded Then
                info.ReviewNeeded = True
            End If

            Return info
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

        Private Shared Function RemoveFrontMatter(content As String) As String
            If String.IsNullOrWhiteSpace(content) Then Return ""
            Dim normalized = content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If Not normalized.StartsWith("---" & vbLf, StringComparison.Ordinal) Then Return content

            Dim endMarker = vbLf & "---" & vbLf
            Dim endPos = normalized.IndexOf(endMarker, 4, StringComparison.Ordinal)
            If endPos < 0 Then Return content

            Return normalized.Substring(endPos + endMarker.Length).Replace(vbLf, vbCrLf)
        End Function

        Private Shared Function GetFrontMatterScalar(frontMatter As String, key As String) As String
            If String.IsNullOrWhiteSpace(frontMatter) Then Return ""

            For Each rawLine In frontMatter.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim line = rawLine.Trim()
                If line.StartsWith(key & ":", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = line.Substring(key.Length + 1).Trim()
                    If value.StartsWith("""") AndAlso value.EndsWith("""") AndAlso value.Length >= 2 Then
                        value = value.Substring(1, value.Length - 2)
                    ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") AndAlso value.Length >= 2 Then
                        value = value.Substring(1, value.Length - 2).Replace("''", "'")
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
                    If trimmed.StartsWith("- ") Then
                        Dim value = trimmed.Substring(2).Trim()
                        If value.StartsWith("""") AndAlso value.EndsWith("""") AndAlso value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2)
                        ElseIf value.StartsWith("'") AndAlso value.EndsWith("'") AndAlso value.Length >= 2 Then
                            value = value.Substring(1, value.Length - 2).Replace("''", "'")
                        End If
                        If Not String.IsNullOrWhiteSpace(value) Then result.Add(value)
                    ElseIf trimmed.Contains(":") Then
                        Exit For
                    End If
                End If
            Next

            Return result
        End Function

        Private Shared Sub RebuildIndex(kbRootPath As String)
            Dim wikiRoot = GetWikiRootPath(kbRootPath)
            If Not Directory.Exists(wikiRoot) Then Return

            Dim schema = GetCachedSchema(kbRootPath)

            Dim pages = GetAllWikiPages(kbRootPath).
                OrderBy(Function(p) p.Kind, StringComparer.OrdinalIgnoreCase).
                ThenBy(Function(p) p.Title, StringComparer.OrdinalIgnoreCase).
                ToList()

            Dim potentialSupersededPagePaths = GetPotentialSupersededPagePaths(pages, schema)

            Dim sb As New StringBuilder()
            sb.AppendLine("# Knowledge Base Index")
            sb.AppendLine()
            sb.AppendLine("> Auto-maintained catalog of wiki pages.")
            sb.AppendLine()

            For Each grp In pages.GroupBy(Function(p) If(String.IsNullOrWhiteSpace(p.Kind), "Concept", p.Kind), StringComparer.OrdinalIgnoreCase)
                sb.AppendLine("## " & grp.Key & " Pages")
                sb.AppendLine()

                For Each page In grp
                    Dim statusBadge = ""
                    If page.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) Then
                        statusBadge = " ⚠ disputed"
                    ElseIf page.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase) Then
                        statusBadge = " ◌ tentative"
                    ElseIf page.Status.Equals("superseded", StringComparison.OrdinalIgnoreCase) Then
                        statusBadge = " ⇢ superseded"
                    End If

                    Dim extra As New List(Of String)()

                    If page.SourceCount > 0 Then
                        extra.Add("sources=" & page.SourceCount.ToString())
                    End If

                    If page.ContradictionCount > 0 Then
                        extra.Add("contradictions=" & page.ContradictionCount.ToString())
                    End If

                    If page.ReviewNeeded Then
                        extra.Add("review")
                    End If

                    If potentialSupersededPagePaths.Contains(page.RelativePath) Then
                        extra.Add("possible-superseded")
                    End If

                    If page.UpdatedUtc.HasValue Then
                        extra.Add("updated=" & page.UpdatedUtc.Value.ToString("yyyy-MM-dd"))
                    End If

                    Dim metaSuffix = ""
                    If extra.Count > 0 Then
                        metaSuffix = " [" & String.Join("; ", extra) & "]"
                    End If

                    sb.AppendLine($"- [{page.Title}]({page.RelativePath}) — {page.Summary}{statusBadge}{metaSuffix}")
                Next

                sb.AppendLine()
            Next

            Dim reviewPages = pages.
                Where(
                    Function(p)
                        Return p.ReviewNeeded OrElse
                               p.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) OrElse
                               p.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase) OrElse
                               potentialSupersededPagePaths.Contains(p.RelativePath)
                    End Function).
                OrderBy(Function(p) p.Title, StringComparer.OrdinalIgnoreCase).
                ToList()

            sb.AppendLine("## Review Queue")
            sb.AppendLine()

            If reviewPages.Count = 0 Then
                sb.AppendLine("- None")
            Else
                For Each page In reviewPages
                    Dim updatedText = If(page.UpdatedUtc.HasValue,
                                         page.UpdatedUtc.Value.ToString("yyyy-MM-dd"),
                                         "unknown")

                    sb.AppendLine(
                        $"- [{page.Title}]({page.RelativePath}) — Status={page.Status}; Contradictions={page.ContradictionCount}; Updated={updatedText}{If(potentialSupersededPagePaths.Contains(page.RelativePath), "; PossibleSuperseded=true", "")}")
                Next
            End If

            sb.AppendLine()

            Dim indexPath = Path.Combine(wikiRoot, KnowledgeStoreCatalog.IndexFile)
            File.WriteAllText(indexPath, sb.ToString().Trim() & vbCrLf, Encoding.UTF8)

            GenerateMkDocsConfig(kbRootPath)
        End Sub

        Private Shared Function BuildFallbackSummary(text As String) As String
            Dim cleaned = Regex.Replace(If(text, ""), "\s+", " ").Trim()
            If cleaned.Length <= 180 Then Return cleaned
            Return cleaned.Substring(0, 180).Trim() & "..."
        End Function

        Private Shared Function EscapeYaml(value As String) As String
            Return If(value, "").Replace("""", "\""")
        End Function

        Private Shared Function BuildFileUri(filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return ""
            Try
                Return New Uri(Path.GetFullPath(filePath)).AbsoluteUri
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim cleaned = New String(If(name, "").Where(Function(c) Not invalid.Contains(c)).ToArray()).Trim()
            If cleaned.Length > 120 Then cleaned = cleaned.Substring(0, 120).Trim()
            If String.IsNullOrWhiteSpace(cleaned) Then cleaned = "unnamed"
            Return cleaned
        End Function

        Private Shared Function NormalizeWikiKey(value As String) As String
            Dim s = Regex.Replace(If(value, "").ToLowerInvariant(), "[^a-z0-9]+", "")
            Return s.Trim()
        End Function

        Private Shared Function Tokenize(text As String) As HashSet(Of String)
            Dim tokens = Regex.Split(If(text, "").ToLowerInvariant(), "[^a-z0-9äöüß]+").
                Where(Function(t) t.Length >= 3).
                ToHashSet(StringComparer.OrdinalIgnoreCase)

            Return tokens
        End Function

        Private Shared Function ScoreTokens(a As HashSet(Of String), b As HashSet(Of String)) As Double
            If a Is Nothing OrElse b Is Nothing OrElse a.Count = 0 OrElse b.Count = 0 Then Return 0
            Return a.Intersect(b, StringComparer.OrdinalIgnoreCase).Count()
        End Function

        Private Shared Function Jaccard(a As HashSet(Of String), b As HashSet(Of String)) As Double
            If a Is Nothing OrElse b Is Nothing OrElse a.Count = 0 OrElse b.Count = 0 Then Return 0
            Dim intersect = a.Intersect(b, StringComparer.OrdinalIgnoreCase).Count()
            Dim union = a.Union(b, StringComparer.OrdinalIgnoreCase).Count()
            If union = 0 Then Return 0
            Return CDbl(intersect) / CDbl(union)
        End Function

        Public Shared Async Function LintWikiAsync(kbRootPath As String,
                                                   context As ISharedContext,
                                                   Optional autoApply As Boolean = False,
                                                   Optional progressCallback As Action(Of String) = Nothing) As Task(Of String)

            If String.IsNullOrWhiteSpace(kbRootPath) Then Return "Invalid KB path."

            Dim pages = GetAllWikiPages(kbRootPath)
            If pages.Count = 0 Then Return "Wiki is empty or not initialized."

            ' Exclude dossier pages from health-checks; they are managed deterministically.
            pages = pages.
                        Where(Function(p) Not p.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase)).
                        ToList()

            Dim pageLookup = pages.ToDictionary(Function(p) p.RelativePath, StringComparer.OrdinalIgnoreCase)
            Dim inboundCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim brokenLinks As New List(Of String)()
            Dim orphanPages As New List(Of String)()
            Dim missingSources As New List(Of String)()
            Dim reviewPages As New List(Of String)()
            Dim claimsWithoutEvidence As New List(Of String)()
            Dim claimsWithoutSources As New List(Of String)()
            Dim contradictedClaims As New List(Of String)()

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)

            Dim operationLabel As String = If(autoApply, "Lint Wiki", "Health Check")
            ReportMaintenanceProgress(progressCallback, $"{operationLabel}: analyzing wiki structure.")

            For Each page In pages
                inboundCounts(page.RelativePath) = 0
            Next

            For Each page In pages
                Dim internalLinks = ExtractInternalMarkdownLinks(page.Content, page.FilePath, kbRootPath)

                For Each link In internalLinks
                    If pageLookup.ContainsKey(link) Then
                        inboundCounts(link) += 1
                    Else
                        brokenLinks.Add($"{page.RelativePath} -> {link}")
                    End If
                Next

                If (schema Is Nothing OrElse schema.AlwaysAddSourceLinks) AndAlso
                   page.Kind.Equals("Source", StringComparison.OrdinalIgnoreCase) AndAlso
                   (page.SourcePaths Is Nothing OrElse page.SourcePaths.Count = 0) AndAlso
                   page.Content.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) < 0 Then
                    missingSources.Add(page.RelativePath)
                End If

                Dim unsupportedClaims = CountClaimsWithoutEvidence(page, schema)
                If unsupportedClaims > 0 Then
                    claimsWithoutEvidence.Add($"{page.RelativePath} — Unsupported claims={unsupportedClaims}")
                End If

                Dim unsourcedClaims = CountClaimsWithoutSources(page, schema)
                If unsourcedClaims > 0 Then
                    claimsWithoutSources.Add($"{page.RelativePath} — Unsourced claims={unsourcedClaims}")
                End If

                Dim contradictedClaimCount = CountContradictedClaims(page, schema)
                If contradictedClaimCount > 0 Then
                    contradictedClaims.Add($"{page.RelativePath} — Contradicted claims={contradictedClaimCount}")
                End If

                If page.ReviewNeeded OrElse
                   page.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase) OrElse
                   page.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase) Then
                    reviewPages.Add($"{page.RelativePath} — Status={page.Status}; Contradictions={page.ContradictionCount}")
                End If
            Next

            For Each page In pages
                If inboundCounts.ContainsKey(page.RelativePath) AndAlso inboundCounts(page.RelativePath) = 0 Then
                    orphanPages.Add(page.RelativePath)
                End If
            Next

            Dim contradictionPairs = FindPotentialContradictionPairs(pages, schema)
            Dim supersededCandidates = FindPotentialSupersededPages(pages)

            Dim report As New StringBuilder()
            report.AppendLine("# Wiki Health Report")
            report.AppendLine()
            report.AppendLine($"**Date Generated:** {DateTime.Now:yyyy-MM-dd HH:mm}")
            report.AppendLine($"**Total Pages:** {pages.Count}")
            report.AppendLine()

            report.AppendLine("## Orphan Pages")
            report.AppendLine()
            If orphanPages.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In orphanPages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Broken Internal Links")
            report.AppendLine()
            If brokenLinks.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In brokenLinks.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Source Pages Missing Source Citations")
            report.AppendLine()
            If missingSources.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In missingSources.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Pages Requiring Review")
            report.AppendLine()
            If reviewPages.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In reviewPages.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Claims With Weak Evidence")
            report.AppendLine()
            If claimsWithoutEvidence.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In claimsWithoutEvidence.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Claims Without Sources")
            report.AppendLine()
            If claimsWithoutSources.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In claimsWithoutSources.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Contradicted Claims")
            report.AppendLine()
            If contradictedClaims.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In contradictedClaims.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Potential Contradictions")
            report.AppendLine()
            If contradictionPairs.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each pair In contradictionPairs
                    report.AppendLine($"- {pair.PageA.RelativePath} <-> {pair.PageB.RelativePath}")
                    report.AppendLine($"  - Claim A: {pair.ClaimA}")
                    report.AppendLine($"  - Claim B: {pair.ClaimB}")
                    report.AppendLine($"  - Reason: {pair.Reason}")
                Next
            End If
            report.AppendLine()

            report.AppendLine("## Possible Superseded Pages")
            report.AppendLine()
            If supersededCandidates.Count = 0 Then
                report.AppendLine("- None")
            Else
                For Each item In supersededCandidates.OrderBy(Function(x) x, StringComparer.OrdinalIgnoreCase)
                    report.AppendLine($"- {item}")
                Next
            End If
            report.AppendLine()

            Dim promptUser As New StringBuilder()

            If schema IsNot Nothing Then
                promptUser.AppendLine(schema.ToPromptBlock())
                promptUser.AppendLine()
            End If

            promptUser.AppendLine("=== CURRENT INDEX ===")
            promptUser.AppendLine(LoadIndexText(kbRootPath))
            promptUser.AppendLine()
            promptUser.AppendLine("=== STRUCTURAL REPORT ===")
            promptUser.AppendLine(report.ToString().Trim())
            promptUser.AppendLine()

            If contradictionPairs.Count > 0 Then
                promptUser.AppendLine("=== SUSPICIOUS PAGE PAIRS ===")
                For Each pair In contradictionPairs.Take(5)
                    promptUser.AppendLine($"PAIR: {pair.PageA.RelativePath} <-> {pair.PageB.RelativePath}")
                    promptUser.AppendLine("PAGE A EXCERPT:")
                    promptUser.AppendLine(TruncateForPrompt(RemoveFrontMatter(pair.PageA.Content), 1600))
                    promptUser.AppendLine()
                    promptUser.AppendLine("PAGE B EXCERPT:")
                    promptUser.AppendLine(TruncateForPrompt(RemoveFrontMatter(pair.PageB.Content), 1600))
                    promptUser.AppendLine()
                Next
            End If

            Dim systemPrompt = "You are a Knowledge Base Health Inspector. Review the structural report and the supplied page excerpts. " &
                               "Identify duplicate concepts, contradictions, stale claims, weakly-supported claims, missing cross-links, worthwhile next questions, and possible superseded pages. " &
                               "Be explicit about which pages should be marked disputed, tentative, superseded, or review_needed."

            ReportMaintenanceProgress(progressCallback, $"{operationLabel}: preparing AI structural analysis.")

            Dim llmAnalysis = Await ExecuteKnowledgeStoreLlmAsync(
                context:=context,
                promptSystem:=systemPrompt,
                promptUser:=promptUser.ToString().Trim(),
                hideSplash:=False,
                progressCallback:=progressCallback,
                operationName:=operationLabel).ConfigureAwait(False)

            If Not String.IsNullOrWhiteSpace(llmAnalysis) Then
                report.AppendLine("## AI Structural Analysis")
                report.AppendLine()
                report.AppendLine(llmAnalysis.Trim())
                report.AppendLine()
            End If

            If autoApply Then
                ReportMaintenanceProgress(progressCallback, $"{operationLabel}: applying automatic fixes.")

                Dim fixSummary = Await ApplyWikiHealthFixesAsync(
                    kbRootPath:=kbRootPath,
                    context:=context,
                    includeLlmRepairs:=True,
                    progressCallback:=progressCallback,
                    operationName:=operationLabel).ConfigureAwait(False)

                If Not String.IsNullOrWhiteSpace(fixSummary) Then
                    report.AppendLine("## Auto-Applied Fixes")
                    report.AppendLine()
                    report.AppendLine(fixSummary.Trim())
                    report.AppendLine()
                End If
            End If

            Dim reportPath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, "health_report.md")
            File.WriteAllText(reportPath, report.ToString().Trim() & vbCrLf, Encoding.UTF8)

            Dim reviewQueuePath = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder, "review_queue.md")
            Dim reviewQueueDocument = BuildReviewQueueDocument(
                pages:=pages,
                contradictionPairs:=contradictionPairs,
                claimsWithoutEvidence:=claimsWithoutEvidence,
                claimsWithoutSources:=claimsWithoutSources,
                supersededCandidates:=supersededCandidates)
            File.WriteAllText(reviewQueuePath, reviewQueueDocument, Encoding.UTF8)

            AppendToLog(
                kbRootPath,
                "lint",
                $"Pages={pages.Count}; Orphans={orphanPages.Count}; BrokenLinks={brokenLinks.Count}; Reviews={reviewPages.Count}; PotentialContradictions={contradictionPairs.Count}; SupersededCandidates={supersededCandidates.Count}; AutoApplied={autoApply}")

            ReportMaintenanceProgress(progressCallback, $"{operationLabel}: finished.")

            Return report.ToString()
        End Function

        Public Shared Async Function ApplyWikiHealthFixesAsync(kbRootPath As String,
                                                               context As ISharedContext,
                                                               Optional includeLlmRepairs As Boolean = True,
                                                               Optional progressCallback As Action(Of String) = Nothing,
                                                               Optional operationName As String = "Repair") As Task(Of String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return ""
            InitializeWikiStructure(kbRootPath)

            ' Same scope pattern as IngestSourceAsync: collapse the per-saved-page
            ' RebuildIndex / RefreshReviewArtifacts / BackfillRelatedLinks work
            ' (driven by ParseAndSaveAgentResponseAsync inside the repair loops)
            ' into a single end-of-run flush. The explicit RebuildIndex /
            ' RefreshReviewArtifacts calls already present inside
            ' ApplyWikiHealthFixesCoreAsync remain — they are idempotent and the
            ' final pair simply overlaps harmlessly with EndIngestScopeAsync.
            BeginEmbeddingBatch(kbRootPath)
            BeginIngestScope(kbRootPath)

            Dim resultText As String = ""
            Dim captured As ExceptionDispatchInfo = Nothing

            Try
                resultText = Await ApplyWikiHealthFixesCoreAsync(
                    kbRootPath:=kbRootPath,
                    context:=context,
                    includeLlmRepairs:=includeLlmRepairs,
                    progressCallback:=progressCallback,
                    operationName:=operationName).ConfigureAwait(False)
            Catch ex As Exception
                captured = ExceptionDispatchInfo.Capture(ex)
            End Try

            ' End ingest scope BEFORE the embedding batch (same ordering as
            ' IngestSourceAsync): the scope's deferred backfill rewrites must
            ' still queue their embedding updates into the active batch.
            Await EndIngestScopeAsync(context).ConfigureAwait(False)
            Await EndEmbeddingBatchAsync(context).ConfigureAwait(False)

            If captured IsNot Nothing Then captured.Throw()
            Return resultText
        End Function

        Private Shared Async Function ApplyWikiHealthFixesCoreAsync(kbRootPath As String,
                                                                    context As ISharedContext,
                                                                    includeLlmRepairs As Boolean,
                                                                    progressCallback As Action(Of String),
                                                                    operationName As String) As Task(Of String)

            ' Idempotent: safe to run on every auto-apply pass.
            Dim migrationSchema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            ReportMaintenanceProgress(progressCallback, $"{operationName}: migrating sources to dossiers / shared _files folder.")
            Dim migrated As Integer = 0
            For Each page In GetAllWikiPages(kbRootPath)
                If page.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase) Then Continue For
                If page.SourcePaths Is Nothing OrElse page.SourcePaths.Count = 0 Then Continue For

                Dim refs = EnsureDossiersForPage(kbRootPath, page.FilePath, page.Title, page.SourcePaths, migrationSchema)
                If refs.Count = 0 Then Continue For

                Dim body = RemoveFrontMatter(page.Content).Trim()
                body = RebuildSourcesSectionDeterministically(body, page.FilePath, kbRootPath, refs)

                Dim state = ComputePageComputedState(body, page.SourcePaths, migrationSchema)
                Dim recomposed = ComposePageDocument(
                    title:=page.Title,
                    summary:=page.Summary,
                    kind:=page.Kind,
                    body:=body,
                    sourceFilePath:="",
                    additionalSourcePaths:=page.SourcePaths,
                    status:=state.Status,
                    reviewNeeded:=state.ReviewNeeded,
                    contradictionCount:=state.ContradictionCount,
                    sourceCount:=state.SourceCount)

                If Not String.Equals(page.Content, recomposed, StringComparison.Ordinal) Then
                    File.WriteAllText(page.FilePath, recomposed, Encoding.UTF8)
                    Await QueueOrUpdateEmbeddingAsync(kbRootPath, page.FilePath, recomposed, context).ConfigureAwait(False)
                    migrated += 1
                End If
            Next
            If migrated > 0 Then
                AppendToLog(kbRootPath, "dossier-migration", $"PagesMigrated={migrated}")
            End If

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim pages = GetAllWikiPages(kbRootPath)
            If pages.Count = 0 Then Return "No wiki pages found."

            Dim result As New WikiAutoFixResult()
            Dim skippedRiskyWrites As Integer = 0
            Dim potentialSupersededPagePaths = GetPotentialSupersededPagePaths(pages)

            ReportMaintenanceProgress(progressCallback, $"{operationName}: running deterministic repairs.")

            Dim pageNumber As Integer = 0

            For Each page In pages
                pageNumber += 1

                If page.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                ReportMaintenanceProgress(
                    progressCallback,
                    $"{operationName}: normalizing page {pageNumber} of {pages.Count}: {If(page.Title, page.RelativePath)}")

                Dim originalContent = page.Content

                Dim candidates = FindCandidateWikiPages(
                    kbRootPath,
                    page.Title & " " & page.Summary & " " & RemoveFrontMatter(originalContent),
                    8).
                    Where(Function(p) Not p.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase)).
                    ToList()

                Dim updatedBody = NormalizePageBody(
                    body:=RemoveFrontMatter(originalContent).Trim(),
                    currentPagePath:=page.FilePath,
                    sourcePaths:=page.SourcePaths,
                    relatedCandidates:=candidates,
                    schema:=schema)

                Dim repairedLinksCount As Integer = 0
                updatedBody = RepairResolvableMarkdownLinks(updatedBody, page.FilePath, kbRootPath, repairedLinksCount)
                If repairedLinksCount > 0 Then
                    result.RepairedLinks += repairedLinksCount
                End If

                Dim removedBrokenLinksCount As Integer = 0
                updatedBody = StripUnresolvableInternalLinks(updatedBody, page.FilePath, kbRootPath, removedBrokenLinksCount)
                If removedBrokenLinksCount > 0 Then
                    result.RemovedBrokenLinks += removedBrokenLinksCount
                End If

                If (schema Is Nothing OrElse schema.AlwaysAddSourceLinks) AndAlso
                   originalContent.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   updatedBody.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    result.AddedSourcesSections += 1
                End If

                If (schema Is Nothing OrElse schema.AlwaysCreateCrossLinks) AndAlso
                   originalContent.IndexOf("## Related Pages", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   updatedBody.IndexOf("## Related Pages", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    result.AddedRelatedSections += 1
                End If

                Dim pageState = ComputePageComputedState(updatedBody, page.SourcePaths, schema)

                If potentialSupersededPagePaths.Contains(page.RelativePath) Then
                    pageState.ReviewNeeded = True
                End If

                If IsRiskyAutoApplyRewrite(page, pageState) Then
                    skippedRiskyWrites += 1
                    AppendToLog(
                        kbRootPath,
                        "repair-skip",
                        $"{page.Title} | Risky deterministic rewrite skipped | OldStatus={page.Status}; NewStatus={pageState.Status}; OldSources={page.SourceCount}; NewSources={pageState.SourceCount}; OldContradictions={page.ContradictionCount}; NewContradictions={pageState.ContradictionCount}")
                    Continue For
                End If

                Dim recomposedContent = ComposePageDocument(
                    title:=page.Title,
                    summary:=page.Summary,
                    kind:=page.Kind,
                    body:=updatedBody,
                    sourceFilePath:="",
                    additionalSourcePaths:=page.SourcePaths,
                    status:=pageState.Status,
                    reviewNeeded:=pageState.ReviewNeeded,
                    contradictionCount:=pageState.ContradictionCount,
                    sourceCount:=pageState.SourceCount)

                If Not String.Equals(originalContent, recomposedContent, StringComparison.Ordinal) Then
                    File.WriteAllText(page.FilePath, recomposedContent, Encoding.UTF8)
                    Await QueueOrUpdateEmbeddingAsync(kbRootPath, page.FilePath, recomposedContent, context).ConfigureAwait(False)
                    result.UpdatedPages += 1
                End If
            Next

            RebuildIndex(kbRootPath)
            RefreshReviewArtifacts(kbRootPath)

            If includeLlmRepairs Then
                ReportMaintenanceProgress(progressCallback, $"{operationName}: running LLM-assisted structural repairs.")

                Dim llmCount = Await ApplyLlmStructuralRepairsAsync(
                    kbRootPath,
                    context,
                    progressCallback,
                    operationName).ConfigureAwait(False)

                result.LlmRepairedPages = llmCount
            End If

            CleanupUnusedSourceCopies(kbRootPath)
            RebuildIndex(kbRootPath)
            RefreshReviewArtifacts(kbRootPath)

            Dim sb As New StringBuilder()
            sb.AppendLine($"- Updated pages: {result.UpdatedPages}")
            sb.AppendLine($"- Added Sources sections: {result.AddedSourcesSections}")
            sb.AppendLine($"- Added Related Pages sections: {result.AddedRelatedSections}")
            sb.AppendLine($"- Repaired markdown links: {result.RepairedLinks}")
            sb.AppendLine($"- Removed broken internal links: {result.RemovedBrokenLinks}")
            sb.AppendLine($"- LLM-repaired pages: {result.LlmRepairedPages}")
            sb.AppendLine($"- Risky writes skipped: {skippedRiskyWrites}")

            AppendToLog(
                kbRootPath,
                "repair",
                $"Updated={result.UpdatedPages}; Sources={result.AddedSourcesSections}; Related={result.AddedRelatedSections}; Links={result.RepairedLinks}; BrokenLinksRemoved={result.RemovedBrokenLinks}; LlmPages={result.LlmRepairedPages}; Skipped={skippedRiskyWrites}")

            ReportMaintenanceProgress(progressCallback, $"{operationName}: finished.")
            Return sb.ToString().Trim()
        End Function



        Private Shared Async Function ApplyLlmStructuralRepairsAsync(kbRootPath As String,
                                                                     context As ISharedContext,
                                                                     Optional progressCallback As Action(Of String) = Nothing,
                                                                     Optional operationName As String = "Repair") As Task(Of Integer)

            Dim schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            Dim pages = GetAllWikiPages(kbRootPath)
            If pages.Count = 0 Then Return 0

            Dim pageLookup = pages.ToDictionary(Function(p) p.RelativePath, StringComparer.OrdinalIgnoreCase)
            Dim inboundCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

            For Each page In pages
                inboundCounts(page.RelativePath) = 0
            Next

            For Each page In pages
                Dim internalLinks = ExtractInternalMarkdownLinks(page.Content, page.FilePath, kbRootPath)
                For Each link In internalLinks
                    If pageLookup.ContainsKey(link) Then
                        inboundCounts(link) += 1
                    End If
                Next
            Next

            Dim contradictionPairs = FindPotentialContradictionPairs(pages, schema)
            Dim problemPages = FindPagesNeedingLlmRepair(pages, inboundCounts, schema)

            If problemPages.Count = 0 Then
                ReportMaintenanceProgress(progressCallback, $"{operationName}: no LLM repairs were needed.")
                Return 0
            End If

            Dim repairedCount As Integer = 0
            Dim problemPageNumber As Integer = 0

            For Each page In problemPages
                problemPageNumber += 1

                ReportMaintenanceProgress(
                    progressCallback,
                    $"{operationName}: LLM repairing page {problemPageNumber} of {problemPages.Count}: {If(page.Title, page.RelativePath)}")

                Dim candidates = FindCandidateWikiPages(
                    kbRootPath,
                    page.Title & " " & page.Summary & " " & RemoveFrontMatter(page.Content),
                    8).
                    Where(Function(p) Not p.FilePath.Equals(page.FilePath, StringComparison.OrdinalIgnoreCase)).
                    ToList()

                Dim promptSystem As New StringBuilder()
                If schema IsNot Nothing Then promptSystem.Append(schema.ResolveLanguageDirective())
                promptSystem.AppendLine("You are an autonomous Wiki Maintainer.")
                promptSystem.AppendLine("Repair the provided wiki page so that it integrates better with the rest of the wiki.")
                promptSystem.AppendLine("Preserve true facts, preserve source links, add or improve related-page links, and keep contradictions where relevant.")
                promptSystem.AppendLine("If evidence conflicts, do not hide the conflict.")
                promptSystem.AppendLine()
                promptSystem.AppendLine(BuildSchemaSectionsGuidance(schema))
                promptSystem.AppendLine()
                promptSystem.AppendLine("Only create clickable markdown page links to pages that already exist in the supplied wiki context.")
                promptSystem.AppendLine("If a useful concept does not yet have a page, mention it as plain text or under ## Open Questions, but do not create a broken page link.")
                promptSystem.AppendLine("Return exactly:")
                promptSystem.AppendLine("TITLE: ...")
                promptSystem.AppendLine("SUMMARY: ...")
                promptSystem.AppendLine($"KIND: {FormatKindList(GetEffectivePageKinds(schema))}")
                promptSystem.AppendLine()
                promptSystem.AppendLine("<markdown body>")

                Dim repairedResponse = Await ExecuteKnowledgeStoreLlmAsync(
                    context:=context,
                    promptSystem:=promptSystem.ToString().Trim(),
                    promptUser:=BuildRepairPromptUser(page, candidates, contradictionPairs, kbRootPath, schema),
                    hideSplash:=True,
                    progressCallback:=progressCallback,
                    operationName:=operationName).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(repairedResponse) Then
                    Continue For
                End If

                Dim saved = Await ParseAndSaveAgentResponseAsync(
                    kbRootPath:=kbRootPath,
                    agentResponse:=repairedResponse,
                    context:=context,
                    sourceFilePath:="",
                    defaultKind:=page.Kind,
                    relatedCandidates:=candidates,
                    actionName:="repair-llm",
                    additionalSourcePaths:=page.SourcePaths,
                    schema:=schema).ConfigureAwait(False)

                If saved Then
                    repairedCount += 1
                End If
            Next

            Return repairedCount
        End Function


        Private Shared Async Function BackfillRelatedLinksForExistingPagesAsync(kbRootPath As String,
                                                                                savedPage As WikiPageInfo,
                                                                                context As ISharedContext,
                                                                                Optional schema As KnowledgeStoreSchema = Nothing) As Task(Of Integer)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return 0
            If savedPage Is Nothing OrElse String.IsNullOrWhiteSpace(savedPage.FilePath) Then Return 0

            If schema Is Nothing Then
                schema = KnowledgeStoreSchema.LoadOrCreate(kbRootPath)
            End If

            If schema IsNot Nothing AndAlso Not schema.AlwaysCreateCrossLinks Then
                Return 0
            End If

            Dim seedText =
                savedPage.Title & " " &
                savedPage.Summary & " " &
                RemoveFrontMatter(savedPage.Content)

            Dim candidatePages = FindCandidateWikiPages(kbRootPath, seedText, 8).
                        Where(Function(p) Not p.FilePath.Equals(savedPage.FilePath, StringComparison.OrdinalIgnoreCase)).
                        Where(Function(p) Not p.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase)).
                        ToList()

            Dim updatedCount As Integer = 0

            For Each candidate In candidatePages
                If candidate.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim latestCandidate As WikiPageInfo = Nothing

                Try
                    latestCandidate = ReadWikiPageInfo(kbRootPath, candidate.FilePath, schema)
                Catch
                    Continue For
                End Try

                Dim originalContent = latestCandidate.Content
                Dim updatedBody = NormalizePageBody(
                    body:=RemoveFrontMatter(originalContent).Trim(),
                    currentPagePath:=latestCandidate.FilePath,
                    sourcePaths:=latestCandidate.SourcePaths,
                    relatedCandidates:=New List(Of WikiPageInfo) From {savedPage},
                    schema:=schema)

                Dim repairedLinksCount As Integer = 0
                updatedBody = RepairResolvableMarkdownLinks(
                    updatedBody,
                    latestCandidate.FilePath,
                    kbRootPath,
                    repairedLinksCount)

                Dim removedBrokenLinksCount As Integer = 0
                updatedBody = StripUnresolvableInternalLinks(
                    updatedBody,
                    latestCandidate.FilePath,
                    kbRootPath,
                    removedBrokenLinksCount)

                Dim pageState = ComputePageComputedState(updatedBody, latestCandidate.SourcePaths, schema)

                Dim recomposedContent = ComposePageDocument(
                    title:=latestCandidate.Title,
                    summary:=latestCandidate.Summary,
                    kind:=latestCandidate.Kind,
                    body:=updatedBody,
                    sourceFilePath:="",
                    additionalSourcePaths:=latestCandidate.SourcePaths,
                    status:=pageState.Status,
                    reviewNeeded:=pageState.ReviewNeeded,
                    contradictionCount:=pageState.ContradictionCount,
                    sourceCount:=pageState.SourceCount)

                If Not String.Equals(originalContent, recomposedContent, StringComparison.Ordinal) Then
                    File.WriteAllText(latestCandidate.FilePath, recomposedContent, Encoding.UTF8)
                    Await QueueOrUpdateEmbeddingAsync(
                        kbRootPath,
                        latestCandidate.FilePath,
                        recomposedContent,
                        context).ConfigureAwait(False)

                    updatedCount += 1
                End If
            Next

            If updatedCount > 0 Then
                AppendToLog(
                    kbRootPath,
                    "related-backfill",
                    $"{savedPage.Title} | UpdatedExistingPages={updatedCount}")
            End If

            Return updatedCount
        End Function


        Private Shared Function FindPagesNeedingLlmRepair(pages As List(Of WikiPageInfo),
                                                          inboundCounts As Dictionary(Of String, Integer),
                                                          Optional schema As KnowledgeStoreSchema = Nothing) As List(Of WikiPageInfo)
            Return pages.
                    Where(
                        Function(p)
                            ' Dossier pages are managed deterministically; never send them to the LLM.
                            If p.Kind.Equals("SourceDossier", StringComparison.OrdinalIgnoreCase) Then
                                Return False
                            End If

                            Dim hasSourcesIssue =
                                (schema Is Nothing OrElse schema.AlwaysAddSourceLinks) AndAlso
                                p.Kind.Equals("Source", StringComparison.OrdinalIgnoreCase) AndAlso
                                (p.SourcePaths Is Nothing OrElse p.SourcePaths.Count = 0) AndAlso
                                p.Content.IndexOf("## Sources", StringComparison.OrdinalIgnoreCase) < 0

                            Dim hasRelatedIssue =
                                (schema Is Nothing OrElse schema.AlwaysCreateCrossLinks) AndAlso
                                p.Content.IndexOf("## Related Pages", StringComparison.OrdinalIgnoreCase) < 0

                            Dim isOrphan = inboundCounts.ContainsKey(p.RelativePath) AndAlso inboundCounts(p.RelativePath) = 0
                            Dim isDisputed = p.Status.Equals("disputed", StringComparison.OrdinalIgnoreCase)
                            Dim isTentative = p.Status.Equals("tentative", StringComparison.OrdinalIgnoreCase)
                            Dim weakEvidence = CountClaimsWithoutEvidence(p, schema) > 0
                            Dim needsReview = p.ReviewNeeded
                            Dim hasContradictions = p.ContradictionCount > 0

                            Return hasSourcesIssue OrElse
                                   hasRelatedIssue OrElse
                                   isOrphan OrElse
                                   isDisputed OrElse
                                   isTentative OrElse
                                   weakEvidence OrElse
                                   needsReview OrElse
                                   hasContradictions
                        End Function).
                    OrderByDescending(Function(p) If(p.ReviewNeeded, 1, 0)).
                    ThenByDescending(Function(p) p.ContradictionCount).
                    Take(12).
                    ToList()
        End Function

        Private Shared Function ExtractInternalMarkdownLinks(content As String,
                                                             currentPagePath As String,
                                                             kbRootPath As String) As List(Of String)
            Dim result As New List(Of String)()
            If String.IsNullOrWhiteSpace(content) Then Return result

            For Each match As Match In Regex.Matches(content, "\[[^\]]+\]\(([^)]+)\)")
                Dim target = match.Groups(1).Value.Trim()
                If target.EndsWith(".md", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not target.StartsWith("https://", StringComparison.OrdinalIgnoreCase) AndAlso
                   Not target.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then

                    Dim resolved = ResolveWikiLinkTarget(currentPagePath, target, kbRootPath)
                    If Not String.IsNullOrWhiteSpace(resolved) Then
                        result.Add(resolved)
                    End If
                End If
            Next

            For Each match As Match In Regex.Matches(content, "\[\[([^\]]+)\]\]")
                Dim target = match.Groups(1).Value.Trim()
                If String.IsNullOrWhiteSpace(target) Then
                    Continue For
                End If

                Dim existingPagePath = FindExistingWikiPagePath(kbRootPath, target)
                If Not String.IsNullOrWhiteSpace(existingPagePath) Then
                    result.Add(GetRelativeWikiPath(kbRootPath, existingPagePath))
                End If
            Next

            Return result.
                Where(Function(x) Not String.IsNullOrWhiteSpace(x)).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        End Function

        Private Shared Function StripUnresolvableInternalLinks(content As String,
                                                               currentPagePath As String,
                                                               kbRootPath As String,
                                                               ByRef removedCount As Integer) As String
            removedCount = 0
            If String.IsNullOrWhiteSpace(content) Then Return content

            Dim localRemovedCount As Integer = 0

            Dim updatedContent = Regex.Replace(
                content,
                "\[([^\]]+)\]\(([^)]+)\)",
                Function(m)
                    Dim linkText = m.Groups(1).Value.Trim()
                    Dim rawTarget = m.Groups(2).Value.Trim()

                    If String.IsNullOrWhiteSpace(rawTarget) Then
                        Return m.Value
                    End If

                    If rawTarget.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                       rawTarget.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
                       rawTarget.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                        Return m.Value
                    End If

                    Dim target = rawTarget
                    Dim fragmentIndex = target.IndexOf("#"c)
                    If fragmentIndex >= 0 Then
                        target = target.Substring(0, fragmentIndex).Trim()
                    End If

                    If Not target.EndsWith(".md", StringComparison.OrdinalIgnoreCase) Then
                        Return m.Value
                    End If

                    Dim resolved = ResolveWikiLinkTarget(currentPagePath, target, kbRootPath)
                    If String.IsNullOrWhiteSpace(resolved) Then
                        localRemovedCount += 1
                        Return If(String.IsNullOrWhiteSpace(linkText), target, linkText)
                    End If

                    Dim absoluteTarget = Path.Combine(GetWikiRootPath(kbRootPath), resolved.Replace("/"c, Path.DirectorySeparatorChar))
                    If Not File.Exists(absoluteTarget) Then
                        localRemovedCount += 1
                        Return If(String.IsNullOrWhiteSpace(linkText), Path.GetFileNameWithoutExtension(target), linkText)
                    End If

                    Return m.Value
                End Function)

            updatedContent = Regex.Replace(
                updatedContent,
                "\[\[([^\]]+)\]\]",
                Function(m)
                    Dim target = m.Groups(1).Value.Trim()
                    If String.IsNullOrWhiteSpace(target) Then
                        Return m.Value
                    End If

                    Dim existingPagePath = FindExistingWikiPagePath(kbRootPath, target)
                    If String.IsNullOrWhiteSpace(existingPagePath) Then
                        localRemovedCount += 1
                        Return target
                    End If

                    Dim relativeLink = NormalizeWikiRelativePath(GetRelativePathCompat(Path.GetDirectoryName(currentPagePath), existingPagePath))
                    If String.IsNullOrWhiteSpace(relativeLink) Then
                        localRemovedCount += 1
                        Return target
                    End If

                    Return $"[{target}]({relativeLink})"
                End Function)

            removedCount = localRemovedCount
            Return updatedContent
        End Function

        Private Shared Function RepairResolvableMarkdownLinks(content As String,
                                                              currentPagePath As String,
                                                              kbRootPath As String,
                                                              ByRef repairedCount As Integer) As String
            repairedCount = 0
            If String.IsNullOrWhiteSpace(content) Then Return content

            Dim localRepairCount As Integer = 0

            Dim updatedContent = Regex.Replace(
                content,
                "\[([^\]]+)\]\(([^)]+)\)",
                Function(m)
                    Dim linkText = m.Groups(1).Value
                    Dim target = m.Groups(2).Value.Trim()

                    If Not target.EndsWith(".md", StringComparison.OrdinalIgnoreCase) Then
                        Return m.Value
                    End If

                    If target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                       target.StartsWith("https://", StringComparison.OrdinalIgnoreCase) OrElse
                       target.StartsWith("file://", StringComparison.OrdinalIgnoreCase) Then
                        Return m.Value
                    End If

                    Dim resolved = ResolveWikiLinkTarget(currentPagePath, target, kbRootPath)
                    If String.IsNullOrWhiteSpace(resolved) Then
                        Return m.Value
                    End If

                    Dim absoluteTarget = Path.Combine(GetWikiRootPath(kbRootPath), resolved.Replace("/"c, Path.DirectorySeparatorChar))
                    If Not File.Exists(absoluteTarget) Then
                        Return m.Value
                    End If

                    Dim corrected = NormalizeWikiRelativePath(GetRelativePathCompat(Path.GetDirectoryName(currentPagePath), absoluteTarget))
                    If String.Equals(corrected, target, StringComparison.OrdinalIgnoreCase) Then
                        Return m.Value
                    End If

                    localRepairCount += 1
                    Return $"[{linkText}]({corrected})"
                End Function)

            repairedCount = localRepairCount
            Return updatedContent
        End Function

        Private Shared Function ParseSupplementalPageDrafts(response As String) As List(Of SupplementalWikiPageDraft)
            Dim result As New List(Of SupplementalWikiPageDraft)()

            If String.IsNullOrWhiteSpace(response) Then Return result
            If response.Trim().Equals("NONE", StringComparison.OrdinalIgnoreCase) Then Return result

            Dim matches = Regex.Matches(response, "<PAGE>(.*?)</PAGE>", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
            For Each match As Match In matches
                Dim inner = match.Groups(1).Value.Trim()
                If String.IsNullOrWhiteSpace(inner) Then Continue For

                Dim parsed = ParseAgentResponse(inner, "Concept")
                If String.IsNullOrWhiteSpace(parsed.Title) OrElse String.IsNullOrWhiteSpace(parsed.Body) Then
                    Continue For
                End If

                result.Add(New SupplementalWikiPageDraft With {
                    .Title = parsed.Title,
                    .Summary = parsed.Summary,
                    .Kind = parsed.Kind,
                    .Body = parsed.Body
                })
            Next

            Return result
        End Function

        Private Shared Function ComposeSyntheticAgentResponse(title As String,
                                                              summary As String,
                                                              kind As String,
                                                              body As String) As String
            Return ComposeSyntheticAgentResponse(title, summary, kind, body, Nothing)
        End Function

        Private Shared Function ComposeSyntheticAgentResponse(title As String,
                                                              summary As String,
                                                              kind As String,
                                                              body As String,
                                                              schema As KnowledgeStoreSchema) As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"TITLE: {If(title, "").Trim()}")
            sb.AppendLine($"SUMMARY: {If(summary, "").Trim()}")
            sb.AppendLine($"KIND: {NormalizePageKind(kind, "Concept", schema)}")
            sb.AppendLine()
            sb.AppendLine(If(body, "").Trim())
            Return sb.ToString().Trim()
        End Function



        Private Shared Sub ReportMaintenanceProgress(progressCallback As Action(Of String), statusText As String)
            If progressCallback Is Nothing OrElse String.IsNullOrWhiteSpace(statusText) Then
                Return
            End If

            Try
                progressCallback.Invoke(statusText.Trim())
            Catch
            End Try
        End Sub


    End Class
End Namespace