' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStore.Schema.vb
' Purpose:
'   Defines and persists per-store schema settings that steer ingest, wiki
'   structure, linking, contradiction analysis, and query behavior.
'
' Responsibilities:
'   - Define store-level schema metadata such as domain, focus areas, entity
'     types, page kinds, and required sections.
'   - Expose feature flags for cross-linking, source-linking, contradiction
'     detection, query filing, and ignored topics.
'   - Load existing schema files or create defaults on first use.
'   - Persist schema settings to `.redink\schema.json`.
'   - Provide safe runtime defaults when schema loading fails.
'
' Notes:
'   - The schema is consumed by indexing, wiki generation, linting, repair,
'     and query workflows.
'   - The file name intentionally follows the existing project naming
'     convention.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json

Namespace SharedLibrary

    ''' <summary>
    ''' Per-store schema that steers ingest, linking, linting, and query behavior.
    ''' Persisted as .redink\schema.json inside each Knowledge Store root.
    ''' </summary>
    ''' <summary>
    ''' Per-store schema that steers ingest, linking, linting, and query behavior.
    ''' Persisted as .redink\schema.json inside each Knowledge Store root.
    ''' </summary>
    Public Class KnowledgeStoreSchema

        Public Property SchemaVersion As Integer = 1
        Public Property Domain As String = ""
        Public Property FocusAreas As String() = {
            "key entities",
            "important concepts",
            "contradictions",
            "changes over time",
            "source-backed claims"
        }
        Public Property EntityTypes As String() = {
            "Person",
            "Organization",
            "Concept",
            "Topic",
            "Artifact",
            "Event"
        }
        Public Property PreferredPageKinds As String() = {
            "Source",
            "Entity",
            "Concept",
            "Analysis"
        }
        Public Property RequiredSections As String() = {
            "Summary",
            "Key Points",
            "Related Pages",
            "Sources"
        }
        Public Property AlwaysCreateCrossLinks As Boolean = True
        Public Property AlwaysAddSourceLinks As Boolean = True
        Public Property DetectContradictions As Boolean = True
        Public Property QueryFilingEnabled As Boolean = False
        Public Property IgnoredTopics As String() = {}
        Public Property AdditionalInstructions As String = ""

        ''' <summary>
        ''' Minimum number of supplemental wiki pages the LLM should aim to produce
        ''' per ingested source (in addition to the primary source page). The model
        ''' is told this is a soft floor; it may return fewer for genuinely thin
        ''' sources. Set together with <see cref="MaxSupplementalPagesPerSource"/>
        ''' to control fan-out cost during ingest.
        ''' </summary>
        Public Property MinSupplementalPagesPerSource As Integer = 2

        ''' <summary>
        ''' Hard upper bound on supplemental wiki pages applied to the wiki per
        ''' ingested source. Drafts beyond this count are discarded after parsing.
        ''' Set to 0 to disable supplemental page generation entirely (skips the
        ''' supplemental LLM call as well, so it is also a cost optimization).
        ''' </summary>
        Public Property MaxSupplementalPagesPerSource As Integer = 15

        ''' <summary>
        ''' When True, contradiction and superseded-page scanning use a token→page
        ''' inverted index to skip pair scoring for pages with no meaningful token
        ''' overlap. Faster on large wikis, but pairs whose shared-token count is
        ''' below <see cref="PairScanMinSharedTokens"/> are no longer considered.
        ''' Defaults to False to preserve historical recall.
        ''' </summary>
        Public Property UseInvertedIndexPairScan As Boolean = False

        ''' <summary>
        ''' Minimum shared tokens required for two pages to be considered a
        ''' candidate pair under the inverted-index pre-filter. Ignored when
        ''' <see cref="UseInvertedIndexPairScan"/> is False. Lower values cost
        ''' more CPU but lose less recall; 2 is a reasonable starting point.
        ''' </summary>
        Public Property PairScanMinSharedTokens As Integer = 2

        Public Property StoreName As String = ""

        ''' <summary>
        ''' Free-text description of the store's content and purpose, surfaced to the LLM
        ''' in tooling mode so it can decide when to query this store. Example:
        ''' "Contains internal compliance policies, anti-money-laundering procedures,
        ''' and client onboarding checklists for VISCHER AG."
        ''' </summary>
        Public Property ToolingDescription As String = ""

        Public Shared Function GetSchemaPath(kbRootPath As String) As String
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return ""
            Return Path.Combine(kbRootPath, ".redink", "schema.json")
        End Function

        Public Shared Function LoadOrCreate(kbRootPath As String) As KnowledgeStoreSchema
            Dim schema As KnowledgeStoreSchema = Nothing
            Dim schemaPath = GetSchemaPath(kbRootPath)

            If String.IsNullOrWhiteSpace(schemaPath) Then
                Return New KnowledgeStoreSchema()
            End If

            Try
                If File.Exists(schemaPath) Then
                    Dim json = File.ReadAllText(schemaPath, Encoding.UTF8)
                    If Not String.IsNullOrWhiteSpace(json) Then
                        schema = JsonConvert.DeserializeObject(Of KnowledgeStoreSchema)(json)
                    End If
                End If
            Catch
                schema = Nothing
            End Try

            If schema Is Nothing Then
                schema = New KnowledgeStoreSchema()
                Save(kbRootPath, schema)
            End If

            Return schema
        End Function

        Public Shared Sub Save(kbRootPath As String, schema As KnowledgeStoreSchema)
            If schema Is Nothing Then schema = New KnowledgeStoreSchema()

            Dim schemaPath = GetSchemaPath(kbRootPath)
            If String.IsNullOrWhiteSpace(schemaPath) Then Return

            Try
                Dim dir = Path.GetDirectoryName(schemaPath)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim json = JsonConvert.SerializeObject(schema, Formatting.Indented)
                File.WriteAllText(schemaPath, json, Encoding.UTF8)
            Catch
            End Try
        End Sub

        Public Function ToPromptBlock() As String
            Dim sb As New StringBuilder()

            sb.AppendLine("STORE-SPECIFIC WIKI SCHEMA:")
            sb.AppendLine($"- Domain: {If(String.IsNullOrWhiteSpace(Domain), "(not specified)", Domain)}")

            If FocusAreas IsNot Nothing AndAlso FocusAreas.Length > 0 Then
                sb.AppendLine("- Focus Areas: " & String.Join(", ", FocusAreas))
            End If

            If EntityTypes IsNot Nothing AndAlso EntityTypes.Length > 0 Then
                sb.AppendLine("- Entity Types: " & String.Join(", ", EntityTypes))
            End If

            If PreferredPageKinds IsNot Nothing AndAlso PreferredPageKinds.Length > 0 Then
                sb.AppendLine("- Preferred Page Kinds: " & String.Join(", ", PreferredPageKinds))
            End If

            If RequiredSections IsNot Nothing AndAlso RequiredSections.Length > 0 Then
                sb.AppendLine("- Required Sections: " & String.Join(", ", RequiredSections))
            End If

            sb.AppendLine($"- Always Create Cross Links: {AlwaysCreateCrossLinks}")
            sb.AppendLine($"- Always Add Source Links: {AlwaysAddSourceLinks}")
            sb.AppendLine($"- Detect Contradictions: {DetectContradictions}")
            sb.AppendLine($"- Query Filing Enabled: {QueryFilingEnabled}")

            If IgnoredTopics IsNot Nothing AndAlso IgnoredTopics.Length > 0 Then
                sb.AppendLine("- Ignored Topics: " & String.Join(", ", IgnoredTopics))
            End If

            If Not String.IsNullOrWhiteSpace(AdditionalInstructions) Then
                sb.AppendLine("- Additional Instructions: " & AdditionalInstructions)
            End If

            Return sb.ToString().Trim()
        End Function

        Public Property SourceAccess As New SourceAccessOptions()
        Public Property SourceRegistry As New SourceRegistryOptions()
        Public Property PerClaimCitations As New PerClaimCitationOptions()

        Public Property ConceptUrlTemplates As New List(Of ConceptUrlTemplate)()

        Public Class SourceAccessOptions
            Public Property FilesFolderName As String = "_files"
            Public Property DeduplicateByHash As Boolean = True
            Public Property MaxInlineFileSizeMB As Integer = 50
            Public Property AllowedExtensions As String() = {".pdf", ".docx", ".doc", ".txt", ".md", ".html", ".htm", ".rtf", ".pptx", ".xlsx", ".png", ".jpg", ".jpeg"}
        End Class

        Public Class SourceRegistryOptions
            Public Property Enabled As Boolean = True
            Public Property FolderName As String = "Sources"          ' dossier folder
            ' Front-matter fields the LLM is asked to populate on dossier pages.
            Public Property MetadataFields As String() = {"citation", "author", "issued_date", "jurisdiction", "language", "doc_type"}
        End Class

        Public Class PerClaimCitationOptions
            Public Property Enabled As Boolean = True
            Public Property Style As String = "bracket-id"            ' bracket-id | footnote | none
        End Class

        ''' <summary>
        ''' Free-form natural language identifier (e.g. "German", "de-CH", "English",
        ''' "French"). When non-empty, all LLM-generated page content (titles, summaries,
        ''' bullets, prose) is produced in this language. Technical identifiers (kind
        ''' values, file names, [S#] markers, YAML keys) remain unchanged.
        ''' </summary>
        Public Property OutputLanguage As String = ""

        ''' <summary>
        ''' Optional explicit names for the four analytical roles. When set, these win
        ''' over fuzzy detection from RequiredSections. Leave empty to use the
        ''' RequiredSections-based fallback.
        ''' </summary>
        Public Property AnalyticalSections As New AnalyticalSectionsOptions()

        ''' <summary>
        ''' Visible UI labels that the wiki engine emits deterministically (section
        ''' headings, badges, placeholder text). Keys are stable English identifiers;
        ''' values are the desired display text. Missing keys fall back to the English
        ''' default in code.
        ''' </summary>
        Public Property Labels As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Returns the configured label for the given key, or the supplied English
        ''' default when no override exists. The default is also returned if the
        ''' configured value is whitespace-only.
        ''' </summary>
        Public Function GetLabel(key As String, defaultText As String) As String
            If Labels IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(key) AndAlso
               Labels.ContainsKey(key) AndAlso
               Not String.IsNullOrWhiteSpace(Labels(key)) Then
                Return Labels(key)
            End If
            Return defaultText
        End Function

        ''' <summary>
        ''' Returns either an empty string or a ready-to-prepend prompt block that
        ''' instructs the LLM to write in the configured OutputLanguage.
        ''' </summary>
        Public Function ResolveLanguageDirective() As String
            If String.IsNullOrWhiteSpace(OutputLanguage) Then Return ""
            Dim sb As New StringBuilder()
            sb.AppendLine("LANGUAGE:")
            sb.AppendLine($"- Write ALL page content (titles, summaries, headings, bullets, prose) in {OutputLanguage.Trim()}.")
            sb.AppendLine("- Keep technical identifiers in their original form: kind values, file names, source paths, [S#] markers, YAML keys, dossier slugs.")
            sb.AppendLine("- If the source is in another language, translate naturally; do not transliterate proper names or quoted legal text.")
            sb.AppendLine($"- Use ONLY the section headings supplied in the REQUIRED BODY SECTIONS list. Do NOT add any additional sections, and do NOT include English fallback equivalents (e.g. ""Key Claims"", ""Evidence"", ""Open Questions"", ""Contradictions"", ""Sources"", ""Related Pages"") if they are not on that list.")
            sb.AppendLine($"- For empty sections, write the empty-state placeholder in {OutputLanguage.Trim()} (e.g. German: ""- _Keine._""; French: ""- _Aucune._""; English: ""- _None._""). Never use English when the page language is not English.")
            Return sb.ToString().TrimEnd() & vbCrLf & vbCrLf
        End Function

        Public Class AnalyticalSectionsOptions
            Public Property Claims As String = ""
            Public Property Evidence As String = ""
            Public Property Contradictions As String = ""
            Public Property OpenQuestions As String = ""
        End Class


        Public Class ConceptUrlTemplate
            Public Property Match As String = ""        ' simple substring/regex on concept title
            Public Property MatchKind As String = "contains" ' contains | regex
            Public Property Url As String = ""          ' may contain {title} placeholder
            Public Property Description As String = ""  ' "Use when the concept is a Swiss federal law on fedlex.admin.ch"
        End Class

    End Class

End Namespace