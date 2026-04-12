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
'   - Schema model:
'       * Define configurable store-level guidance such as domain, focus areas,
'         entity types, preferred page kinds, and required sections.
'       * Expose feature flags for cross-linking, source-linking, contradiction
'         detection, query filing, and ignored topics.
'   - Persistence:
'       * Store schema data in `.redink\schema.json` inside each Knowledge
'         Store root.
'       * Load existing schema files or create defaults on first use.
'       * Save schema updates back to disk.
'   - Runtime defaults:
'       * Provide sensible baseline settings when no schema exists or loading
'         fails.
'
' Notes:
'   - The schema is consumed by indexing, wiki generation, linting, and query
'     workflows.
'   - File name intentionally matches the existing project naming convention.
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

    End Class

End Namespace