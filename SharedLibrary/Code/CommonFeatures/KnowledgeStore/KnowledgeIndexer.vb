' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeIndexer.vb
' Purpose: Indexes documents for the Knowledge Store by extracting text content
'          and generating searchable metadata (title, summary, keywords).
'
' Architecture / How it works:
'  - Reads the document using existing SharedLibrary file-reading utilities.
'  - Extracts a title (first heading or first line).
'  - Generates a short summary and keyword list for search matching.
'  - When UseLLMIndex is True, delegates summary/keyword generation to the LLM
'    for richer, more context-aware metadata.
'  - Returns a KnowledgeEntry ready for persistence in the index.
'
' External Dependencies:
'  - SharedMethods file-reading utilities (ReadFileContent).
'  - KnowledgeStoreManager.KnowledgeEntry for the data model.
'  - ISharedContext for configuration.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Indexes documents for Knowledge Store inclusion by extracting text, titles, and keywords.
    ''' </summary>
    Public Class KnowledgeIndexer

        ''' <summary>
        ''' File extensions supported for Knowledge Store indexing.
        ''' </summary>
        Public Shared ReadOnly SupportedExtensions As String() = {
            ".txt", ".rtf", ".docx", ".pdf", ".pptx", ".xlsx",
            ".json", ".xml", ".html", ".htm", ".md", ".csv",
            ".yaml", ".yml", ".ini", ".log",
            ".vb", ".cs", ".js", ".ts", ".py", ".java",
            ".cpp", ".c", ".h", ".sql"
        }

        ''' <summary>
        ''' Maximum characters to extract from a document for the content summary.
        ''' </summary>
        Private Const MaxSummaryChars As Integer = 2000

        ''' <summary>
        ''' Maximum characters to store as the full content snapshot.
        ''' </summary>
        Private Const MaxContentChars As Integer = 80000

        ''' <summary>
        ''' Maximum characters sent to the LLM for summary/keyword generation.
        ''' </summary>
        Private Const MaxLLMInputChars As Integer = 12000

        ''' <summary>
        ''' System prompt for LLM-based summary and keyword extraction.
        ''' </summary>
        Private Const LLMIndexPrompt As String =
            "You are an indexing assistant. Given the following document excerpt, produce:" & vbCrLf &
            "1. A concise title (max 100 chars) on the first line, prefixed with ""TITLE: ""." & vbCrLf &
            "2. A summary of the document (max 500 chars) on the next line(s), prefixed with ""SUMMARY: ""." & vbCrLf &
            "3. A comma-separated list of 10-20 keywords on the last line, prefixed with ""KEYWORDS: ""." & vbCrLf &
            "Output ONLY these three sections. No markdown, no numbering, no extra text."

        ''' <summary>
        ''' Indexes a single document and returns a KnowledgeEntry.
        ''' </summary>
        ''' <param name="filePath">Full or environment-variable path to the file.</param>
        ''' <param name="context">Shared context for configuration.</param>
        ''' <param name="useLLMIndex">When True, uses the LLM for richer summary/keyword generation.</param>
        ''' <returns>A populated KnowledgeEntry, or Nothing if the file cannot be read.</returns>
        Public Shared Async Function IndexDocumentAsync(
                filePath As String,
                context As ISharedContext,
                Optional useLLMIndex As Boolean = False) As Task(Of KnowledgeStoreManager.KnowledgeEntry)

            Dim expandedPath = SharedMethods.ExpandEnvironmentVariables(If(filePath, "").Trim().Trim(""""c, "'"c))
            If Not File.Exists(expandedPath) Then Return Nothing

            ' Check extension
            Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()
            If Not SupportedExtensions.Contains(ext) Then Return Nothing

            ' Read content using existing infrastructure
            Dim content As String = ""
            Try
                content = Await Task.Run(Function() ReadFileForIndex(expandedPath))
            Catch ex As Exception
                Debug.WriteLine($"KnowledgeIndexer: Error reading '{expandedPath}': {ex.Message}")
                Return Nothing
            End Try

            If String.IsNullOrWhiteSpace(content) Then Return Nothing

            ' Default: local extraction
            Dim title = ExtractTitle(content, expandedPath)
            Dim summary = BuildSummary(content)
            Dim keywords = ExtractKeywords(content)

            ' LLM enrichment — overrides title/summary/keywords if successful
            If useLLMIndex Then
                Try
                    Dim llmInput = If(content.Length > MaxLLMInputChars,
                                      content.Substring(0, MaxLLMInputChars),
                                      content)

                    Dim llmResponse = Await SharedMethods.LLM(
                        context, LLMIndexPrompt, llmInput,
                        "", "", 0L, False, True).ConfigureAwait(False)

                    If Not String.IsNullOrWhiteSpace(llmResponse) Then
                        Dim parsed = ParseLLMIndexResponse(llmResponse)
                        If Not String.IsNullOrWhiteSpace(parsed.Title) Then title = parsed.Title
                        If Not String.IsNullOrWhiteSpace(parsed.Summary) Then summary = parsed.Summary
                        If parsed.Keywords IsNot Nothing AndAlso parsed.Keywords.Length > 0 Then keywords = parsed.Keywords
                    End If
                Catch ex As Exception
                    ' LLM failure is non-fatal — fall back to local extraction (already set above)
                    Debug.WriteLine($"KnowledgeIndexer: LLM enrichment failed for '{expandedPath}': {ex.Message}")
                End Try
            End If

            ' Create entry
            Dim entry As New KnowledgeStoreManager.KnowledgeEntry() With {
                .FilePath = filePath,
                .Title = title,
                .Summary = summary,
                .Keywords = keywords,
                .ContentSnapshot = If(content.Length > MaxContentChars,
                                      content.Substring(0, MaxContentChars),
                                      content),
                .IndexedDate = DateTime.Now,
                .Tags = Nothing,
                .IsFromCentralIndex = False
            }

            Return entry
        End Function

        ''' <summary>
        ''' Parses the structured LLM response into title, summary, and keywords.
        ''' </summary>
        Private Shared Function ParseLLMIndexResponse(response As String) As (Title As String, Summary As String, Keywords As String())
            Dim title As String = ""
            Dim summary As String = ""
            Dim keywords As String() = Nothing

            Dim lines = response.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            Dim summaryLines As New List(Of String)()
            Dim inSummary As Boolean = False

            For Each line In lines
                Dim trimmed = line.Trim()

                If trimmed.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                    title = trimmed.Substring("TITLE:".Length).Trim()
                    inSummary = False
                ElseIf trimmed.StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) Then
                    summaryLines.Add(trimmed.Substring("SUMMARY:".Length).Trim())
                    inSummary = True
                ElseIf trimmed.StartsWith("KEYWORDS:", StringComparison.OrdinalIgnoreCase) Then
                    inSummary = False
                    Dim kwText = trimmed.Substring("KEYWORDS:".Length).Trim()
                    keywords = kwText.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries).
                        Select(Function(k) k.Trim().ToLowerInvariant()).
                        Where(Function(k) k.Length >= 2).
                        Distinct().
                        Take(25).
                        ToArray()
                ElseIf inSummary Then
                    summaryLines.Add(trimmed)
                End If
            Next

            summary = String.Join(" ", summaryLines).Trim()
            If summary.Length > MaxSummaryChars Then
                summary = summary.Substring(0, MaxSummaryChars) & "..."
            End If

            Return (title, summary, keywords)
        End Function

        ''' <summary>
        ''' Reads file content using appropriate method based on extension.
        ''' </summary>
        Private Shared Function ReadFileForIndex(expandedPath As String) As String
            Dim ext = Path.GetExtension(expandedPath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                     ".md", ".yaml", ".yml", ".vb", ".cs", ".js", ".ts", ".py",
                     ".java", ".cpp", ".c", ".h", ".sql"
                    Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)

                Case ".docx"
                    Return SharedMethods.ReadDocxSandboxed(expandedPath)

                Case ".pdf"
                    ' ReadPdfAsText is Async; run synchronously for indexing
                    Dim t = SharedMethods.ReadPdfAsText(expandedPath, ReturnErrorInsteadOfEmpty:=False, DoOCR:=False, AskUser:=False)
                    t.Wait()
                    Return If(t.Result, "")

                Case ".rtf"
                    ' Read RTF as raw text, stripping RTF control words
                    Try
                        Using rtb As New System.Windows.Forms.RichTextBox()
                            rtb.LoadFile(expandedPath, System.Windows.Forms.RichTextBoxStreamType.RichText)
                            Return rtb.Text
                        End Using
                    Catch
                        Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
                    End Try

                Case ".xlsx"
                    Return SharedMethods.ReadXlsxSandboxed(expandedPath)

                Case ".pptx"
                    Return SharedMethods.ReadPptxSandboxed(expandedPath)

                Case Else
                    Return IO.File.ReadAllText(expandedPath, System.Text.Encoding.UTF8)
            End Select
        End Function

        ''' <summary>
        ''' Extracts a title from the content (first heading or first non-empty line).
        ''' Falls back to the filename without extension.
        ''' </summary>
        Private Shared Function ExtractTitle(content As String, filePath As String) As String
            ' Try to find a markdown heading
            Dim headingMatch = Regex.Match(content, "^#{1,3}\s+(.+)$", RegexOptions.Multiline)
            If headingMatch.Success Then
                Dim heading = headingMatch.Groups(1).Value.Trim()
                If heading.Length > 0 AndAlso heading.Length <= 200 Then Return heading
            End If

            ' Try first non-empty line (if reasonable length)
            Dim lines = content.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            For Each line In lines
                Dim trimmed = line.Trim()
                If trimmed.Length >= 3 AndAlso trimmed.Length <= 200 Then
                    Return trimmed
                End If
            Next

            ' Fall back to filename
            Return Path.GetFileNameWithoutExtension(filePath)
        End Function

        ''' <summary>
        ''' Builds a short summary from the document content.
        ''' </summary>
        Private Shared Function BuildSummary(content As String) As String
            ' Clean whitespace
            Dim cleaned = Regex.Replace(content, "\s+", " ").Trim()
            If cleaned.Length <= MaxSummaryChars Then Return cleaned
            ' Cut at word boundary
            Dim cutoff = cleaned.LastIndexOf(" "c, MaxSummaryChars)
            If cutoff < MaxSummaryChars \ 2 Then cutoff = MaxSummaryChars
            Return cleaned.Substring(0, cutoff).Trim() & "..."
        End Function

        ''' <summary>
        ''' Extracts keywords from the content using simple frequency analysis.
        ''' Returns the top N most frequent meaningful words.
        ''' </summary>
        Private Shared Function ExtractKeywords(content As String) As String()
            Dim stopWords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for",
                "of", "with", "by", "from", "is", "are", "was", "were", "be", "been",
                "being", "have", "has", "had", "do", "does", "did", "will", "would",
                "could", "should", "may", "might", "shall", "can", "this", "that",
                "these", "those", "it", "its", "not", "no", "nor", "as", "if", "then",
                "than", "so", "such", "der", "die", "das", "und", "oder", "in", "von",
                "zu", "den", "dem", "des", "ein", "eine", "einer", "einem", "einen",
                "mit", "auf", "ist", "sind", "war", "wird", "werden", "hat", "haben",
                "le", "la", "les", "de", "du", "des", "un", "une", "et", "ou", "en"
            }

            ' Tokenize
            Dim words = Regex.Matches(content, "\b[a-zA-ZÀ-ÿ]{3,}\b")
            Dim freq As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For Each m As Match In words
                Dim w = m.Value
                If stopWords.Contains(w) Then Continue For
                If Not freq.ContainsKey(w) Then freq(w) = 0
                freq(w) += 1
            Next

            ' Return top 20 keywords
            Return freq.OrderByDescending(Function(kv) kv.Value).
                Take(20).
                Select(Function(kv) kv.Key.ToLowerInvariant()).
                ToArray()
        End Function

    End Class

End Namespace