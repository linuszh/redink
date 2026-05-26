' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeEmbeddingService.vb
' Purpose:
'   Provides embedding generation, embedding persistence, and local semantic
'   search for Knowledge Store content.
'
' Responsibilities:
'   - Acquire embeddings through the configured embedding model or special-task
'     model override.
'   - Split source and wiki content into chunks suitable for embedding.
'   - Maintain `.redink\.embeddings.json` safely for each Knowledge Store root.
'   - Rebuild or incrementally update embeddings for changed wiki pages.
'   - Execute semantic retrieval with keyword fallback when necessary.
'
' Notes:
'   - Embedding updates are serialized to avoid concurrent index corruption.
'   - Search operates on the local embedding cache and does not require a live
'     external vector database.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext
Imports Newtonsoft.Json

Namespace SharedLibrary

    ''' <summary>
    ''' Represents a single semantic text chunk stored in the Knowledge Base.
    ''' </summary>
    Public Class EmbeddingRecord
        Public Property Id As String
        Public Property FilePath As String
        Public Property TextChunk As String
        Public Property Vector As Single()
    End Class

    ''' <summary>
    ''' Represents an extracted result during a Knowledge Base semantic or keyword search.
    ''' </summary>
    Public Class KnowledgeSearchResult
        Public Property FilePath As String
        Public Property TextChunk As String
        Public Property Score As Double
    End Class

    ''' <summary>
    ''' Handles universal interaction with Embedding APIs and local fast-math Cosine Similarity & BM25 search.
    ''' </summary>
    Public Class KnowledgeEmbeddingService

        Private Shared ReadOnly VectorCacheLock As New Object()
        Private Shared ReadOnly EmbeddingUpdateLock As New Threading.SemaphoreSlim(1, 1)

        ''' <summary>
        ''' Gets an embedding vector for a given text by calling the configured LLM API.
        ''' </summary>
        Public Shared Async Function GetEmbeddingAsync(text As String, context As ISharedContext) As Task(Of Single())
            Dim backupConfig As ModelConfig = Nothing
            Dim useAlternateAPI As Boolean = False

            Try
                ' 1. Try to load an "Embedding" Model Config
                Dim altModelPath As String = ""
                Dim propInfo = context.GetType().GetProperty("INI_AlternateModelPath")
                If propInfo IsNot Nothing Then
                    altModelPath = TryCast(propInfo.GetValue(context), String)
                End If

                If Not String.IsNullOrWhiteSpace(altModelPath) Then
                    backupConfig = SharedMethods.GetCurrentConfig(context)

                    Dim success = SharedMethods.GetSpecialTaskModel(context, altModelPath, "Embedding")
                    If success Then
                        useAlternateAPI = True
                    Else
                        ' If it failed to find [Embedding], discard the backup so we don't accidentally overwrite anything
                        backupConfig = Nothing
                    End If
                End If

                Dim config = SharedMethods.GetCurrentConfig(context)

                ' If the user hasn't explicitly configured a response key starting with "array:", fail gracefully.
                If config Is Nothing OrElse String.IsNullOrWhiteSpace(config.Response) OrElse Not config.Response.StartsWith("array:", StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If

                ' 2. Call API via existing SharedMethods.LLM (using the correct API stream!)
                Dim responseString As String = Await SharedMethods.LLM(
                    context:=context,
                    promptSystem:="",
                    promptUser:=text,
                    Model:="",
                    Temperature:="",
                    Timeout:=0,
                    UseSecondAPI:=useAlternateAPI,
                    Hidesplash:=True)

                If String.IsNullOrWhiteSpace(responseString) Then Return Nothing

                ' PREVENT CORRUPTION:
                ' If there's no pipe or comma, it's an API error string (e.g. "HTTP Error 404..."), not a float array!
                If Not responseString.Contains("|") AndAlso Not responseString.Contains(",") Then
                    Try
                        System.IO.File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "LLM_Embedding_Failed_Response.txt"), "The API returned an error string instead of an Array:" & vbCrLf & responseString)
                    Catch
                    End Try
                    Return Nothing
                End If

                ' 3. Parse Float Array from the separator-delimited string
                Dim separator As Char = If(responseString.Contains("|"), "|"c, ","c)
                Dim stringParts As String() = responseString.Split(separator)

                Dim vector(stringParts.Length - 1) As Single
                For i As Integer = 0 To stringParts.Length - 1
                    Single.TryParse(stringParts(i).Trim(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, vector(i))
                Next

                Return vector

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KnowledgeEmbeddingService] Failed to fetch embedding: {ex.Message}")
                Return Nothing
            Finally
                ' Restore primary chat model configuration to the _2 variables safely
                If backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(context, backupConfig)
                End If
            End Try
        End Function

        ''' <summary>
        ''' Computes the Cosine Similarity between two vectors. Pure VB.NET, no dependencies.
        ''' </summary>
        Public Shared Function CosineSimilarity(vectorA As Single(), vectorB As Single()) As Double
            If vectorA Is Nothing OrElse vectorB Is Nothing OrElse vectorA.Length <> vectorB.Length Then Return 0

            Dim dotProduct As Double = 0.0
            Dim normA As Double = 0.0
            Dim normB As Double = 0.0

            For i As Integer = 0 To vectorA.Length - 1
                dotProduct += vectorA(i) * vectorB(i)
                normA += vectorA(i) * vectorA(i)
                normB += vectorB(i) * vectorB(i)
            Next

            If normA = 0 OrElse normB = 0 Then Return 0
            Return dotProduct / (Math.Sqrt(normA) * Math.Sqrt(normB))
        End Function

        Private Shared Function GetEmbeddingStatePath(kbRootPath As String) As String
            Return Path.Combine(kbRootPath, ".redink", ".embedding-state.json")
        End Function

        Private Shared Function ComputeContentHash(text As String) As String
            Using sha = System.Security.Cryptography.SHA256.Create()
                Dim bytes = sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(If(text, "")))
                Return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Private Shared Function LoadContentHashState(path As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            If Not File.Exists(path) Then
                Return result
            End If

            Try
                Dim json = File.ReadAllText(path)
                Dim raw = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(json)

                If raw Is Nothing Then
                    Return result
                End If

                For Each kvp In raw
                    If String.IsNullOrWhiteSpace(kvp.Key) Then
                        Continue For
                    End If

                    result(kvp.Key) = If(kvp.Value, "")
                Next
            Catch
            End Try

            Return result
        End Function

        Private Shared Sub SaveContentHashState(path As String, state As Dictionary(Of String, String))
            Try
                Dim dir = System.IO.Path.GetDirectoryName(path)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim json = JsonConvert.SerializeObject(state, Formatting.None)
                Dim tmpPath = path & ".tmp"

                File.WriteAllText(tmpPath, json)
                If File.Exists(path) Then
                    File.Delete(path)
                End If
                File.Move(tmpPath, path)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"Failed to save embedding state: {ex.Message}")
            End Try
        End Sub

        ' =====================================================================
        ' KB INDEXING & STORAGE
        ' =====================================================================

        ''' <summary>
        ''' Chunks text by paragraph and combines them to approximately the desired character count.
        ''' Subdivides large paragraphs to avoid huge chunk overflow boundaries.
        ''' </summary>
        Public Shared Function ChunkText(text As String, Optional maxChars As Integer = 1500) As List(Of String)
            Dim chunks As New List(Of String)()
            If String.IsNullOrWhiteSpace(text) Then Return chunks

            ' Split by double newlines or single newlines (including varying environments' returns)
            Dim paragraphs = text.Split({vbCrLf & vbCrLf, vbLf & vbLf, vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
            Dim currentChunk As String = ""

            For Each p In paragraphs
                Dim cleanP = p.Trim()
                If String.IsNullOrWhiteSpace(cleanP) Then Continue For

                ' Break down a singular paragraph if it is excessively large (larger than the chunk text itself)
                While cleanP.Length > maxChars
                    Dim part = cleanP.Substring(0, maxChars)
                    ' Try to find a space to break near instead of abruptly truncating words
                    Dim lastSpace = part.LastIndexOf(" "c)
                    If lastSpace > 0 Then
                        part = part.Substring(0, lastSpace)
                    End If

                    If currentChunk.Length > 0 Then
                        chunks.Add(currentChunk.Trim())
                        currentChunk = ""
                    End If
                    chunks.Add(part.Trim())
                    cleanP = cleanP.Substring(part.Length).Trim()
                End While

                If cleanP.Length = 0 Then Continue For

                If currentChunk.Length + cleanP.Length > maxChars AndAlso currentChunk.Length > 0 Then
                    chunks.Add(currentChunk.Trim())
                    currentChunk = cleanP
                Else
                    currentChunk &= If(currentChunk.Length > 0, " ", "") & cleanP
                End If
            Next

            If currentChunk.Trim().Length > 0 Then chunks.Add(currentChunk.Trim())
            Return chunks
        End Function

        ''' <summary>
        ''' Chunks a file's content, grabs Vector embeddings if available, and updates the local store.
        ''' </summary>
        Public Shared Async Function UpdateFileEmbeddingsAsync(kbRootPath As String, filePath As String, text As String, context As ISharedContext) As Task
            If String.IsNullOrWhiteSpace(kbRootPath) OrElse
               String.IsNullOrWhiteSpace(filePath) OrElse
               String.IsNullOrWhiteSpace(text) Then
                Return
            End If

            Dim embeddingText = RemoveFrontMatter(text)
            Dim contentHash = ComputeContentHash(embeddingText)

            Await EmbeddingUpdateLock.WaitAsync().ConfigureAwait(False)
            Try
                Dim indexPath As String = Path.Combine(kbRootPath, ".redink", ".embeddings.json")
                Dim statePath As String = GetEmbeddingStatePath(kbRootPath)
                Dim contentHashes = LoadContentHashState(statePath)

                Dim previousHash As String = ""
                If contentHashes.TryGetValue(filePath, previousHash) AndAlso
                   String.Equals(previousHash, contentHash, StringComparison.OrdinalIgnoreCase) Then
                    Return
                End If

                Dim records As List(Of EmbeddingRecord) = LoadIndex(indexPath)
                records.RemoveAll(Function(r) String.Equals(r.FilePath, filePath, StringComparison.OrdinalIgnoreCase))

                Dim chunks = ChunkText(embeddingText)

                For Each chunk In chunks
                    Dim vector = Await GetEmbeddingAsync(chunk, context).ConfigureAwait(False)

                    records.Add(New EmbeddingRecord With {
                        .Id = Guid.NewGuid().ToString(),
                        .FilePath = filePath,
                        .TextChunk = chunk,
                        .Vector = vector
                    })
                Next

                SaveIndex(indexPath, records)

                contentHashes(filePath) = contentHash
                SaveContentHashState(statePath, contentHashes)
            Finally
                EmbeddingUpdateLock.Release()
            End Try
        End Function


        ''' <summary>
        ''' Rebuilds the entire embedding index for a Knowledge Store from the existing wiki pages.
        ''' Useful after changing the configured embedding model.
        ''' </summary>
        Public Shared Async Function RebuildAllWikiEmbeddingsAsync(kbRootPath As String,
                                                                   context As ISharedContext,
                                                                   Optional progressPrefix As String = "",
                                                                   Optional progressOffset As Integer = 0,
                                                                   Optional progressTotal As Integer = 0) As Task(Of Integer)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return 0

            Dim wikiRoot As String = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then Return 0

            Dim wikiPages As New List(Of String)()
            For Each filePath In Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories)
                Dim name As String = Path.GetFileName(filePath)

                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                wikiPages.Add(filePath)
            Next

            Dim indexPath As String = Path.Combine(kbRootPath, ".redink", ".embeddings.json")

            ' Start from a clean embedding index so no stale vectors survive model changes.
            SaveIndex(indexPath, New List(Of EmbeddingRecord)())

            SaveContentHashState(GetEmbeddingStatePath(kbRootPath), New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))

            Dim rebuiltCount As Integer = 0

            For i As Integer = 0 To wikiPages.Count - 1
                If ProgressBarModule.CancelOperation Then
                    Exit For
                End If

                Dim filePath As String = wikiPages(i)
                Dim absoluteIndex As Integer = progressOffset + i + 1

                If progressTotal > 0 Then
                    ProgressBarModule.GlobalProgressMax = progressTotal
                    ProgressBarModule.GlobalProgressValue = absoluteIndex - 1

                    Dim labelPrefix As String = If(String.IsNullOrWhiteSpace(progressPrefix),
                                                   "Refreshing embeddings",
                                                   $"Refreshing embeddings - {progressPrefix}")

                    ProgressBarModule.GlobalProgressLabel =
                        $"{labelPrefix} ({absoluteIndex}/{progressTotal}): {Path.GetFileName(filePath)}"
                End If

                Try
                    Dim text As String = File.ReadAllText(filePath, System.Text.Encoding.UTF8)
                    If String.IsNullOrWhiteSpace(text) Then
                        If progressTotal > 0 Then
                            ProgressBarModule.GlobalProgressValue = absoluteIndex
                        End If
                        Continue For
                    End If

                    Await UpdateFileEmbeddingsAsync(
                        kbRootPath:=kbRootPath,
                        filePath:=filePath,
                        text:=text,
                        context:=context).ConfigureAwait(False)

                    rebuiltCount += 1
                Catch ex As Exception
                    KnowledgeWikiService.LogWikiError(kbRootPath, filePath, $"Embedding rebuild failed: {ex.Message}")
                Finally
                    If progressTotal > 0 Then
                        ProgressBarModule.GlobalProgressValue = absoluteIndex
                    End If
                End Try
            Next

            KnowledgeWikiService.AppendOperationalLog(kbRootPath, "embeddings-refresh", $"Pages={rebuiltCount}")

            Return rebuiltCount
        End Function

        ''' <summary>
        ''' Rebuilds the entire embedding index for a Knowledge Store from the existing wiki pages.
        ''' Useful after changing the configured embedding model.
        ''' </summary>
        Public Shared Async Function RebuildAllWikiEmbeddingsAsync(kbRootPath As String,
                                                                   context As ISharedContext) As Task(Of Integer)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return 0

            Dim wikiRoot As String = Path.Combine(kbRootPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then Return 0

            Dim indexPath As String = Path.Combine(kbRootPath, ".redink", ".embeddings.json")

            ' Start from a clean embedding index so no stale vectors survive model changes.
            SaveIndex(indexPath, New List(Of EmbeddingRecord)())

            SaveContentHashState(GetEmbeddingStatePath(kbRootPath), New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase))

            Dim rebuiltCount As Integer = 0

            For Each filePath In Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories)
                Dim name As String = Path.GetFileName(filePath)

                If name.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) Then Continue For
                If name.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                Try
                    Dim text As String = File.ReadAllText(filePath, System.Text.Encoding.UTF8)
                    If String.IsNullOrWhiteSpace(text) Then Continue For

                    Await UpdateFileEmbeddingsAsync(
                        kbRootPath:=kbRootPath,
                        filePath:=filePath,
                        text:=text,
                        context:=context).ConfigureAwait(False)

                    rebuiltCount += 1
                Catch ex As Exception
                    KnowledgeWikiService.LogWikiError(kbRootPath, filePath, $"Embedding rebuild failed: {ex.Message}")
                End Try
            Next

            KnowledgeWikiService.AppendOperationalLog(kbRootPath, "embeddings-refresh", $"Pages={rebuiltCount}")

            Return rebuiltCount
        End Function


        Private Shared Function LoadIndex(path As String) As List(Of EmbeddingRecord)
            If Not File.Exists(path) Then Return New List(Of EmbeddingRecord)()
            Try
                Dim json = File.ReadAllText(path)
                Return JsonConvert.DeserializeObject(Of List(Of EmbeddingRecord))(json)
            Catch
                Return New List(Of EmbeddingRecord)()
            End Try
        End Function

        Private Shared Sub SaveIndex(path As String, records As List(Of EmbeddingRecord))
            Try
                SyncLock VectorCacheLock
                    Dim dir = System.IO.Path.GetDirectoryName(path)
                    If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If

                    Dim json = JsonConvert.SerializeObject(records, Formatting.None)
                    Dim tmpPath = path & ".tmp"

                    File.WriteAllText(tmpPath, json)
                    If File.Exists(path) Then
                        File.Delete(path)
                    End If
                    File.Move(tmpPath, path)
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"Failed to save embeddings: {ex.Message}")
            End Try
        End Sub

        ' =====================================================================
        ' KB SEARCH ORCHESTRATION 
        ' =====================================================================

        ''' <summary>
        ''' Searches the KB using Semantic Embeddings or Okapi BM25 Keyword fallback.
        ''' </summary>
        Public Shared Async Function SearchAsync(kbRootPath As String, query As String, topK As Integer, context As ISharedContext) As Task(Of List(Of KnowledgeSearchResult))
            Dim indexPath As String = Path.Combine(kbRootPath, ".redink", ".embeddings.json")
            Dim records As List(Of EmbeddingRecord) = LoadIndex(indexPath)
            If records.Count = 0 Then Return New List(Of KnowledgeSearchResult)()

            Dim queryVec = Await GetEmbeddingAsync(query, context)

            If queryVec IsNot Nothing Then
                ' VECTOR SEARCH (Cosine Similarity)
                System.Diagnostics.Debug.WriteLine($"[KB SEARCH] Valid vectors found. Executing VECTOR SEARCH (Cosine) for '{query}'")
                Dim results As New List(Of KnowledgeSearchResult)()
                For Each r In records
                    If r.Vector IsNot Nothing Then
                        Dim score = CosineSimilarity(queryVec, r.Vector)
                        If score > 0.4 Then ' Arbitrary threshold to ignore garbage
                            results.Add(New KnowledgeSearchResult With {.FilePath = r.FilePath, .TextChunk = r.TextChunk, .Score = score})
                        End If
                    End If
                Next
                Return results.OrderByDescending(Function(x) x.Score).Take(topK).ToList()
            Else
                ' BM25 KEYWORD SEARCH (Fallback when no vector model is assigned)
                System.Diagnostics.Debug.WriteLine($"[KB SEARCH] No valid vector config. Executing KEYWORD SEARCH (BM25) fallback for '{query}'")
                Return PerformBM25Search(query, records, topK)
            End If
        End Function

        Private Shared Function Tokenize(text As String) As String()
            If String.IsNullOrWhiteSpace(text) Then Return New String() {}
            Dim cleaned = Regex.Replace(text.ToLowerInvariant(), "[^\w\s]", " ")
            Return cleaned.Split(New Char() {" "c, ControlChars.Cr, ControlChars.Lf, ControlChars.Tab}, StringSplitOptions.RemoveEmptyEntries)
        End Function

        ''' <summary>
        ''' Pure VB.NET implementation of Okapi BM25 scoring for keyword search fallback.
        ''' </summary>
        Private Shared Function PerformBM25Search(query As String, records As List(Of EmbeddingRecord), topK As Integer) As List(Of KnowledgeSearchResult)
            Dim results As New List(Of KnowledgeSearchResult)()
            Dim queryTerms = Tokenize(query)
            If queryTerms.Length = 0 Then Return results

            ' Standard BM25 Tuning constants
            Dim k1 As Double = 1.5
            Dim b As Double = 0.75

            Dim docLengths As New List(Of Integer)(records.Count)
            Dim avgDocLength As Double = 0
            Dim termDocFrequency As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim parsedDocs As New List(Of String())()

            ' Pre-calculate Document Lengths and Global Term frequencies
            For Each r In records
                Dim terms = Tokenize(r.TextChunk)
                parsedDocs.Add(terms)
                docLengths.Add(terms.Length)
                avgDocLength += terms.Length

                Dim uniqueTerms = terms.Distinct(StringComparer.OrdinalIgnoreCase)
                For Each term In uniqueTerms
                    If Not termDocFrequency.ContainsKey(term) Then termDocFrequency(term) = 0
                    termDocFrequency(term) += 1
                Next
            Next
            avgDocLength /= records.Count
            If avgDocLength = 0 Then avgDocLength = 1

            ' Calculate Inverse Document Frequency (IDF) for query terms
            Dim idf As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
            Dim N As Double = records.Count
            For Each term In queryTerms
                Dim nq As Double = If(termDocFrequency.ContainsKey(term), termDocFrequency(term), 0)
                Dim idfVal As Double = Math.Log((N - nq + 0.5) / (nq + 0.5) + 1)
                idf(term) = idfVal
            Next

            ' Score all documents against the query
            For i As Integer = 0 To records.Count - 1
                Dim score As Double = 0
                Dim docTerms = parsedDocs(i)
                Dim Ld As Double = docLengths(i)

                For Each term In queryTerms
                    If idf.ContainsKey(term) AndAlso idf(term) > 0 Then
                        Dim tf As Double = docTerms.Count(Function(t) String.Equals(t, term, StringComparison.OrdinalIgnoreCase))
                        If tf > 0 Then
                            Dim numerator = tf * (k1 + 1)
                            Dim denominator = tf + k1 * (1 - b + b * (Ld / avgDocLength))
                            score += idf(term) * (numerator / denominator)
                        End If
                    End If
                Next

                If score > 0 Then
                    results.Add(New KnowledgeSearchResult With {
                        .FilePath = records(i).FilePath,
                        .TextChunk = records(i).TextChunk,
                        .Score = score
                    })
                End If
            Next

            Return results.OrderByDescending(Function(x) x.Score).Take(topK).ToList()
        End Function

        Private Shared Function RemoveFrontMatter(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return ""

            Dim normalized = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            If Not normalized.StartsWith("---" & vbLf, StringComparison.Ordinal) Then
                Return text
            End If

            Dim endMarker = vbLf & "---" & vbLf
            Dim endPos = normalized.IndexOf(endMarker, 4, StringComparison.Ordinal)
            If endPos < 0 Then
                Return text
            End If

            Return normalized.Substring(endPos + endMarker.Length).Replace(vbLf, vbCrLf)
        End Function


    End Class
End Namespace