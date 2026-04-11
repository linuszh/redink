' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeWikiService.vb
' Purpose: Orchestrates the physical Wiki directory structure, Markdown page
'          generation, indexing updates, and Agentic ingestion loops.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Public Class KnowledgeWikiService

        Public Const MaxChunkChars As Integer = 500000

        ' =====================================================================
        ' FOLDER & INDEX MANAGEMENT
        ' =====================================================================

        ''' <summary>
        ''' Ensures the standard Obsidian-compatible directory structure exists for a KB.
        ''' </summary>
        Public Shared Sub InitializeWikiStructure(kbRootPath As String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return

            Dim pagesPath = Path.Combine(kbRootPath, ".redink", "Wiki")
            Dim rawPath = Path.Combine(kbRootPath, ".redink", "Raw")

            If Not Directory.Exists(pagesPath) Then Directory.CreateDirectory(pagesPath)
            If Not Directory.Exists(rawPath) Then Directory.CreateDirectory(rawPath)

            ' Create master index if missing
            Dim indexPath = Path.Combine(pagesPath, "index.md")
            If Not File.Exists(indexPath) Then
                Dim defaultIndex = "# Knowledge Base Index" & vbCrLf & vbCrLf &
                                   "> Welcome to the Wiki. Auto-generated pages and concept summaries are linked below." & vbCrLf & vbCrLf &
                                   "## Concept Pages" & vbCrLf
                File.WriteAllText(indexPath, defaultIndex, System.Text.Encoding.UTF8)
            End If

            ' Create master chronological log if missing
            Dim logPath = Path.Combine(pagesPath, "log.md")
            If Not File.Exists(logPath) Then
                Dim defaultLog = "# Wiki Audit Log" & vbCrLf & vbCrLf &
                                 "> Chronological record of all ingestions, updates, and interactions." & vbCrLf & vbCrLf
                File.WriteAllText(logPath, defaultLog, System.Text.Encoding.UTF8)
            End If
        End Sub

        ''' <summary>
        ''' Safely appends a new link to the master index.md file.
        ''' </summary>
        Private Shared Sub AppendToIndex(kbRootPath As String, fileName As String, title As String, summary As String)
            Dim indexPath = Path.Combine(kbRootPath, ".redink", "Wiki", "index.md")
            If Not File.Exists(indexPath) Then InitializeWikiStructure(kbRootPath)

            Dim linkEntry = $"- **[[{fileName}]]** ({title}): {summary}" & vbCrLf
            File.AppendAllText(indexPath, linkEntry, System.Text.Encoding.UTF8)
        End Sub

        ''' <summary>
        ''' Appends a chronological entry to the log.md file.
        ''' </summary>
        Private Shared Sub AppendToLog(kbRootPath As String, action As String, details As String)
            Dim logPath = Path.Combine(kbRootPath, ".redink", "Wiki", "log.md")
            If Not File.Exists(logPath) Then InitializeWikiStructure(kbRootPath)

            Dim datePrefix = DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            Dim logEntry = $"## [{datePrefix}] {action} | {details}" & vbCrLf
            File.AppendAllText(logPath, logEntry, System.Text.Encoding.UTF8)
        End Sub

        ' =====================================================================
        ' AGENTIC WIKI INGESTION & CREATION
        ' =====================================================================


        ''' <summary>
        ''' Universal Clipboard Save: Formats clipboard text into a rich Markdown wiki page,
        ''' extracts metadata, saves it, and updates the vector database.
        ''' </summary>
        Public Shared Async Function CreatePageFromClipboardAsync(kbRootPath As String, clipboardText As String, context As ISharedContext) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(kbRootPath) OrElse String.IsNullOrWhiteSpace(clipboardText) Then Return False

            InitializeWikiStructure(kbRootPath)

            Dim indexPath = Path.Combine(kbRootPath, ".redink", "Wiki", "index.md")
            Dim wikiIndex As String = ""
            If File.Exists(indexPath) Then
                wikiIndex = File.ReadAllText(indexPath)
                If wikiIndex.Length > 15000 Then wikiIndex = wikiIndex.Substring(0, 15000) & "... [truncated]"
            End If

            Dim systemPrompt = "You are a Wiki Assistant. Formulate the user's raw text into a clean, well-structured Markdown (.md) wiki page. " &
                               "Provide a concise Title on the very first line prefixed with 'TITLE: '. " &
                               "Provide a 1-sentence summary on the second line prefixed with 'SUMMARY: '. " &
                               "Start the actual markdown content on the fourth line."

            If Not String.IsNullOrWhiteSpace(wikiIndex) Then
                systemPrompt &= vbCrLf & vbCrLf & "CURRENT WIKI STRUCTURE / INDEX:" & vbCrLf & wikiIndex & vbCrLf &
                                "IMPORTANT: Analyze the above index. Connect concepts by creating meaningful Markdown links (e.g., [[Concept Name]]) to existing pages whenever related topics appear in the raw text."
            End If

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
                LogWikiError(kbRootPath, "Clipboard", "LLM returned empty or failed to process (Timeout or context limits).")
                Return False
            End If

            Return Await ParseAndSaveAgentResponseAsync(kbRootPath, responseString, context)
        End Function

        ''' <summary>
        ''' The Agent Loop Entry: Processes a newly dropped Raw Source, synthesizes it into 
        ''' a structured Wiki Concept page, and updates the master Index.
        ''' </summary>
        Public Shared Async Function IngestSourceAsync(kbRootPath As String, sourceFilePath As String, context As ISharedContext) As Task(Of Boolean)
            If Not File.Exists(sourceFilePath) Then Return False

            InitializeWikiStructure(kbRootPath)

            Dim sourceText As String = ""
            Try
                sourceText = File.ReadAllText(sourceFilePath)
            Catch ex As Exception
                LogWikiError(kbRootPath, sourceFilePath, $"File could not be read: {ex.Message}")
                Return False
            End Try

            Dim indexPath = Path.Combine(kbRootPath, ".redink", "Wiki", "index.md")
            Dim wikiIndex As String = ""
            If File.Exists(indexPath) Then
                wikiIndex = File.ReadAllText(indexPath)
                If wikiIndex.Length > 15000 Then wikiIndex = wikiIndex.Substring(0, 15000) & "... [truncated]"
            End If

            Dim systemPrompt = "You are an autonomous Wiki Maintainer. Read the following source document. " &
                               "Synthesize its key facts into a comprehensive Markdown wiki page. " &
                               "Provide a concise Title on the very first line prefixed with 'TITLE: '. " &
                               "Provide a short summary on the second line prefixed with 'SUMMARY: '. " &
                               "Start the body of the wiki page on the fourth line. Use headings, bullet points, and bold text."

            If Not String.IsNullOrWhiteSpace(wikiIndex) Then
                systemPrompt &= vbCrLf & vbCrLf & "CURRENT WIKI INVENTORY / INDEX:" & vbCrLf & wikiIndex & vbCrLf &
                                "IMPORTANT: You MUST connect this new information to the existing knowledge base. " &
                                "Use the provided inventory to identify existing concepts and create Markdown links (e.g., [[Existing Concept Name]]) directly within the text. If the source contradicts or strongly aligns with an existing concept, please mention the relationship."
            End If

            Dim responseString As String = Await ExecuteKnowledgeStoreScopedAsync(
                context,
                Async Function(useAlternateAPI)
                    If sourceText.Length <= MaxChunkChars Then
                        Return Await SharedMethods.LLM(
                            context:=context,
                            promptSystem:=systemPrompt,
                            promptUser:=$"Source File: {Path.GetFileName(sourceFilePath)}{vbCrLf}{vbCrLf}{sourceText}",
                            UseSecondAPI:=useAlternateAPI,
                            Hidesplash:=True)
                    End If

                    Return Await ProcessLargeTextInChunksAsync(context, systemPrompt, sourceText, useAlternateAPI, Path.GetFileName(sourceFilePath))
                End Function)

            If String.IsNullOrWhiteSpace(responseString) Then
                LogWikiError(kbRootPath, sourceFilePath, "LLM returned empty or failed to process (Possibly too large or timeout).")
                Return False
            End If

            Return Await ParseAndSaveAgentResponseAsync(kbRootPath, responseString, context)
        End Function

        ''' <summary>
        ''' Breaks massively large documents into chunks, passing previous context forward for a unified summary.
        ''' </summary>
        Public Shared Async Function ProcessLargeTextInChunksAsync(context As ISharedContext, systemPrompt As String, fullText As String, useAlternateAPI As Boolean, Optional fileName As String = "Document") As Task(Of String)
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

                ' The final chunk uses the strict output formatting requirements
                Dim iterPrompt = If(i = chunks.Count - 1, systemPrompt, "You are a rolling document summarizer. Summarize the following document chunk, combining it with the previous context provided. Output ONLY the updated running comprehensive summary without required strict markdown output formatting yet.")

                rollingSummary = Await SharedMethods.LLM(
                    context:=context,
                    promptSystem:=iterPrompt,
                    promptUser:=chunkPrompt,
                    UseSecondAPI:=useAlternateAPI,
                    Hidesplash:=True)
            Next

            Return rollingSummary
        End Function

        ''' <summary>
        ''' Logs errors locally in the KB.
        ''' </summary>
        Public Shared Sub LogWikiError(kbRootPath As String, source As String, errorMessage As String)
            Try
                Dim errorLogPath = Path.Combine(kbRootPath, ".redink", "knowledge_errors.log")
                Dim msg = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Error in {Path.GetFileName(source)}: {errorMessage}{vbCrLf}"
                File.AppendAllText(errorLogPath, msg, System.Text.Encoding.UTF8)
            Catch
            End Try
        End Sub



        ' =====================================================================
        ' UTILITIES
        ' =====================================================================

        ''' <summary>
        ''' Runs a KnowledgeStore operation within a temporary alternate-model scope.
        ''' Falls back to the primary model if no special model is available and always restores
        ''' the original configuration afterwards.
        ''' </summary>
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

                Return Await operation(useAlternateAPI)
            Finally
                If backupConfig IsNot Nothing Then
                    SharedMethods.RestoreDefaults(context, backupConfig)
                End If
            End Try
        End Function

        ''' <summary>
        ''' Executes an LLM call for KnowledgeStore tasks using the alternate model if configured.
        ''' Falls back to the primary model if no special model is available and always restores
        ''' the original configuration afterwards.
        ''' </summary>
        Private Shared Async Function ExecuteKnowledgeStoreLlmAsync(context As ISharedContext,
                                                                    promptSystem As String,
                                                                    promptUser As String,
                                                                    Optional hideSplash As Boolean = True) As Task(Of String)
            Return Await ExecuteKnowledgeStoreScopedAsync(
                context,
                Async Function(useAlternateAPI)
                    Return Await SharedMethods.LLM(
                        context:=context,
                        promptSystem:=promptSystem,
                        promptUser:=promptUser,
                        UseSecondAPI:=useAlternateAPI,
                        Hidesplash:=hideSplash)
                End Function)
        End Function

        ''' <summary>
        ''' Parses the Agent's structured output, writes the physical .md file, 
        ''' updates the index, and queues the file for Vector Embedding.
        ''' </summary>
        Private Shared Async Function ParseAndSaveAgentResponseAsync(kbRootPath As String, agentResponse As String, context As ISharedContext) As Task(Of Boolean)
            Dim lines = agentResponse.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None)

            Dim title As String = "Untitled Concept"
            Dim summary As String = "Auto-generated wiki page."
            Dim contentStartIdx As Integer = 0

            For i As Integer = 0 To Math.Min(lines.Length - 1, 5)
                Dim line = lines(i).Trim()
                If line.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                    title = line.Substring(6).Trim()
                ElseIf line.StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) Then
                    summary = line.Substring(8).Trim()
                ElseIf String.IsNullOrWhiteSpace(line) AndAlso contentStartIdx = 0 Then
                    ' Assuming content starts after blanks
                Else
                    If contentStartIdx = 0 AndAlso Not line.StartsWith("TITLE:") AndAlso Not line.StartsWith("SUMMARY:") Then
                        contentStartIdx = i
                    End If
                End If
            Next

            Dim bodyContent = String.Join(vbCrLf, lines.Skip(contentStartIdx))

            ' Sanitize a filename out of the title
            Dim safeName = Regex.Replace(title, "[^a-zA-Z0-9_\-\s]", "").Trim()
            If String.IsNullOrWhiteSpace(safeName) Then safeName = "WikiPage_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
            Dim fileName = safeName & ".md"
            Dim fullPath = Path.Combine(kbRootPath, ".redink", "Wiki", fileName)

            ' WIKI INCREMENTAL MERGE LOGIC:
            ' If the topic already exists, do not duplicate (no _1, _2). Merge new information gracefully.
            Dim isUpdate As Boolean = False
            If File.Exists(fullPath) Then
                isUpdate = True
                Dim existingContent = File.ReadAllText(fullPath, System.Text.Encoding.UTF8)

                Dim mergeSystemPrompt = "You are an autonomous Wiki Maintainer. Merge the new extracted knowledge into the existing wiki page. " &
                                        "Synthesize smoothly, resolving contradictions, aligning facts, and eliminating redundancies. Retain structural elements and links." & vbCrLf &
                                        "Provide a concise Title on the very first line prefixed with 'TITLE: '." & vbCrLf &
                                        "Provide a 1-sentence summary on the second line prefixed with 'SUMMARY: '." & vbCrLf &
                                        "Start the actual markdown content on the fourth line."

                Dim mergeUserPrompt = $"=== EXISTING WIKI PAGE ==={vbCrLf}{existingContent}{vbCrLf}=== NEW EXTRACTED KNOWLEDGE ==={vbCrLf}{bodyContent}"

                Dim mergedResponse = Await ExecuteKnowledgeStoreLlmAsync(
                    context:=context,
                    promptSystem:=mergeSystemPrompt,
                    promptUser:=mergeUserPrompt,
                    hideSplash:=True)

                If Not String.IsNullOrWhiteSpace(mergedResponse) Then
                    ' Reparse merged result
                    lines = mergedResponse.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None)
                    contentStartIdx = 0
                    For i As Integer = 0 To Math.Min(lines.Length - 1, 5)
                        Dim line = lines(i).Trim()
                        If line.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase) Then
                            title = line.Substring(6).Trim()
                        ElseIf line.StartsWith("SUMMARY:", StringComparison.OrdinalIgnoreCase) Then
                            summary = line.Substring(8).Trim()
                        ElseIf String.IsNullOrWhiteSpace(line) AndAlso contentStartIdx = 0 Then
                        Else
                            If contentStartIdx = 0 AndAlso Not line.StartsWith("TITLE:") AndAlso Not line.StartsWith("SUMMARY:") Then
                                contentStartIdx = i
                            End If
                        End If
                    Next
                    bodyContent = String.Join(vbCrLf, lines.Skip(contentStartIdx))
                End If
            End If

            Try
                ' 1. Save the file to the Wiki Folder (overwrites merged or saves new)
                File.WriteAllText(fullPath, bodyContent, System.Text.Encoding.UTF8)

                ' 2. Append link to Index.md and Chronicle the update in log.md
                If Not isUpdate Then
                    AppendToIndex(kbRootPath, Path.GetFileNameWithoutExtension(fileName), title, summary)
                    AppendToLog(kbRootPath, "ingest", title)
                Else
                    AppendToLog(kbRootPath, "update", title)
                End If

                ' 3. Index it into our Vector Database for Semantic Search
                Await KnowledgeEmbeddingService.UpdateFileEmbeddingsAsync(kbRootPath, fullPath, bodyContent, context)

                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"Wiki Save Error: {ex.Message}")
                Return False
            End Try
        End Function

        ' =====================================================================
        ' WIKI LINTING & HEALTH CHECK
        ' =====================================================================

        ''' <summary>
        ''' Performs a health check on the Wiki, finding orphaned pages and asking the LLM
        ''' to review the index for inconsistencies, missing links, or structural improvements.
        ''' </summary>
        Public Shared Async Function LintWikiAsync(kbRootPath As String, context As ISharedContext) As Task(Of String)
            If String.IsNullOrWhiteSpace(kbRootPath) Then Return "Invalid KB path."

            ' FIXED: Proper path matching the defined Wiki structure
            Dim pagesPath = Path.Combine(kbRootPath, ".redink", "Wiki")
            Dim indexPath = Path.Combine(pagesPath, "index.md")

            If Not Directory.Exists(pagesPath) OrElse Not File.Exists(indexPath) Then
                Return "Wiki is empty or not initialized."
            End If

            ' 1. Detect Orphaned Pages via physical file scan
            Dim allMdFiles = Directory.GetFiles(pagesPath, "*.md")
            Dim indexContent = File.ReadAllText(indexPath, System.Text.Encoding.UTF8)
            Dim orphans As New List(Of String)()

            For Each file In allMdFiles
                Dim fileName = Path.GetFileName(file)
                If fileName.Equals("index.md", StringComparison.OrdinalIgnoreCase) Then Continue For
                If fileName.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) Then Continue For

                ' Check if the filename or title is referenced in the index
                If Not indexContent.Contains(fileName) AndAlso Not indexContent.Contains(Path.GetFileNameWithoutExtension(fileName)) Then
                    orphans.Add(fileName)
                End If
            Next

            ' 2. Gather context for the Health Report
            Dim reportBuilder As New System.Text.StringBuilder()
            reportBuilder.AppendLine("# Wiki Health Report")
            reportBuilder.AppendLine($"**Date Generated:** {DateTime.Now.ToString("yyyy-MM-dd HH:mm")}")
            reportBuilder.AppendLine($"**Total Pages:** {allMdFiles.Length - 1}")

            If orphans.Count > 0 Then
                reportBuilder.AppendLine()
                reportBuilder.AppendLine("## ⚠️ Orphaned Pages Detected")
                reportBuilder.AppendLine("These physical files exist but are missing from `index.md`:")
                For Each orphan In orphans
                    reportBuilder.AppendLine($"- [[{Path.GetFileNameWithoutExtension(orphan)}]]")
                Next
            Else
                reportBuilder.AppendLine()
                reportBuilder.AppendLine("## ✅ Orphan Check")
                reportBuilder.AppendLine("No orphaned pages found. All files are properly linked in the index.")
            End If

            ' 3. LLM Structural Analysis & Contradiction Check
            ' We pass the index and the orphan status to the Agent so it can flag duplicate concepts or messy architecture
            Dim systemPrompt = "You are a Knowledge Base Health Inspector. Review the provided wiki index and orphaned files list. " &
                               "1. Identify any duplicate or heavily overlapping concepts. " &
                               "2. Flag any contradictory categorizations or missing structural links. " &
                               "3. Provide short, actionable recommendations to improve the wiki architecture."

            Dim userPrompt = $"=== CURRENT INDEX CONTENT ==={vbCrLf}{indexContent}{vbCrLf}=== ORPHANS DETECTED ==={vbCrLf}{String.Join(", ", orphans)}"

            Dim llmAnalysis = Await ExecuteKnowledgeStoreLlmAsync(
                context:=context,
                promptSystem:=systemPrompt,
                promptUser:=userPrompt,
                hideSplash:=False)

            If Not String.IsNullOrWhiteSpace(llmAnalysis) Then
                reportBuilder.AppendLine()
                reportBuilder.AppendLine("## 🧠 AI Structural Analysis")
                reportBuilder.AppendLine(llmAnalysis)
            End If

            ' 4. Save report dynamically into the Wiki
            Dim reportPath = Path.Combine(pagesPath, "health_report.md")
            File.WriteAllText(reportPath, reportBuilder.ToString(), System.Text.Encoding.UTF8)

            Return reportBuilder.ToString()
        End Function


    End Class
End Namespace