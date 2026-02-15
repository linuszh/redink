' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.vb
' Purpose: Defines and executes AutoPilot-specific internal tools for document
'          processing, attachment handling, and binary analysis. These tools are
'          registered alongside existing external tools and the built-in web
'          retrieval tool.
'
' Tools Provided:
'  - process_word_document: Applies any prompt (translate, correct, freestyle) to
'    DOCX attachments using the ported OpenXML paragraph-batch processing.
'  - comment_word_document: Adds Word comment bubbles to DOCX attachments using a
'    review instruction and optional author name.
'  - extract_pdf_text: Extracts text content from PDF attachments, optionally
'    running OCR when needed and available.
'  - merge_pdfs: Merges multiple PDF attachments into a single PDF.
'  - read_attachment: Reads the text content of any supported attachment
'    (DOCX, PDF, TXT, CSV, HTML, XML, JSON, XLSX, XLS, PPTX).
'  - list_attachments: Lists all attachments with metadata.
'  - describe_binary_attachment: Sends a supported binary attachment (image, audio,
'    video) to the AI for description or transcription.
'  - compare_word_documents: Compares two Word document attachments using Word's
'    built-in comparison and produces a tracked-changes document.
'
' Architecture:
'  - Tools are registered as ModelConfig entries with ToolOnly=True, consistent
'    with the existing tool loading pattern from specialservices INI files.
'  - Tool execution is routed through the existing ExecuteToolCall method which
'    checks for internal tools before calling external endpoints.
'  - Binary analysis is enabled only when the API configuration supports file
'    objects, and is restricted to specific extensions.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Tool name for binary attachment description or transcription.
    ''' </summary>
    Private Const AP_Tool_DescribeBinary As String = "describe_binary_attachment"

    ''' <summary>
    ''' Supported binary extensions that can be passed as FileObject to the LLM.
    ''' </summary>
    Private Shared ReadOnly AP_BinaryExtensions As String() = {
        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tiff", ".tif",
        ".mp3", ".wav", ".ogg", ".m4a", ".flac", ".aac", ".wma",
        ".mp4", ".avi", ".mov", ".mkv", ".webm"
    }

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL NAMES (constants for matching)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Tool name for Word document processing.
    ''' </summary>
    Private Const AP_Tool_ProcessWordDoc As String = "process_word_document"

    ''' <summary>
    ''' Tool name for adding comment bubbles to Word documents.
    ''' </summary>
    Private Const AP_Tool_CommentWordDoc As String = "comment_word_document"

    ''' <summary>
    ''' Tool name for extracting text from PDF attachments.
    ''' </summary>
    Private Const AP_Tool_ExtractPdfText As String = "extract_pdf_text"

    ''' <summary>
    ''' Tool name for merging multiple PDF attachments.
    ''' </summary>
    Private Const AP_Tool_MergePdfs As String = "merge_pdfs"

    ''' <summary>
    ''' Tool name for reading attachment content.
    ''' </summary>
    Private Const AP_Tool_ReadAttachment As String = "read_attachment"

    ''' <summary>
    ''' Tool name for listing attachments.
    ''' </summary>
    Private Const AP_Tool_ListAttachments As String = "list_attachments"

    ''' <summary>
    ''' Prefix used for AutoPilot tool naming.
    ''' </summary>
    Private Const AP_ToolPrefix As String = "autopilot_"

    ''' <summary>
    ''' Tool name for comparing Word documents.
    ''' </summary>
    Private Const AP_Tool_CompareWordDocs As String = "compare_word_documents"

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL REGISTRATION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Creates ModelConfig entries for all AutoPilot internal tools.
    ''' These integrate seamlessly with the existing tool selection and execution system.
    ''' </summary>
    Friend Function GetAutoPilotInternalTools() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        ' ── process_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_ProcessWordDoc,
            .ModelDescription = "Process Word Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ProcessWordDoc & ": Processes one or more Word document (.docx) attachments by applying a prompt/instruction. " &
                "Use this for translation, correction, proofreading, anonymization, or any text transformation. " &
                "Returns both a clean version and a compare document showing changes.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ProcessWordDoc & """," &
                """description"":""Applies a text processing instruction to Word document attachments. Supports translation, correction, anonymization, and freestyle text operations. Produces clean output and a compare document with tracked changes.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""The instruction to apply to the document (e.g., 'Translate to German', 'Correct spelling and grammar', 'Anonymize all personal names')""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the Word document attachments to process. If empty or omitted, processes all .docx attachments.""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── extract_pdf_text ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_ExtractPdfText,
            .ModelDescription = "Extract PDF Text (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ExtractPdfText & ": Extracts the text content from one or more PDF attachments.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExtractPdfText & """," &
                """description"":""Extracts text from PDF file attachments""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to extract text from. If empty, processes all PDFs.""}" &
                "},""required"":[]}}"
        })

        ' ── merge_pdfs ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_MergePdfs,
            .ModelDescription = "Merge PDFs (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_MergePdfs & ": Merges multiple PDF attachments into a single PDF file.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_MergePdfs & """," &
                """description"":""Merges multiple PDF file attachments into a single combined PDF""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to merge, in order. If empty, merges all PDFs.""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the merged output PDF (default: merged.pdf)""}" &
                "},""required"":[]}}"
        })

        ' ── read_attachment ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_ReadAttachment,
            .ModelDescription = "Read Attachment Content (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReadAttachment & ": Reads and returns the text content of a supported attachment " &
                "(DOCX, PDF, TXT, CSV, HTML, XML, JSON, XLSX, XLS, PPTX).",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReadAttachment & """," &
                """description"":""Reads and returns the text content of an attachment file. " &
                "Supports Word documents (.docx), PDFs (.pdf), Excel spreadsheets (.xlsx, .xls), " &
                "PowerPoint presentations (.pptx), and text-based files (.txt, .csv, .html, .xml, .json, .md, .log).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the attachment to read""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── list_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_ListAttachments,
            .ModelDescription = "List Attachments (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ListAttachments & ": Lists all attachments of the current email with name, type, and size.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ListAttachments & """," &
                """description"":""Lists all email attachments with their filename, type, size, and processing status""," &
                """parameters"":{""type"":""object"",""properties"":{},""required"":[]}}"
        })

        ' ── describe_binary_attachment ───────────────────────────────────
        ' Only add if the model supports file objects (APICall_Object is configured)
        Dim apiCallObj As String = If(_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi,
                                      INI_APICall_Object_2, INI_APICall_Object)
        If Not String.IsNullOrWhiteSpace(apiCallObj) Then
            tools.Add(New ModelConfig() With {
                .ToolOnly = True,
                .Tool = True,
                .ToolName = AP_Tool_DescribeBinary,
                .ModelDescription = "Describe or transcribe a binary attachment (built-in)",
                .ToolInstructionsPrompt =
                    AP_Tool_DescribeBinary & ": Sends a binary attachment (image, audio, voicemail, video) to the AI for description or transcription. " &
                    "Do NOT use this for every image — only when the user explicitly asks about an attachment or when it appears to be substantive content. " &
                    "Common footer/signature images should be ignored.",
                .ToolDefinition =
                    "{""name"":""" & AP_Tool_DescribeBinary & """," &
                    """description"":""Sends a binary attachment (image, audio file, voicemail, video) directly to the AI for description, transcription, or analysis. " &
                    "Use for .png, .jpg, .mp3, .wav, .m4a, .mp4, etc. Do NOT use for footer/signature images.""," &
                    """parameters"":{""type"":""object"",""properties"":{" &
                    """attachment_name"":{""type"":""string"",""description"":""Filename of the binary attachment to analyze""}," &
                    """prompt"":{""type"":""string"",""description"":""Instructions for the AI (e.g. 'describe this image', 'transcribe this voicemail')""}" &
                    "},""required"":[""attachment_name"",""prompt""]}}"
            })
        End If

        ' ── comment_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_CommentWordDoc,
            .ModelDescription = "Comment Word Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CommentWordDoc & ": Adds review comments (Word comment bubbles) to a Word (.docx) attachment. " &
                "Use this when the user wants the document annotated, commented, reviewed with margin notes, " &
                "or marked up with feedback directly inside the document as Word comment bubbles. " &
                "Do NOT use this when the user wants a textual summary or analysis — only when comments " &
                "should appear as annotations within the document itself. " &
                "Supports an optional author parameter: if the user asks for comments under a specific name " &
                "(e.g. the sender's name), pass it as author. If not specified, comments are authored as 'Inky'.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CommentWordDoc & """," &
                """description"":""Adds review comments as Word comment bubbles to a .docx attachment. " &
                "Use when the user wants in-document annotations, margin comments, or review feedback placed directly " &
                "inside the Word file. Do NOT use for plain textual summaries or analyses. " &
                "Supports an optional author name for the comments.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""The review instruction (e.g., 'Review for legal risks', 'Check for inconsistencies', 'Suggest improvements')""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the .docx attachments to annotate. If empty or omitted, annotates all .docx attachments.""}," &
                """author"":{""type"":""string"",""description"":""Optional author name for the comment bubbles. Use this when the user requests a specific name (e.g. the sender's name). If omitted, defaults to 'Inky'.""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── compare_word_documents ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True,
            .Tool = True,
            .ToolName = AP_Tool_CompareWordDocs,
            .ModelDescription = "Compare two Word documents (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CompareWordDocs & ": Compares exactly two Word document (.doc/.docx) attachments using Word's " &
                "built-in comparison engine (track changes). Use 'original_filename' for the BASE/earlier/reference " &
                "version and 'revised_filename' for the MODIFIED/newer/changed version. Returns a textual summary of " &
                "revisions found and produces a comparison document with tracked changes attached to the reply. " &
                "This tool requires exactly two attachments to be specified — it cannot compare more than two at once. " &
                "If the sender provides more than two Word documents, ask which two should be compared, or run " &
                "multiple comparisons.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CompareWordDocs & """," &
                """description"":""Compares exactly two Word documents (.doc/.docx) using Word's built-in comparison (track changes). " &
                "The 'original_filename' is the BASE document (the earlier or reference version). " &
                "The 'revised_filename' is the MODIFIED document (the newer or changed version). " &
                "Returns a textual summary of the differences found and produces a comparison document with tracked changes as a result attachment. " &
                "IMPORTANT: 'original_filename' = the source/baseline; 'revised_filename' = the version that was changed or updated. " &
                "This tool compares exactly two documents per call.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """original_filename"":{""type"":""string"",""description"":""Exact filename of the original/baseline/source Word attachment (the earlier version)""}," &
                """revised_filename"":{""type"":""string"",""description"":""Exact filename of the revised/modified/updated Word attachment (the newer version)""}" &
                "},""required"":[""original_filename"",""revised_filename""]}}"
        })

        Return tools
    End Function

    ''' <summary>
    ''' Checks if a tool call is an AutoPilot internal tool and executes it.
    ''' Returns Nothing if the tool is not an AutoPilot tool (lets the existing
    ''' external tool pipeline handle it).
    ''' This method is called from the existing ExecuteToolCall pathway.
    ''' </summary>
    Friend Async Function TryExecuteAutoPilotTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            Optional cancellationToken As CancellationToken = Nothing) As Task(Of ToolResponse)

        Select Case toolCall.ToolName

            Case AP_Tool_ProcessWordDoc
                Return Await ExecuteProcessWordDocTool(toolCall, context, cancellationToken)

            Case AP_Tool_ExtractPdfText
                Return Await ExecuteExtractPdfTextTool(toolCall, context, cancellationToken)

            Case AP_Tool_MergePdfs
                Return Await ExecuteMergePdfsTool(toolCall, context, cancellationToken)

            Case AP_Tool_ReadAttachment
                Return Await ExecuteReadAttachmentTool(toolCall, context, cancellationToken)

            Case AP_Tool_ListAttachments
                Return ExecuteListAttachmentsTool(toolCall, context)

            Case AP_Tool_DescribeBinary
                Return Await ExecuteDescribeBinaryTool(toolCall, context, cancellationToken)

            Case AP_Tool_CommentWordDoc
                Return Await ExecuteCommentWordDocTool(toolCall, context, cancellationToken)

            Case AP_Tool_CompareWordDocs
                Return Await ExecuteCompareWordDocsTool(toolCall, context, cancellationToken)

            Case Else
                Return Nothing ' Not our tool — let existing pipeline handle it
        End Select
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: comment_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the comment_word_document tool against matching Word attachments.
    ''' </summary>
    Private Async Function ExecuteCommentWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instructionObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("instruction") Then instructionObj = toolCall.Arguments("instruction")
            Dim instruction = instructionObj?.ToString()
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Extract optional author parameter
            Dim authorObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("author") Then authorObj = toolCall.Arguments("author")
            Dim author As String = authorObj?.ToString()

            ' Determine which attachments to process
            Dim targetNames As New List(Of String)()
            If toolCall.Arguments?.ContainsKey("attachment_names") Then
                Dim namesObj = toolCall.Arguments("attachment_names")
                If TypeOf namesObj Is JArray Then
                    For Each item In DirectCast(namesObj, JArray)
                        targetNames.Add(item.ToString())
                    Next
                End If
            End If

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) (a.Extension = ".docx" OrElse a.Extension = ".doc") AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word document attachments found."
                Return response
            End If

            Dim effectiveAuthor = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
            Dim authorNote = If(effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase),
                                "", $" (author: {effectiveAuthor})")

            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Adding comments to: {att.OriginalFileName} with instruction: {instruction}{authorNote}")
                ApDashboardLog($"💬 Adding comments to: {att.OriginalFileName}{authorNote}", "step")

                Dim inputPath = att.TempFilePath

                ' Only .docx is supported for OpenXML comment insertion
                If Not inputPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) Then
                    resultMessages.Add($"✗ {att.OriginalFileName}: Only .docx files are supported for comment insertion.")
                    Continue For
                End If

                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_commented.docx"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                Dim success = Await CommentDocxForAutoPilot(inputPath, outputPath, instruction, ct, author)

                If success Then
                    att.OutputFiles.Add(outputPath)
                    resultMessages.Add($"✓ {att.OriginalFileName}: Comments added successfully. Output: {outputName}")
                    ApDashboardLog($"✓ Comments added to: {att.OriginalFileName}", "info")
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Failed to add comments (document may be empty or unsupported).")
                    ApDashboardLog($"⚠ Failed to add comments to: {att.OriginalFileName}", "warn")
                End If
            Next

            response.Success = resultMessages.Any(Function(m) m.StartsWith("✓"))
            response.Response = String.Join(vbCrLf, resultMessages)

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error adding comments to Word document(s): {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: compare_word_documents
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Compares two Word document attachments using Word's Compare feature.
    ''' Produces a comparison .docx with tracked changes and returns a textual summary.
    ''' </summary>
    Private Async Function ExecuteCompareWordDocsTool(
                toolCall As ToolCall,
                context As ToolExecutionContext,
                ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow,
            .OriginalCallJson = toolCall.RawJson
        }

        Try
            Dim originalFilename As String = ""
            Dim revisedFilename As String = ""

            If toolCall.Arguments IsNot Nothing Then
                If toolCall.Arguments.ContainsKey("original_filename") Then
                    originalFilename = toolCall.Arguments("original_filename")?.ToString().Trim()
                End If
                If toolCall.Arguments.ContainsKey("revised_filename") Then
                    revisedFilename = toolCall.Arguments("revised_filename")?.ToString().Trim()
                End If
            End If

            If String.IsNullOrWhiteSpace(originalFilename) OrElse String.IsNullOrWhiteSpace(revisedFilename) Then
                response.Success = False
                response.ErrorMessage = "Both 'original_filename' and 'revised_filename' are required."
                response.Response = response.ErrorMessage
                Return response
            End If

            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count < 2 Then
                response.Success = False
                response.ErrorMessage = "At least two attachments are required for comparison."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim originalAtt = _apCurrentAttachments.FirstOrDefault(
                Function(a) a.OriginalFileName.Equals(originalFilename, StringComparison.OrdinalIgnoreCase))
            Dim revisedAtt = _apCurrentAttachments.FirstOrDefault(
                Function(a) a.OriginalFileName.Equals(revisedFilename, StringComparison.OrdinalIgnoreCase))

            If originalAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Original attachment '{originalFilename}' not found. Available: {String.Join(", ", _apCurrentAttachments.Select(Function(a) a.OriginalFileName))}"
                response.Response = response.ErrorMessage
                Return response
            End If

            If revisedAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Revised attachment '{revisedFilename}' not found. Available: {String.Join(", ", _apCurrentAttachments.Select(Function(a) a.OriginalFileName))}"
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Validate extensions
            Dim origExt = IO.Path.GetExtension(originalAtt.TempFilePath).ToLowerInvariant()
            Dim revExt = IO.Path.GetExtension(revisedAtt.TempFilePath).ToLowerInvariant()
            Dim supportedExts = {".doc", ".docx"}

            If Not supportedExts.Contains(origExt) OrElse Not supportedExts.Contains(revExt) Then
                response.Success = False
                response.ErrorMessage = "Both documents must be Word files (.doc or .docx)."
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Build comparison output path
            Dim compareName = $"Comparison_{IO.Path.GetFileNameWithoutExtension(originalFilename)}_vs_{IO.Path.GetFileNameWithoutExtension(revisedFilename)}.docx"
            Dim comparePath = IO.Path.Combine(_apCurrentTempDir, compareName)

            context.Log($"Comparing: {originalFilename} (original) vs {revisedFilename} (revised)")
            ApDashboardLog($"📊 Comparing: {originalFilename} vs {revisedFilename}", "step")

            ' Run the comparison on the UI/STA thread (Word COM requires it)
            Dim success As Boolean = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(
                originalAtt.TempFilePath, revisedAtt.TempFilePath, comparePath))

            If success AndAlso IO.File.Exists(comparePath) Then
                ' Register the comparison doc as an output file so it gets attached to the reply
                originalAtt.OutputFiles.Add(comparePath)

                ' Build a textual summary of changes by reading the comparison doc's tracked changes
                Dim summaryText As String = ""
                Try
                    summaryText = Await SwitchToUi(Function()
                                                       Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                                       Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                                       Try
                                                           wordApp = New Microsoft.Office.Interop.Word.Application() With {.Visible = False}
                                                           doc = wordApp.Documents.Open(comparePath, ReadOnly:=True)

                                                           Dim revCount = doc.Revisions.Count
                                                           Dim sb As New System.Text.StringBuilder()
                                                           sb.AppendLine($"Comparison complete: {revCount} revision(s) found between '{originalFilename}' (original) and '{revisedFilename}' (revised).")
                                                           sb.AppendLine()

                                                           Dim maxRevisions = Math.Min(revCount, 50)
                                                           For i As Integer = 1 To maxRevisions
                                                               Dim rev = doc.Revisions(i)
                                                               Dim revType = rev.Type.ToString()
                                                               Dim revText = rev.Range.Text
                                                               If revText IsNot Nothing AndAlso revText.Length > 200 Then
                                                                   revText = revText.Substring(0, 200) & "..."
                                                               End If
                                                               sb.AppendLine($"  [{revType}] {revText}")
                                                           Next

                                                           If revCount > maxRevisions Then
                                                               sb.AppendLine($"  ... and {revCount - maxRevisions} more revision(s).")
                                                           End If

                                                           Return sb.ToString()
                                                       Finally
                                                           Try : If doc IsNot Nothing Then doc.Close(False)
                                                           Catch
                                                           End Try
                                                           Try : If wordApp IsNot Nothing Then wordApp.Quit(False)
                                                           Catch
                                                           End Try
                                                       End Try
                                                   End Function)
                Catch ex As Exception
                    summaryText = $"Comparison document created successfully but could not extract revision summary: {ex.Message}"
                End Try

                response.Success = True
                response.Response = summaryText & vbCrLf & $"The comparison document '{compareName}' has been generated and will be attached to the reply."
                ApDashboardLog($"✓ Comparison complete: {compareName}", "info")
            Else
                response.Success = False
                response.ErrorMessage = "Word comparison failed. The documents may be incompatible or corrupted."
                response.Response = response.ErrorMessage
                ApDashboardLog($"⚠ Comparison failed for: {originalFilename} vs {revisedFilename}", "warn")
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = $"Error comparing documents: {ex.Message}"
            response.Response = response.ErrorMessage
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: process_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the process_word_document tool against matching Word attachments.
    ''' </summary>
    Private Async Function ExecuteProcessWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instructionObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("instruction") Then instructionObj = toolCall.Arguments("instruction")
            Dim instruction = instructionObj?.ToString()
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Determine which attachments to process
            Dim targetNames As New List(Of String)()
            If toolCall.Arguments?.ContainsKey("attachment_names") Then
                Dim namesObj = toolCall.Arguments("attachment_names")
                If TypeOf namesObj Is JArray Then
                    For Each item In DirectCast(namesObj, JArray)
                        targetNames.Add(item.ToString())
                    Next
                End If
            End If

            ' If no specific names, process all .docx attachments
            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) (a.Extension = ".docx" OrElse a.Extension = ".doc") AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word document attachments found."
                Return response
            End If

            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Processing: {att.OriginalFileName} with instruction: {instruction}")

                Dim inputPath = att.TempFilePath
                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_processed.docx"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                ' Use the ported DOCX processor
                Dim success = Await ProcessDocxForAutoPilot(inputPath, outputPath, instruction, ct)

                If success Then
                    att.OutputFiles.Add(outputPath)

                    ' Create compare document using Word interop
                    Dim comparePath = Path.Combine(_apCurrentTempDir,
                        Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_compare.docx")
                    Dim compareSuccess = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(inputPath, outputPath, comparePath))
                    If compareSuccess Then
                        att.OutputFiles.Add(comparePath)
                        resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} + compare document.")
                    Else
                        resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} (compare document creation failed).")
                    End If
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Processing failed.")
                End If
            Next

            response.Success = True
            response.Response = String.Join(vbCrLf, resultMessages)

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error processing Word document(s): {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Creates a Word compare document between original and processed files.
    ''' Must run on the UI thread (uses Word interop).
    ''' </summary>
    Private Function CreateWordCompareDocumentForAutoPilot(originalPath As String, processedPath As String, comparePath As String) As Boolean
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
        Dim originalDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim processedDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing

        Try
            ' Try to get or create a Word application instance
            Try
                wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch
                wordApp = New Microsoft.Office.Interop.Word.Application()
            End Try

            Dim wasScreenUpdating = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            processedDoc = wordApp.Documents.Open(processedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=processedDoc,
                Destination:=Microsoft.Office.Interop.Word.WdCompareDestination.wdCompareDestinationNew,
                Granularity:=Microsoft.Office.Interop.Word.WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True,
                CompareCaseChanges:=True,
                CompareWhitespace:=True,
                CompareTables:=True,
                CompareHeaders:=True,
                CompareFootnotes:=True,
                CompareTextboxes:=True,
                CompareFields:=True,
                CompareComments:=True,
                RevisedAuthor:=AN6,
                IgnoreAllComparisonWarnings:=True)

            compareDoc.SaveAs2(comparePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
            compareDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            compareDoc = Nothing

            processedDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            processedDoc = Nothing
            originalDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            originalDoc = Nothing

            wordApp.ScreenUpdating = wasScreenUpdating
            Return True

        Catch ex As Exception
            Debug.WriteLine($"CreateWordCompareDocumentForAutoPilot error: {ex.Message}")
            Return False
        Finally
            If compareDoc IsNot Nothing Then Try : compareDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            If processedDoc IsNot Nothing Then Try : processedDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            If originalDoc IsNot Nothing Then Try : originalDoc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            If wordApp IsNot Nothing Then Try : wordApp.ScreenUpdating = True : Catch : End Try
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_pdf_text
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the extract_pdf_text tool against matching PDF attachments.
    ''' </summary>
    Private Async Function ExecuteExtractPdfTextTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames As New List(Of String)()
            If toolCall.Arguments?.ContainsKey("attachment_names") Then
                Dim namesObj = toolCall.Arguments("attachment_names")
                If TypeOf namesObj Is JArray Then
                    For Each item In DirectCast(namesObj, JArray)
                        targetNames.Add(item.ToString())
                    Next
                End If
            End If

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No PDF attachments found to extract."
                Return response
            End If

            Dim sb As New System.Text.StringBuilder()
            For Each att In toProcess
                context.Log($"Extracting text from: {att.OriginalFileName}")
                ApDashboardLog($"📄 Extracting text from: {att.OriginalFileName}", "step")

                ' Step 1: Standard text extraction (no OCR yet)
                Dim pdfResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                    att.TempFilePath,
                    ReturnErrorInsteadOfEmpty:=True,
                    DoOCR:=False,
                    AskUser:=False,
                    context:=_context
                )

                Dim text As String = If(pdfResult IsNot Nothing, pdfResult.Content, "")
                Dim usedOcr As Boolean = False

                ' Step 2: Determine if OCR is needed
                Dim ocrAvailable As Boolean = SharedMethods.IsOcrAvailable(_context)
                Dim needsOcr As Boolean = False

                If String.IsNullOrWhiteSpace(text) Then
                    needsOcr = True
                ElseIf pdfResult IsNot Nothing AndAlso pdfResult.OcrWasSkippedDueToHeuristics Then
                    needsOcr = True
                End If

                If needsOcr AndAlso ocrAvailable Then
                    ApDashboardLog($"🔍 Running OCR on: {att.OriginalFileName}", "step")
                    context.Log($"OCR: {att.OriginalFileName}")

                    Dim ocrResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                        att.TempFilePath,
                        ReturnErrorInsteadOfEmpty:=True,
                        DoOCR:=True,
                        AskUser:=False,
                        context:=_context
                    )

                    Dim ocrText As String = If(ocrResult IsNot Nothing, ocrResult.Content, "")

                    If Not String.IsNullOrWhiteSpace(ocrText) Then
                        text = ocrText
                        usedOcr = True
                        ApDashboardLog($"✓ OCR completed for: {att.OriginalFileName} ({ocrText.Length:N0} chars)", "info")
                    Else
                        ApDashboardLog($"⚠ OCR returned no content for: {att.OriginalFileName}, using standard extraction", "warn")
                    End If
                End If

                sb.AppendLine($"[{att.OriginalFileName}]")
                If String.IsNullOrWhiteSpace(text) Then
                    Dim ocrNote = If(Not ocrAvailable,
                        "(no extractable text; OCR is not available in the current configuration)",
                        "(no extractable text)")
                    sb.AppendLine(ocrNote)
                    ApDashboardLog($"⚠ No text extracted from: {att.OriginalFileName}", "warn")
                Else
                    If Not usedOcr AndAlso pdfResult IsNot Nothing AndAlso pdfResult.OcrWasSkippedDueToHeuristics AndAlso Not ocrAvailable Then
                        sb.AppendLine("(Note: This PDF may contain scanned images. Some content may be missing because OCR is not available.)")
                    End If
                    sb.AppendLine(text)
                End If
                sb.AppendLine()
            Next

            response.Success = True
            response.Response = sb.ToString().TrimEnd()

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting PDF text: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: describe_binary_attachment
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the describe_binary_attachment tool for a supported binary attachment.
    ''' </summary>
    Private Async Function ExecuteDescribeBinaryTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileNameObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("attachment_name") Then fileNameObj = toolCall.Arguments("attachment_name")
            Dim fileName = fileNameObj?.ToString()

            Dim promptObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("prompt") Then promptObj = toolCall.Arguments("prompt")
            Dim prompt = If(promptObj?.ToString(), "Describe or transcribe this file.")

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = _apCurrentAttachments?.FirstOrDefault(
                Function(a) a.OriginalFileName.Equals(fileName, StringComparison.OrdinalIgnoreCase))

            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If att.IsOverSizeLimit Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' exceeds the size limit."
                Return response
            End If

            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' could not be read."
                Return response
            End If

            ' Check extension is a supported binary type
            Dim ext As String = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If Not AP_BinaryExtensions.Contains(ext) Then
                response.Success = False
                response.Response = $"The file format '{ext}' is not supported for binary analysis. " &
                    "Supported formats include images (.png, .jpg, .gif, .webp, .tiff), " &
                    "audio (.mp3, .wav, .ogg, .m4a, .flac, .aac), and video (.mp4, .mov, .webm)."
                Return response
            End If

            context.Log($"Analyzing binary attachment: {fileName}")
            ApDashboardLog($"🖼 Sending binary attachment to AI: {fileName} ({ext})", "step")

            ' Determine which API to use
            Dim useSecond As Boolean = (_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi)

            ' Call LLM with the file as FileObject
            Dim llmResult As String = Await SharedMethods.LLM(
                _context,
                prompt,
                "",
                UseSecondAPI:=useSecond,
                Hidesplash:=True,
                FileObject:=att.TempFilePath,
                cancellationToken:=ct
            )

            If String.IsNullOrWhiteSpace(llmResult) Then
                response.Success = False
                response.Response = $"The AI model could not process the file '{fileName}'. The model may not support this file type."
                ApDashboardLog($"⚠ Binary analysis returned no result for: {fileName}", "warn")
                Return response
            End If

            response.Success = True
            response.Response = $"Analysis of '{fileName}':" & vbCrLf & llmResult
            ApDashboardLog($"✓ Binary analysis completed for: {fileName} ({llmResult.Length:N0} chars)", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error analyzing binary attachment: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: merge_pdfs
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the merge_pdfs tool to combine multiple PDF attachments.
    ''' </summary>
    Private Async Function ExecuteMergePdfsTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames As New List(Of String)()
            If toolCall.Arguments?.ContainsKey("attachment_names") Then
                Dim namesObj = toolCall.Arguments("attachment_names")
                If TypeOf namesObj Is JArray Then
                    For Each item In DirectCast(namesObj, JArray)
                        targetNames.Add(item.ToString())
                    Next
                End If
            End If

            Dim outputNameObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("output_filename") Then outputNameObj = toolCall.Arguments("output_filename")
            Dim outputName = outputNameObj?.ToString()

            If String.IsNullOrWhiteSpace(outputName) Then outputName = "merged.pdf"

            Dim toMerge As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                ' Preserve order as specified
                toMerge = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim found = _apCurrentAttachments?.FirstOrDefault(
                        Function(a) a.OriginalFileName.Equals(name, StringComparison.OrdinalIgnoreCase) AndAlso
                                    Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing)
                    If found IsNot Nothing Then toMerge.Add(found)
                Next
            Else
                toMerge = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toMerge Is Nothing OrElse toMerge.Count < 2 Then
                response.Success = False
                response.Response = "Need at least 2 PDF attachments to merge."
                Return response
            End If

            context.Log($"Merging {toMerge.Count} PDFs into {outputName}")

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            ' Use PDFsharp to merge
            Using outputDoc As New PdfSharp.Pdf.PdfDocument()
                For Each att In toMerge
                    Using inputDoc = PdfSharp.Pdf.IO.PdfReader.Open(att.TempFilePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                        For Each page In inputDoc.Pages
                            outputDoc.AddPage(page)
                        Next
                    End Using
                Next
                outputDoc.Save(outputPath)
            End Using

            ' Register as output attachment
            toMerge(0).OutputFiles.Add(outputPath)

            response.Success = True
            response.Response = $"Successfully merged {toMerge.Count} PDFs into {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error merging PDFs: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: read_attachment
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the read_attachment tool for a single attachment.
    ''' Supports Word, PDF, Excel (.xlsx/.xls), PowerPoint (.pptx), and text-based files.
    ''' </summary>
    Private Async Function ExecuteReadAttachmentTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileNameObj As Object = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("attachment_name") Then fileNameObj = toolCall.Arguments("attachment_name")
            Dim fileName = fileNameObj?.ToString()

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = _apCurrentAttachments?.FirstOrDefault(
                Function(a) a.OriginalFileName.Equals(fileName, StringComparison.OrdinalIgnoreCase))

            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If att.IsOverSizeLimit Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' exceeds the size limit and cannot be processed."
                Return response
            End If

            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' could not be read."
                Return response
            End If

            context.Log($"Reading attachment: {fileName}")

            Dim text As String = Nothing
            Dim label As String = Nothing

            ' Try Office extraction first (handles Word docs)
            Dim extracted As Boolean = False
            Try
                extracted = TryExtractOfficeText(att.TempFilePath, text, label)
            Catch
            End Try

            ' Try Excel / PowerPoint via dedicated extractors
            If Not extracted Then
                Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
                Try
                    Select Case ext
                        Case ".xlsx", ".xls"
                            text = ExtractExcelText(att.TempFilePath)
                            extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                        Case ".pptx"
                            text = ExtractPowerPointText(att.TempFilePath)
                            extracted = Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error")
                    End Select
                Catch
                End Try
            End If

            ' Try text-like files
            If Not extracted Then
                Try
                    extracted = TryExtractTextLike(att.TempFilePath, text, label)
                Catch
                End Try
            End If

            ' Try PDF
            If Not extracted AndAlso att.Extension = ".pdf" Then
                Try
                    text = Await SharedMethods.ReadPdfAsText(att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=False, AskUser:=False)
                    extracted = Not String.IsNullOrWhiteSpace(text)
                Catch
                End Try
            End If

            If extracted AndAlso Not String.IsNullOrWhiteSpace(text) Then
                ' Truncate very large files
                If text.Length > 50000 Then
                    text = text.Substring(0, 50000) & vbCrLf & "[... content truncated at 50,000 characters ...]"
                End If
                response.Success = True
                response.Response = text
            Else
                response.Success = False
                response.Response = $"Could not extract text from '{fileName}'. The file format may not be supported."
            End If

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading attachment: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: list_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Executes the list_attachments tool for the current email.
    ''' </summary>
    Private Function ExecuteListAttachmentsTool(
            toolCall As ToolCall,
            context As ToolExecutionContext) As ToolResponse

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow
        }

        Try
            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count = 0 Then
                response.Success = True
                response.Response = "No attachments in this email."
                Return response
            End If

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine($"Attachments ({_apCurrentAttachments.Count}):")
            For i As Integer = 0 To _apCurrentAttachments.Count - 1
                Dim att = _apCurrentAttachments(i)
                Dim sizeStr = If(att.SizeBytes > 0, $"{att.SizeBytes / 1024:F0} KB", "unknown size")
                Dim statusStr = If(att.IsOverSizeLimit, " [OVER SIZE LIMIT]",
                               If(att.TempFilePath IsNot Nothing, " [available for processing]", " [not available]"))
                sb.AppendLine($"  {i + 1}. {att.OriginalFileName} ({att.Extension}, {sizeStr}){statusStr}")
            Next

            response.Success = True
            response.Response = sb.ToString().TrimEnd()

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error listing attachments: {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Extension point: checks whether a tool call should be handled by AutoPilot
    ''' internal tools before delegating to external tools. Called from the existing
    ''' ExecuteToolCall method — requires a small integration point there.
    ''' </summary>
    Friend Function IsAutoPilotInternalTool(toolName As String) As Boolean
        Select Case toolName
            Case AP_Tool_ProcessWordDoc,
                 AP_Tool_CommentWordDoc,
                 AP_Tool_ExtractPdfText,
                 AP_Tool_MergePdfs,
                 AP_Tool_ReadAttachment,
                 AP_Tool_ListAttachments,
                 AP_Tool_DescribeBinary,
                 AP_Tool_CompareWordDocs
                Return True
            Case Else
                Return False
        End Select
    End Function

End Class