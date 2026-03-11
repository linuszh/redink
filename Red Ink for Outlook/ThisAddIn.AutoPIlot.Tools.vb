' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.vb
' Purpose:
'   Defines, registers, and executes AutoPilot internal tools used by Outlook
'   AutoPilot/Chat-Agent runs for attachment processing, document generation,
'   PDF/Office conversion, binary analysis, and structured fallback handling.
'
' Architecture:
'  - Registration:
'      * Exposes built-in tools as `ModelConfig` entries (`Tool=True`, `ToolOnly=True`)
'        so they participate in the same tool-calling pipeline as external tools.
'      * Tool metadata (`ToolDefinition`, `ToolInstructionsPrompt`) is generated inline
'        and consumed by `ExecuteToolCall` / `ExecuteToolingLoop`.
'  - Dispatch:
'      * `TryExecuteAutoPilotTool` routes a parsed tool call to strongly scoped
'        executors (`Execute*Tool` methods) and returns `ToolResponse` payloads.
'  - Session scope:
'      * Uses AutoPilot session state from `ThisAddIn.Autopilot.vb`:
'          - `_apCurrentAttachments`
'          - `_apCurrentTempDir`
'          - `_apCurrentMailInfo`
'      * Supports tool chaining via output registration (`OutputFiles`) and lookup via
'        `FindAttachment` (original attachments + prior tool outputs).
'  - Office/PDF processing:
'      * Word/PPT/Excel processing and generation via Interop/OpenXML helpers.
'      * PDF extraction/merge/split/watermark/comment flows via PdfPig/PdfSharp and
'        shared OCR-aware extraction helpers.
'  - Text extraction:
'      * Reads Office/text/PDF attachments with cache reuse
'        (`CachedText`, `CachedDocxHint`) to reduce repeated I/O and parsing.
'  - Logging and UX:
'      * Emits execution traces to tooling context and AutoPilot dashboard
'        (`context.Log`, `ApDashboardLog`) with concise success/failure summaries.
'
' Security & Safety:
'  - Path containment:
'      * Tool outputs are created in the per-mail temp directory and re-used only via
'        resolved attachment/output references.
'  - Size and format gates:
'      * Oversized attachments are excluded from processing.
'      * Binary analysis is restricted to explicit allow-listed extensions.
'  - Conversion safeguards:
'      * PDF→Word includes timeout/interop guardrails to avoid indefinite blocking.
'      * PDF font resolver is initialized centrally for deterministic PdfSharp behavior.
'  - Capability fallback:
'      * `report_inability` generates actionable alternatives (Red Ink + optional
'        InternetResearch path) instead of silent/tool-less failures.
'
' Built-in Tools (current):
'  - process_word_document
'  - comment_word_document
'  - comment_pdf_document
'  - compare_word_documents
'  - read_word_document_details
'  - read_attachment
'  - list_attachments
'  - search_in_attachments
'  - summarize_thread
'  - describe_binary_attachment
'  - extract_pdf_text
'  - merge_pdfs
'  - split_pdf
'  - add_pdf_watermark
'  - create_pdf_from_text
'  - word_to_pdf
'  - pdf_to_word
'  - extract_excel_data
'  - create_word_document
'  - create_excel_spreadsheet
'  - create_powerpoint
'  - create_code_file
'  - extract_data_from_attachments
'  - redact_pdf
'  - report_inability
'
' Notes:
'  - This file is a partial `ThisAddIn` implementation and depends on:
'      * AutoPilot state/config from `ThisAddIn.Autopilot.vb`
'      * Tooling contracts from `ThisAddIn.Tooling.vb`
'      * SharedLibrary helpers (`SharedMethods`, OCR/PDF/Office extractors)
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Xml
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

    Private Const AP_Tool_ProcessWordDoc As String = "process_word_document"
    Private Const AP_Tool_CommentWordDoc As String = "comment_word_document"
    Private Const AP_Tool_ExtractPdfText As String = "extract_pdf_text"
    Private Const AP_Tool_MergePdfs As String = "merge_pdfs"
    Private Const AP_Tool_ReadAttachment As String = "read_attachment"
    Private Const AP_Tool_ListAttachments As String = "list_attachments"
    Private Const AP_ToolPrefix As String = "autopilot_"
    Private Const AP_Tool_CompareWordDocs As String = "compare_word_documents"
    Private Const AP_Tool_ReadWordDocDetails As String = "read_word_document_details"
    Private Const AP_Tool_CreatePdfFromText As String = "create_pdf_from_text"
    Private Const AP_Tool_ExtractExcelData As String = "extract_excel_data"
    Private Const AP_Tool_SplitPdf As String = "split_pdf"
    Private Const AP_Tool_AddPdfWatermark As String = "add_pdf_watermark"
    Private Const AP_Tool_WordToPdf As String = "word_to_pdf"
    Private Const AP_Tool_SearchInAttachments As String = "search_in_attachments"
    Private Const AP_Tool_SummarizeThread As String = "summarize_thread"
    Private Const AP_Tool_PdfToWord As String = "pdf_to_word"
    Private Const AP_Tool_CreateWordDoc As String = "create_word_document"
    Private Const AP_Tool_CreateExcel As String = "create_excel_spreadsheet"
    Private Const AP_Tool_CreatePowerPoint As String = "create_powerpoint"
    Private Const AP_Tool_CreateCodeFile As String = "create_code_file"
    Private Const AP_Tool_CommentPdf As String = "comment_pdf_document"
    Private Const AP_Tool_ExtractDataFromAttachments As String = "extract_data_from_attachments"
    Private Const AP_Tool_RedactPdf As String = "redact_pdf"
    Private Const AP_Tool_OverlayPdf As String = "overlay_pdf"
    Private Const AP_Tool_ReportInability As String = "report_inability"


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL REGISTRATION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Builds and returns the full set of AutoPilot internal tool definitions.
    ''' </summary>
    ''' <returns>
    ''' A list of <see cref="ModelConfig"/> items representing built-in tools
    ''' registered for the current AutoPilot run.
    ''' </returns>
    Friend Function GetAutoPilotInternalTools() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        ' ── process_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ProcessWordDoc,
            .ModelDescription = "Process Word/PowerPoint/Excel Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ProcessWordDoc & ": Processes one or more Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) attachments by applying a prompt/instruction. " &
                "Use this for translation, correction, proofreading, anonymization, data updates, formula changes, or any text/data transformation. " &
                "For Word documents, returns both a clean version and a compare document showing changes. " &
                "For PowerPoint and Excel files, returns the processed version (no compare document). " &
                "For Excel files, you can optionally restrict processing to specific sheet names using the sheet_names parameter. " &
                "CRITICAL — ONE OPERATION PER CALL: This tool applies exactly ONE instruction per call. " &
                "If the user requests multiple distinct operations (e.g., 'correct and translate', 'anonymize and summarize', 'fix grammar then make more concise'), " &
                "you MUST split them into separate sequential calls. First call: apply the first operation to the original file. " &
                "Wait for the result. Second call: apply the second operation to the output file from the first call (the '_processed' file). " &
                "Example for 'correct and translate to German': " &
                "(1) Call process_word_document with instruction='Correct spelling, grammar and style' on 'Contract.docx'. Result: 'Contract_processed.docx'. " &
                "(2) Call process_word_document with instruction='Translate to German' on attachment_names=['Contract_processed.docx']. Result: 'Contract_processed_processed.docx'. " &
                "NEVER combine two distinct operations into a single instruction string. " &
                "However, a single coherent task counts as one operation (e.g., 'Translate to German' is one operation even though it involves reading and rewriting). " &
                "Output files are named '<original>_processed.<ext>' and can be referenced in subsequent tool calls by that name.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ProcessWordDoc & """," &
                """description"":""Applies exactly ONE processing instruction to Word (.docx), PowerPoint (.pptx), or Excel (.xlsx) attachments. " &
                "Supports translation, correction, anonymization, data updates, formula modifications, and freestyle operations. " &
                "For Word documents, produces clean output plus a compare document with tracked changes. " &
                "For PowerPoint and Excel, produces the processed file only. " &
                "IMPORTANT: Apply only ONE operation per call. For multi-step requests (e.g. 'correct and translate'), " &
                "make separate sequential calls — first correct, then translate the corrected output file. " &
                "Output files are named '<original>_processed.<ext>' and can be used as input for the next call via attachment_names.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""A single, specific instruction to apply to the document. Must be ONE operation only — " &
                "e.g. 'Translate to German' or 'Correct spelling and grammar' or 'Anonymize all personal names'. " &
                "Do NOT combine multiple operations like 'Correct and translate'. Split those into separate calls.""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the attachments to process. " &
                "Can include output files from previous tool calls (e.g. 'Contract_processed.docx'). " &
                "If empty or omitted, processes all .docx, .pptx, and .xlsx attachments.""}," &
                """sheet_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional: for Excel files only, restrict processing to these sheet names. If omitted, all sheets are processed.""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── extract_pdf_text ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractPdfText,
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
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_MergePdfs,
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
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReadAttachment,
            .ModelDescription = "Read Attachment Content (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReadAttachment & ": Reads and returns the text content of one or more supported attachments " &
                "(DOCX, PDF, TXT, CSV, HTML, XML, JSON, XLSX, XLS, PPTX). " &
                "Embedded mail files (.msg, .eml) are automatically unpacked — their body text and nested attachments " &
                "are extracted recursively and appear as separate attachments that you can reference by name. " &
                "Use attachment_name for a single file or attachment_names for batch reading.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReadAttachment & """," &
                """description"":""Reads and returns the text content of one or more attachment files. " &
                "Supports Word documents (.docx), PDFs (.pdf), Excel spreadsheets (.xlsx, .xls), " &
                "PowerPoint presentations (.pptx), and text-based files (.txt, .csv, .html, .xml, .json, .md, .log). " &
                "Embedded mail files (.msg, .eml) are automatically unpacked at intake — their text content " &
                "and nested attachments appear as separate files in the attachment list. " &
                "For Word documents, also reports if comments or tracked changes are present.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of a single attachment to read""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of multiple attachments to read in batch. Use this instead of attachment_name when reading several files.""}" &
                "},""required"":[]}}"
        })

        ' ── list_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ListAttachments,
            .ModelDescription = "List Attachments (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ListAttachments & ": Lists all attachments of the current email with name, type, and size.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ListAttachments & """," &
                """description"":""Lists all email attachments with their filename, type, size, and processing status""," &
                """parameters"":{""type"":""object"",""properties"":{},""required"":[]}}"
        })

        ' ── describe_binary_attachment ──
        Dim apiCallObj As String = If(_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi,
                                      INI_APICall_Object_2, INI_APICall_Object)
        If Not String.IsNullOrWhiteSpace(apiCallObj) Then
            tools.Add(New ModelConfig() With {
                .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_DescribeBinary,
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
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CommentWordDoc,
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
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CompareWordDocs,
            .ModelDescription = "Compare two Word documents (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CompareWordDocs & ": Compares exactly two Word document (.doc/.docx) attachments using Word's " &
                "built-in comparison engine (track changes). Use 'original_filename' for the BASE/earlier/reference " &
                "version and 'revised_filename' for the MODIFIED/newer/changed version. Returns a textual summary of " &
                "revisions found and produces a comparison document with tracked changes attached to the reply. " &
                "This tool requires exactly two attachments to be specified — it cannot compare more than two at once. " &
                "If the sender provides more than two Word documents, ask which two should be compared, or run " &
                "multiple comparisons. " &
                "IMPORTANT: This tool can also accept output files produced by other tools (e.g. a '_processed.docx' " &
                "from process_word_document). Note that process_word_document already produces its own compare " &
                "document automatically — only use compare_word_documents separately when comparing two independently " &
                "provided attachments or when the user explicitly asks for a comparison between specific files.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CompareWordDocs & """," &
                """description"":""Compares exactly two Word documents (.doc/.docx) using Word's built-in comparison (track changes). " &
                "The 'original_filename' is the BASE document (the earlier or reference version). " &
                "The 'revised_filename' is the MODIFIED document (the newer or changed version). " &
                "Returns a textual summary of the differences found and produces a comparison document with tracked changes as a result attachment. " &
                "IMPORTANT: 'original_filename' = the source/baseline; 'revised_filename' = the version that was changed or updated. " &
                "This tool compares exactly two documents per call. " &
                "Can also reference output files from earlier tools (e.g. '_processed.docx' from process_word_document).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """original_filename"":{""type"":""string"",""description"":""Exact filename of the original/baseline/source Word attachment (the earlier version)""}," &
                """revised_filename"":{""type"":""string"",""description"":""Exact filename of the revised/modified/updated Word attachment (the newer version)""}" &
                "},""required"":[""original_filename"",""revised_filename""]}}"
        })

        ' ── read_word_document_details ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReadWordDocDetails,
            .ModelDescription = "Read Word Document Details (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReadWordDocDetails & ": Deep-reads a Word document (.docx) including body text with inline tracked changes, " &
                "comment bubbles (with author, date, and anchored text), headers, footers, footnotes, and endnotes. " &
                "This is a heavier tool — only use it when the user explicitly asks about comments, tracked changes, " &
                "revisions, review history, headers/footers, or footnotes/endnotes. For general content questions, use read_attachment instead. " &
                "Tracked changes are shown inline using «INS|author|date»...«/INS» and «DEL|author|date»...«/DEL» markers. " &
                "Use tracked_changes_author and tracked_changes_since to filter changes by a specific author or date.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReadWordDocDetails & """," &
                """description"":""Deep-reads a Word document (.docx) with comments, tracked changes, headers/footers, and footnotes/endnotes. " &
                "Only use when the user asks about comments, revisions, changes, review history, headers, footers, footnotes, or endnotes. " &
                "For general content, use read_attachment instead.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the .docx attachment to read""}," &
                """include_comments"":{""type"":""boolean"",""description"":""Include comment bubbles with author, date, and anchored text (default: true)""}," &
                """include_headers_footers"":{""type"":""boolean"",""description"":""Include headers and footers (default: false)""}," &
                """include_footnotes_endnotes"":{""type"":""boolean"",""description"":""Include footnotes and endnotes (default: false)""}," &
                """include_tracked_changes"":{""type"":""boolean"",""description"":""Include tracked changes as inline markers in the body text (default: true)""}," &
                """tracked_changes_author"":{""type"":""string"",""description"":""Optional: only show tracked changes by this author""}," &
                """tracked_changes_since"":{""type"":""string"",""description"":""Optional: only show tracked changes on or after this date (ISO 8601, e.g. '2026-01-15')""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── create_pdf_from_text ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreatePdfFromText,
            .ModelDescription = "Create PDF from Text (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreatePdfFromText & ": Creates a PDF document from provided text content. " &
                "Use this when the user wants a new PDF created with specific content.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreatePdfFromText & """," &
                """description"":""Creates a PDF file from provided text content.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """content"":{""type"":""string"",""description"":""The text content for the PDF""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: output.pdf)""}," &
                """title"":{""type"":""string"",""description"":""Optional title displayed at the top of the PDF""}" &
                "},""required"":[""content""]}}"
        })

        ' ── extract_excel_data ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractExcelData,
            .ModelDescription = "Extract Excel Data (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ExtractExcelData & ": Reads data from an Excel attachment (.xlsx/.xls) with control over which sheet to read. " &
                "Returns data in CSV-like tabular format.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ExtractExcelData & """," &
                """description"":""Reads data from an Excel spreadsheet attachment with sheet selection. Returns tabular data.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Excel attachment""}," &
                """sheet_name"":{""type"":""string"",""description"":""Optional: name of the specific sheet to read. If omitted, reads all sheets.""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── split_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SplitPdf,
            .ModelDescription = "Split PDF (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SplitPdf & ": Extracts a range of pages from a PDF attachment into a new PDF.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SplitPdf & """," &
                """description"":""Extracts a page range from a PDF into a new PDF file.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment""}," &
                """start_page"":{""type"":""integer"",""description"":""First page to extract (1-based)""}," &
                """end_page"":{""type"":""integer"",""description"":""Last page to extract (1-based, inclusive)""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: split.pdf)""}" &
                "},""required"":[""attachment_name"",""start_page"",""end_page""]}}"
        })

        ' ── add_pdf_watermark ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_AddPdfWatermark,
            .ModelDescription = "Add PDF Watermark (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_AddPdfWatermark & ": Adds a diagonal text watermark to every page of a PDF attachment.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_AddPdfWatermark & """," &
                """description"":""Adds a diagonal text watermark to every page of a PDF.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment""}," &
                """watermark_text"":{""type"":""string"",""description"":""Text for the watermark (e.g., 'DRAFT', 'CONFIDENTIAL')""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: watermarked.pdf)""}" &
                "},""required"":[""attachment_name"",""watermark_text""]}}"
        })

        ' ── word_to_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_WordToPdf,
            .ModelDescription = "Convert Word to PDF (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_WordToPdf & ": Converts a Word document (.doc/.docx) attachment to PDF format using Word.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_WordToPdf & """," &
                """description"":""Converts a Word document (.doc/.docx) attachment to PDF format.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the Word document attachment to convert""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── search_in_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SearchInAttachments,
            .ModelDescription = "Search in Attachments (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SearchInAttachments & ": Searches for a keyword or phrase across all readable attachments. " &
                "Returns matching lines with surrounding context.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SearchInAttachments & """," &
                """description"":""Searches for a keyword or phrase across all readable attachments and returns matching excerpts with context.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """search_term"":{""type"":""string"",""description"":""The keyword or phrase to search for (case-insensitive)""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional: limit search to these attachments. If omitted, searches all.""}" &
                "},""required"":[""search_term""]}}"
        })

        ' ── summarize_thread ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_SummarizeThread,
            .ModelDescription = "Summarize Email Thread (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_SummarizeThread & ": Extracts and structures the email conversation thread from the current mail, " &
                "excluding messages sent to/from the monitored AutoPilot mailbox. Returns each message with sender, date, and body.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_SummarizeThread & """," &
                """description"":""Extracts the full email conversation thread, excluding AutoPilot's own replies. Returns structured messages with sender, date, and content.""," &
                """parameters"":{""type"":""object"",""properties"":{},""required"":[]}}"
        })

        ' ── pdf_to_word ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_PdfToWord,
            .ModelDescription = "Convert PDF to Word (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_PdfToWord & ": Converts a PDF attachment to a Word document (.docx) using Word's built-in PDF import. " &
                "The resulting .docx can then be used with compare_word_documents or other Word tools. " &
                "This is the PREFERRED method for PDF-to-Word conversion — use it FIRST. It works well for most PDFs " &
                "that contain real (selectable/searchable) text and preserves layout, tables, and formatting. " &
                "If the conversion result indicates the PDF is scanned/image-only (no extractable text), THEN " &
                "fall back to extract_pdf_text (which supports OCR) to obtain the text, and use create_word_document " &
                "to produce a .docx from that text.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_PdfToWord & """," &
                """description"":""Converts a PDF attachment to a Word document (.docx) using Word's built-in PDF reflow. " &
                "Use this as the PRIMARY method for PDF-to-Word conversion. Works well for text-based PDFs with layout preservation. " &
                "If the result indicates the PDF is scanned/image-only, fall back to extract_pdf_text (OCR) + create_word_document.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to convert""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output .docx (default: derived from PDF name)""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── create_word_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateWordDoc,
            .ModelDescription = "Create Word Document from Markdown (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateWordDoc & ": Creates a new formatted Word document (.docx) from Markdown content. " &
                "Use this when the user asks you to create, generate, or produce a new Word document from any content " &
                "(e.g., from a PDF extract, from research results, from your own generated text, from a summary, etc.). " &
                "Provide the content as Markdown and it will be converted to a properly formatted .docx file with " &
                "headings, bold, italic, lists, etc. The resulting file will be attached to the reply.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateWordDoc & """," &
                """description"":""Creates a new formatted Word document (.docx) from Markdown content. " &
                "Use when the user asks to create, generate, or produce a Word document from any content. " &
                "The Markdown is converted to a properly formatted .docx with headings, bold, italic, lists, etc.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """markdown_content"":{""type"":""string"",""description"":""The full document content in Markdown format. " &
                "Use headings (#, ##, ###), bold (**text**), italic (*text*), lists (- or 1.), etc.""}," &
                """file_name"":{""type"":""string"",""description"":""The desired filename for the output Word document " &
                "(without .docx extension). Defaults to 'Document' if not specified.""}" &
                "},""required"":[""markdown_content""]}}"
        })

        ' ── create_excel_spreadsheet ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateExcel,
            .ModelDescription = "Create Excel Spreadsheet (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateExcel & ": Creates a professionally formatted Excel spreadsheet (.xlsx/.xlsm). " &
                "Use for any spreadsheet, table, budget, tracker, dashboard, or tabular data request. " &
                "ALWAYS call this tool immediately — do NOT describe what you plan to create; just create it. " &
                "MANDATORY: include column_widths for every column, freeze_pane='A2', auto_filter on headers, " &
                "bold headers with bg_color/font_color/borders, alternating row colors, wrap_text on long text, " &
                "number_format on numeric/date columns, and data_validations (dropdowns) for any column with finite valid values.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateExcel & """," &
                """description"":""Creates a new Excel spreadsheet (.xlsx or .xlsm) with MANDATORY professional formatting. " &
                "EVERY spreadsheet MUST include: column_widths sized for content, freeze_pane='A2', auto_filter on headers, " &
                "header row with bold + bg_color '#4472C4' + font_color '#FFFFFF' + border 'all-thin' + h_align 'center', " &
                "data cells with border 'all-thin' + alternating bg_color '#D9E2F3', wrap_text on long text columns, " &
                "number_format on all numeric/date/percentage columns, and data_validations with type 'list' on any column " &
                "with a finite set of valid values (Status, Priority, Yes/No, categories, etc.). " &
                "Also supports formulas, multiple sheets, conditional formatting, charts, VBA macros, and more. " &
                "STYLING RULES (apply to EVERY spreadsheet unless user explicitly asks for plain output): " &
                "1. HEADER ROW: bg_color '#4472C4', font_color '#FFFFFF', bold, font_size 12, h_align 'center', border 'all-thin'. " &
                "2. DATA ROWS: border 'all-thin' on ALL cells, alternating bg_color '#D9E2F3' on even rows. " &
                "3. COLUMN WIDTHS: MUST set for every column. Short labels 12-15, names/descriptions 25-35, numbers/dates 12-18. " &
                "4. ROW HEIGHTS: header row 28. " &
                "5. NUMBER FORMATS: '#,##0.00' for currency, '0%' for percent, 'dd/mm/yyyy' for dates, '#,##0' for integers. " &
                "6. ALIGNMENT: left for text, center for short labels/status, right for numbers. " &
                "7. WRAP TEXT: true for descriptions, notes, addresses, any long content. " &
                "8. FREEZE PANE: ALWAYS 'A2'. " &
                "9. AUTO FILTER: ALWAYS on header range (e.g. 'A1:F1'). " &
                "10. DROPDOWNS: ALWAYS add data_validation type 'list' for columns with finite values (Status, Priority, Yes/No, Rating, Category). " &
                "11. CONDITIONAL FORMATTING: red bg for negative/overdue/failed, green for completed/positive, yellow for pending. " &
                "12. TOTALS ROW: SUM/AVERAGE formulas with bold + border 'bottom-medium'. " &
                "13. TITLE ROW: For dashboards, merge + font_size 16 + bold + distinct bg_color. " &
                "Use English formula syntax with comma separators.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """cells"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """cell"":{""type"":""string"",""description"":""Cell address in A1 notation""}," &
                """value"":{""description"":""Cell value (string or number)""}," &
                """formula"":{""type"":""string"",""description"":""Excel formula starting with =""}," &
                """bold"":{""type"":""boolean""}," &
                """italic"":{""type"":""boolean""}," &
                """underline"":{""type"":""boolean""}," &
                """strikethrough"":{""type"":""boolean""}," &
                """font_name"":{""type"":""string""}," &
                """font_size"":{""type"":""number""}," &
                """font_color"":{""type"":""string"",""description"":""Hex RGB e.g. #FF0000. Use #FFFFFF for headers""}," &
                """bg_color"":{""type"":""string"",""description"":""Hex RGB. Use #4472C4 for headers, #D9E2F3 for alternating rows""}," &
                """number_format"":{""type"":""string"",""description"":""REQUIRED for numbers/dates. #,##0.00 currency, 0% percent, dd/mm/yyyy dates, #,##0 integers""}," &
                """h_align"":{""type"":""string"",""enum"":[""left"",""center"",""right""],""description"":""REQUIRED: left=text, center=headers/labels, right=numbers""}," &
                """v_align"":{""type"":""string"",""enum"":[""top"",""center"",""bottom""]}," &
                """wrap_text"":{""type"":""boolean"",""description"":""REQUIRED true for long text cells""}," &
                """border"":{""type"":""string"",""description"":""REQUIRED: 'all-thin' for all cells. Also: medium, thick, all-medium, bottom-thin, bottom-medium""}," &
                """border_color"":{""type"":""string"",""description"":""Border color hex RGB""}" &
                "}},""description"":""Cells for default/first sheet""}," &
                """sheets"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """name"":{""type"":""string""},""cells"":{""type"":""array"",""items"":{""type"":""object""}}" &
                "}},""description"":""Multiple sheets. Each: name + cells array.""}," &
                """file_name"":{""type"":""string"",""description"":""Filename without extension""}," &
                """sheet_name"":{""type"":""string"",""description"":""Tab name for single-sheet mode""}," &
                """column_widths"":{""type"":""object"",""description"":""REQUIRED: {col_letter: width} for EVERY column. 12-15 short, 25-35 descriptions, 12-18 numbers""}," &
                """row_heights"":{""type"":""object"",""description"":""Row heights. Set header row to 28""}," &
                """merge_ranges"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Ranges to merge""}," &
                """freeze_pane"":{""type"":""string"",""description"":""REQUIRED: Always 'A2'""}," &
                """auto_filter"":{""type"":""string"",""description"":""REQUIRED: Header range e.g. 'A1:F1'""}," &
                """data_validations"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """range"":{""type"":""string""},""type"":{""type"":""string""},""formula1"":{""type"":""string""}," &
                """formula2"":{""type"":""string""},""operator"":{""type"":""string""}," &
                """show_dropdown"":{""type"":""boolean""},""input_title"":{""type"":""string""}," &
                """input_message"":{""type"":""string""},""error_title"":{""type"":""string""}," &
                """error_message"":{""type"":""string""}" &
                "}},""description"":""REQUIRED for finite-value columns: type 'list', formula1='Val1,Val2,Val3'""}," &
                """conditional_formats"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """range"":{""type"":""string""},""type"":{""type"":""string""},""operator"":{""type"":""string""}," &
                """formula1"":{""type"":""string""},""formula2"":{""type"":""string""}," &
                """format_font_color"":{""type"":""string""},""format_bg_color"":{""type"":""string""}," &
                """format_bold"":{""type"":""boolean""}" &
                "}},""description"":""Conditional formatting rules""}," &
                """charts"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """type"":{""type"":""string"",""enum"":[""column"",""bar"",""line"",""pie"",""area"",""scatter"",""doughnut""]}," &
                """data_range"":{""type"":""string""},""title"":{""type"":""string""}," &
                """position"":{""type"":""string""},""width"":{""type"":""number""},""height"":{""type"":""number""}," &
                """sheet_name"":{""type"":""string""}" &
                "}},""description"":""Charts to create""}," &
                """named_ranges"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """name"":{""type"":""string""},""range"":{""type"":""string""}" &
                "}},""description"":""Named ranges""}," &
                """vba_modules"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """name"":{""type"":""string""},""code"":{""type"":""string""},""type"":{""type"":""string""}" &
                "}},""description"":""VBA modules (saves as .xlsm)""}," &
                """print_setup"":{""type"":""object"",""properties"":{" &
                """orientation"":{""type"":""string"",""enum"":[""portrait"",""landscape""]}," &
                """fit_to_pages_wide"":{""type"":""integer""},""fit_to_pages_tall"":{""type"":""integer""}," &
                """header_text"":{""type"":""string""},""footer_text"":{""type"":""string""}" &
                "},""description"":""Print setup options""}" &
                "},""required"":[]}}"
        })

        ' ── create_powerpoint ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreatePowerPoint,
            .ModelDescription = "Create PowerPoint Presentation (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreatePowerPoint & ": Creates a new PowerPoint presentation (.pptx) with slides containing titles, body text, and speaker notes. " &
                "Use this when the user asks you to create, generate, or produce a presentation, slide deck, or pitch deck. " &
                "Provide slide data as a JSON array of slide objects. Each slide object has: " &
                "'title' (string, the slide title), 'body' (string, the main content — use newlines for bullet points), " &
                "and optionally 'notes' (string, speaker notes for that slide). " &
                "The first slide is typically used as a title slide with a short subtitle in 'body'.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreatePowerPoint & """," &
                """description"":""Creates a new PowerPoint presentation (.pptx) with slides. Each slide has a title, body text (use newlines for bullets), and optional speaker notes. " &
                "Use when the user asks to create a presentation, slide deck, or pitch deck.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """slides"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """title"":{""type"":""string"",""description"":""Slide title text""}," &
                """body"":{""type"":""string"",""description"":""Slide body content. Use newline characters for separate bullet points.""}," &
                """notes"":{""type"":""string"",""description"":""Optional speaker notes for this slide""}" &
                "}},""description"":""Array of slide objects defining the presentation""}," &
                """file_name"":{""type"":""string"",""description"":""Desired filename without extension (default: 'Presentation')""}," &
                """title"":{""type"":""string"",""description"":""Presentation title metadata (default: derived from first slide title)""}" &
                "},""required"":[""slides""]}}"
        })

        ' ── create_code_file ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CreateCodeFile,
            .ModelDescription = "Create Code/Script File (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CreateCodeFile & ": Creates a new code, script, or data file with the specified content. " &
                "Use this when the user asks you to create, generate, write, or produce any code file, script, " &
                "configuration file, or structured data file. Examples: HTML pages, Python scripts, JavaScript files, " &
                "JSON/YAML/XML data files, CSS stylesheets, SQL scripts, shell scripts (.sh/.bat/.ps1), " &
                "Markdown documents, CSV files, INI/TOML/ENV config files, Dockerfiles, etc. " &
                "You MUST determine the appropriate file extension based on the content and language. " &
                "You MUST provide the complete, functional, ready-to-execute file content — do NOT use placeholders " &
                "or incomplete code. The resulting file will be attached to the reply email so the user can save and run it.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CreateCodeFile & """," &
                """description"":""Creates a new code, script, or data file with the specified content and attaches it to the reply. " &
                "Supports any text-based file format: HTML, Python, JavaScript, TypeScript, JSON, YAML, XML, CSS, SQL, " &
                "shell scripts, batch files, PowerShell, Markdown, CSV, INI, TOML, Dockerfiles, and more. " &
                "Determine the correct filename and extension based on the content and user request.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """file_name"":{""type"":""string"",""description"":""Full filename including extension (e.g. 'index.html', 'analysis.py', 'config.json', 'setup.sh'). " &
                "Choose a descriptive name and the correct extension for the language/format.""}," &
                """content"":{""type"":""string"",""description"":""The complete file content. Must be functional and ready to use — no placeholders or TODOs.""}," &
                """description"":{""type"":""string"",""description"":""Optional brief description of what the file does, shown to the user in the response.""}" &
                "},""required"":[""file_name"",""content""]}}"
        })

        ' ── comment_pdf_document ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_CommentPdf,
            .ModelDescription = "Comment PDF Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_CommentPdf & ": Adds review comments as highlight annotations with popups to a PDF attachment. " &
                "Use this ONLY when the user explicitly asks to ADD, INSERT, or PLACE comments, annotations, " &
                "review notes, or feedback INSIDE a PDF file — i.e. the user wants the PDF itself modified with " &
                "embedded annotation bubbles. " &
                "Do NOT use this tool when the user asks to READ, EXTRACT, SUMMARIZE, or UNDERSTAND existing " &
                "comments or content from a PDF — use extract_pdf_text or read_attachment for that instead. " &
                "Do NOT use this tool when the user wants a textual summary or analysis of a PDF — only when " &
                "annotations should appear as highlight + popup comment pairs within the PDF itself. " &
                "Supports an optional author parameter: if the user asks for comments under a specific name " &
                "(e.g. the sender's name), pass it as author. If not specified, comments are authored as 'Inky'. " &
                "Comments that cannot be matched to specific text in the PDF are placed as sticky notes " &
                "at the top-right corner of the first page.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_CommentPdf & """," &
                """description"":""Adds review comments as highlight annotations with popup bubbles directly inside a PDF file. " &
                "Use ONLY when the user wants to ADD or INSERT comments/annotations/review feedback INTO the PDF. " &
                "Do NOT use when the user wants to READ or EXTRACT existing content or comments from a PDF. " &
                "Matched text is highlighted in yellow with a popup comment; unmatched comments become sticky notes.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""The review instruction (e.g., 'Review for legal risks', 'Check for inconsistencies', 'Suggest improvements')""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of the PDF attachments to annotate. If empty or omitted, annotates all PDF attachments.""}," &
                """author"":{""type"":""string"",""description"":""Optional author name for the annotations. Use this when the user requests a specific name. If omitted, defaults to 'Inky'.""}" &
                "},""required"":[""instruction""]}}"
        })


        ' ── extract_data_from_attachments ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ExtractDataFromAttachments,
            .ModelDescription = "Extract structured data from attachments into a table (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ExtractDataFromAttachments & ": Extracts structured/tabular data from one or more attachments " &
                "(PDF, Word, Excel, text files, or files from a .zip archive) using AI-driven fact extraction. " &
                "You MUST provide an 'instruction' describing WHAT to extract (e.g. 'Extract invoice number, date, vendor name, and total amount'). " &
                "You SHOULD provide a 'schema' defining the output columns using the format 'ColumnName:type;ColumnName:type' " &
                "where type is one of: text, date, datetime, number, other. Example: 'Invoice Number:text;Date:date;Vendor:text;Amount:number'. " &
                "If no schema is provided, the AI will infer one automatically. " &
                "The tool processes each file individually with the AI, then merges and returns the combined result as a JSON table. " &
                "After receiving the result, YOU decide the best output format based on the user's request: " &
                "- Use create_excel_spreadsheet to produce a formatted .xlsx file " &
                "- Use create_word_document to produce a formatted .docx report " &
                "- Include the data directly in your reply as a formatted text table " &
                "- Or any other appropriate presentation. " &
                "This tool ONLY extracts and returns the structured data — it does NOT create any files.",
                .ToolDefinition =
                "{""name"":""" & AP_Tool_ExtractDataFromAttachments & """," &
                """description"":""Extracts structured/tabular data from one or more attachments using AI-driven fact extraction. " &
                "Returns a JSON object with 'schema' (column definitions) and 'rows' (extracted data). " &
                "Supports PDF, Word, Excel, text files, and files unpacked from .zip archives. " &
                "You MUST then decide how to present the result (create_excel_spreadsheet, create_word_document, or inline table).""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """instruction"":{""type"":""string"",""description"":""Natural-language instruction describing what data to extract. " &
                "Be specific about the fields/facts to capture. Example: 'Extract party names, contract date, governing law, and termination clauses from each document.'""}," &
                """schema"":{""type"":""string"",""description"":""Optional but recommended: column definitions in 'Name:type;Name:type' format. " &
                "Types: text, date, datetime, number, other. Append * to mark the sort column. " &
                "Example: 'Invoice No:text;Date:date*;Vendor:text;Amount:number;Notes:text'. " &
                "If omitted, the AI infers the schema automatically.""}," &
                """attachment_names"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Filenames of attachments to extract data from. " &
                "If empty or omitted, processes all readable attachments (PDF, DOCX, XLSX, TXT, CSV, etc.).""}," &
                """output_language"":{""type"":""string"",""description"":""Language for extracted column names and textual values (e.g. 'English', 'German', 'French'). " &
                "Use the language the user expects the output in. If omitted, uses the language of the user's email.""}," &
                """sort_column"":{""type"":""integer"",""description"":""Optional: 1-based column index to sort by. Use 0 or omit for no sorting.""}," &
                """sort_direction"":{""type"":""string"",""enum"":[""ASC"",""DESC""],""description"":""Sort direction (default: ASC)""}," &
                """date_columns"":{""type"":""string"",""description"":""Optional: comma-separated 1-based column indices that contain dates, for normalization. Example: '2,5'""}" &
                "},""required"":[""instruction""]}}"
        })

        ' ── redact_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_RedactPdf,
            .ModelDescription = "Redact PDF Document (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_RedactPdf & ": Redacts a PDF document by identifying text that matches the given instruction " &
                "and placing redaction boxes over it. Uses AI to analyze the PDF text and determine what should be redacted. " &
                "Operates in three modes controlled by the 'mode' parameter: " &
                "(1) 'prepare' (default): Creates removable red annotation boxes over identified text. " &
                "The user can review and adjust these in a PDF viewer before finalizing. " &
                "(2) 'finalize': Takes a previously prepared PDF (with redaction annotation boxes) and burns " &
                "them into permanent black rectangles by rasterizing each page. No AI analysis is performed. " &
                "(3) 'prepare_and_finalize': Performs both steps in one call — identifies text, places boxes, " &
                "and immediately burns them in as permanent black redactions. " &
                "IMPORTANT: When mode is 'prepare' or 'prepare_and_finalize', an 'instruction' is REQUIRED " &
                "describing what to redact (e.g. 'Redact all personal names and addresses', " &
                "'Redact financial information', 'Redact everything except party names and dates'). " &
                "When mode is 'finalize', no instruction is needed — it just burns in existing annotations. " &
                "The 'include_reason_codes' parameter adds brief labels (e.g. 'name', 'address') to each " &
                "redaction box, which are visible as white text inside the black box after finalization.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_RedactPdf & """," &
                """description"":""Redacts a PDF by identifying text with AI and placing redaction boxes, " &
                "or finalizes existing redaction boxes into permanent black rectangles. " &
                "Modes: 'prepare' (removable red boxes), 'finalize' (burn in existing boxes), " &
                "'prepare_and_finalize' (identify + burn in one step). " &
                "Requires 'instruction' for prepare modes (e.g. 'Redact all personal data').""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to redact""}," &
                """instruction"":{""type"":""string"",""description"":""What to redact — required for 'prepare' and 'prepare_and_finalize' modes. " &
                "Examples: 'Redact all personal names, addresses, and phone numbers', " &
                "'Redact financial data including account numbers and amounts', " &
                "'Redact everything except the contract parties and effective dates'""}," &
                """mode"":{""type"":""string"",""enum"":[""prepare"",""finalize"",""prepare_and_finalize""]," &
                """description"":""Operation mode. 'prepare' = AI-driven removable boxes (default). " &
                "'finalize' = burn in existing annotation boxes. " &
                "'prepare_and_finalize' = AI-driven + immediate burn-in.""}," &
                """include_reason_codes"":{""type"":""boolean"",""description"":""Include brief reason labels (e.g. 'name', 'address') in each redaction box. Default: false.""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: derived from input with '_redacted' or '_final' suffix)""}" &
                "},""required"":[""attachment_name""]}}"
        })

        ' ── overlay_pdf ──
        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_OverlayPdf,
            .ModelDescription = "Overlay text and images on PDF pages (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_OverlayPdf & ": Places text labels and/or images at precise positions on PDF pages. " &
                "Use this when the user wants to add a logo, stamp, header, footer, badge, label, signature image, " &
                "or any positioned content onto a PDF. Supports per-element page targeting (single page, page range, or all pages), " &
                "font family/size/style/color, rotation, opacity, and image scaling. " &
                "Coordinates use PDF points (1 pt = 1/72 inch). A4 page = 595 × 842 pt. Letter = 612 × 792 pt. " &
                "Origin (0,0) is the TOP-LEFT corner of the page. " &
                "For images, reference an existing attachment by name via 'image_attachment_name'. " &
                "Text elements and image elements can be freely mixed in the same call. " &
                "The tool draws elements in array order (later elements overlay earlier ones).",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_OverlayPdf & """," &
                """description"":""Places text labels and/or images at precise positions on PDF pages. " &
                "Coordinates are in PDF points (1/72 inch). Origin (0,0) = top-left. A4 = 595×842 pt, Letter = 612×792 pt. " &
                "Elements are drawn in array order. Use for logos, stamps, headers, footers, labels, signatures, badges, etc.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """attachment_name"":{""type"":""string"",""description"":""Filename of the PDF attachment to overlay onto""}," &
                """elements"":{""type"":""array"",""items"":{""type"":""object"",""properties"":{" &
                """type"":{""type"":""string"",""enum"":[""text"",""image""],""description"":""Element type: 'text' for a text label, 'image' for an image file""}," &
                """pages"":{""type"":""string"",""description"":""Target pages: 'all' for every page, '1' for page 1 only, '1,3,5' for specific pages, '2-5' for a range. Default: 'all'""}," &
                """x"":{""type"":""number"",""description"":""X position in points from the left edge of the page""}," &
                """y"":{""type"":""number"",""description"":""Y position in points from the top edge of the page""}," &
                """text"":{""type"":""string"",""description"":""(text only) The text string to render. Supports \\n for line breaks.""}," &
                """font_family"":{""type"":""string"",""description"":""(text only) Font family name, e.g. 'Arial', 'Times New Roman', 'Calibri'. Default: 'Arial'""}," &
                """font_size"":{""type"":""number"",""description"":""(text only) Font size in points. Default: 12""}," &
                """bold"":{""type"":""boolean"",""description"":""(text only) Bold text. Default: false""}," &
                """italic"":{""type"":""boolean"",""description"":""(text only) Italic text. Default: false""}," &
                """font_color"":{""type"":""string"",""description"":""(text only) Hex RGB color, e.g. '#FF0000' for red, '#000000' for black. Default: '#000000'""}," &
                """h_align"":{""type"":""string"",""enum"":[""left"",""center"",""right""],""description"":""(text only) Horizontal alignment relative to x position. 'left' = x is left edge, 'center' = x is center point, 'right' = x is right edge. Default: 'left'""}," &
                """max_width"":{""type"":""number"",""description"":""(text only) Maximum width in points for text bounding box. Text is clipped or wrapped beyond this. Default: no limit""}," &
                """image_attachment_name"":{""type"":""string"",""description"":""(image only) Filename of the image attachment to place (PNG, JPG, BMP, GIF, TIFF, WEBP)""}," &
                """width"":{""type"":""number"",""description"":""(image only) Width in points to scale the image to""}," &
                """height"":{""type"":""number"",""description"":""(image only) Height in points to scale the image to""}," &
                """rotation"":{""type"":""number"",""description"":""Rotation angle in degrees (clockwise). Default: 0""}," &
                """opacity"":{""type"":""number"",""description"":""Opacity from 0.0 (fully transparent) to 1.0 (fully opaque). Default: 1.0""}" &
                "}}," &
                """description"":""Array of overlay elements (text and/or image) to place on the PDF""}," &
                """output_filename"":{""type"":""string"",""description"":""Filename for the output PDF (default: '<original>_overlay.pdf')""}" &
                "},""required"":[""attachment_name"",""elements""]}}"
        })

        ' ── report_inability ──

        tools.Add(New ModelConfig() With {
            .ToolOnly = True, .Tool = True, .ToolName = AP_Tool_ReportInability,
            .ToolPriority = 9999,
            .ModelDescription = "Report Inability to Fulfill Request (built-in)",
            .ToolInstructionsPrompt =
                AP_Tool_ReportInability & ": Call this tool when you determine that you CANNOT fulfill the user's request " &
                "with the available tools and capabilities. Provide a brief reason. " &
                "The tool returns helpful suggestions for the user. You MUST naturally incorporate " &
                "the returned content into your reply — do NOT add labels, headers, or prefixes around it. " &
                "You MUST call this tool instead of simply telling the user you cannot help. " &
                "Also call this tool when attachments exceed the size limit and cannot be processed.",
            .ToolDefinition =
                "{""name"":""" & AP_Tool_ReportInability & """," &
                """description"":""Call this when you cannot fulfill the user's request. Provide the reason. " &
                "The tool returns helpful suggestions for the user. Naturally incorporate the returned text " &
                "into your reply without adding labels or headers around it. " &
                "Always call this instead of simply saying you cannot help.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """reason"":{""type"":""string"",""description"":""Brief reason why the request cannot be fulfilled (e.g. 'attachment exceeds size limit', 'no tool available for image generation', 'task requires manual interaction')""}" &
                "},""required"":[""reason""]}}"
        })

        Return tools

    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL DISPATCH
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Resolves and executes one AutoPilot internal tool call.
    ''' </summary>
    ''' <param name="toolCall">The parsed tool invocation payload.</param>
    ''' <param name="context">Execution context used for logging and correlation.</param>
    ''' <param name="cancellationToken">Optional cancellation token for async operations.</param>
    ''' <returns>
    ''' A <see cref="ToolResponse"/> when the tool is recognized; otherwise <c>Nothing</c>
    ''' so the caller can continue with external tool handling.
    ''' </returns>
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
            Case AP_Tool_CommentPdf
                Return Await ExecuteCommentPdfTool(toolCall, context, cancellationToken)
            Case AP_Tool_CompareWordDocs
                Return Await ExecuteCompareWordDocsTool(toolCall, context, cancellationToken)
            Case AP_Tool_ReadWordDocDetails
                Return Await ExecuteReadWordDocDetailsTool(toolCall, context, cancellationToken)
            Case AP_Tool_CreatePdfFromText
                Return ExecuteCreatePdfFromTextTool(toolCall, context)
            Case AP_Tool_ExtractExcelData
                Return ExecuteExtractExcelDataTool(toolCall, context)
            Case AP_Tool_SplitPdf
                Return ExecuteSplitPdfTool(toolCall, context)
            Case AP_Tool_AddPdfWatermark
                Return ExecuteAddPdfWatermarkTool(toolCall, context)
            Case AP_Tool_WordToPdf
                Return Await ExecuteWordToPdfTool(toolCall, context, cancellationToken)
            Case AP_Tool_SearchInAttachments
                Return Await ExecuteSearchInAttachmentsTool(toolCall, context, cancellationToken)
            Case AP_Tool_SummarizeThread
                Return ExecuteSummarizeThreadTool(toolCall, context)
            Case AP_Tool_PdfToWord
                Return Await ExecutePdfToWordTool(toolCall, context, cancellationToken)
            Case AP_Tool_CreateWordDoc
                Return Await ExecuteCreateWordDocTool(toolCall, context, cancellationToken)
            Case AP_Tool_CreateExcel
                Return Await ExecuteCreateExcelTool(toolCall, context, cancellationToken)
            Case AP_Tool_CreatePowerPoint
                Return Await ExecuteCreatePowerPointTool(toolCall, context, cancellationToken)
            Case AP_Tool_CreateCodeFile
                Return Await ExecuteCreateCodeFileTool(toolCall, context, cancellationToken)
            Case AP_Tool_ExtractDataFromAttachments
                Return Await ExecuteExtractDataFromAttachmentsTool(toolCall, context, cancellationToken)
            Case AP_Tool_RedactPdf
                Return Await ExecuteRedactPdfTool(toolCall, context, cancellationToken)
            Case AP_Tool_OverlayPdf
                Return Await ExecuteOverlayPdfTool(toolCall, context, cancellationToken)
            Case AP_Tool_ReportInability
                Return Await ExecuteReportInabilityTool(toolCall, context, cancellationToken)
            Case Else
                Return Nothing
        End Select
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Get argument values
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Reads a string argument from a tool-argument dictionary.
    ''' </summary>
    ''' <param name="args">Argument dictionary from the tool call.</param>
    ''' <param name="key">Argument key to read.</param>
    ''' <returns>The string value if present; otherwise <c>Nothing</c>.</returns>
    Private Shared Function GetArgString(args As Dictionary(Of String, Object), key As String) As String
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return Nothing
        Return args(key)?.ToString()
    End Function

    ''' <summary>
    ''' Reads a Boolean argument with fallback default.
    ''' </summary>
    Private Shared Function GetArgBool(args As Dictionary(Of String, Object), key As String, defaultVal As Boolean) As Boolean
        Dim s = GetArgString(args, key)
        If String.IsNullOrWhiteSpace(s) Then Return defaultVal
        Dim result As Boolean
        If Boolean.TryParse(s, result) Then Return result
        Return defaultVal
    End Function

    ''' <summary>
    ''' Reads an Integer argument with fallback default.
    ''' </summary>
    Private Shared Function GetArgInt(args As Dictionary(Of String, Object), key As String, defaultVal As Integer) As Integer
        Dim s = GetArgString(args, key)
        If String.IsNullOrWhiteSpace(s) Then Return defaultVal
        Dim result As Integer
        If Integer.TryParse(s, result) Then Return result
        Return defaultVal
    End Function

    ''' <summary>
    ''' Reads a JSON array argument as a list of strings.
    ''' </summary>
    Private Shared Function GetArgStringArray(args As Dictionary(Of String, Object), key As String) As List(Of String)
        Dim result As New List(Of String)()
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return result
        Dim namesObj = args(key)
        If TypeOf namesObj Is JArray Then
            For Each item In DirectCast(namesObj, JArray)
                result.Add(item.ToString())
            Next
        End If
        Return result
    End Function

    ''' <summary>
    ''' Finds an attachment by filename (case-insensitive) from either:
    ''' (1) original mail attachments, or
    ''' (2) output files produced by prior tool calls in the same run.
    ''' </summary>
    ''' <remarks>
    ''' If an output file is matched, this method returns a transient
    ''' <see cref="AutoPilotAttachmentInfo"/> marked with <c>IsToolOutput=True</c>.
    ''' </remarks>
    Private Function FindAttachment(fileName As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(fileName) OrElse _apCurrentAttachments Is Nothing Then Return Nothing

        Dim trimmedName = fileName.Trim()

        Dim found = _apCurrentAttachments.FirstOrDefault(
            Function(a) a.OriginalFileName.Equals(trimmedName, StringComparison.OrdinalIgnoreCase))
        If found IsNot Nothing Then Return found

        For Each att In _apCurrentAttachments
            If att.OutputFiles Is Nothing Then Continue For
            For Each outputPath In att.OutputFiles
                If String.IsNullOrEmpty(outputPath) Then Continue For
                Dim outputName = Path.GetFileName(outputPath)
                If outputName.Equals(trimmedName, StringComparison.OrdinalIgnoreCase) AndAlso
                   File.Exists(outputPath) Then
                    Return New AutoPilotAttachmentInfo() With {
                        .OriginalFileName = outputName,
                        .Extension = Path.GetExtension(outputPath).ToLowerInvariant(),
                        .TempFilePath = outputPath,
                        .SizeBytes = New FileInfo(outputPath).Length,
                        .IsOverSizeLimit = False,
                        .StatusMessage = "Tool output",
                        .IsToolOutput = True,
                        .OutputFiles = New List(Of String)()
                    }
                End If
            Next
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Returns all filenames currently available for tool resolution:
    ''' original attachments plus existing tool output files.
    ''' </summary>
    Private Function GetAllAvailableFileNames() As List(Of String)
        Dim names As New List(Of String)()
        If _apCurrentAttachments Is Nothing Then Return names
        For Each att In _apCurrentAttachments
            names.Add(att.OriginalFileName)
            If att.OutputFiles IsNot Nothing Then
                For Each outputPath In att.OutputFiles
                    If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                        names.Add(Path.GetFileName(outputPath))
                    End If
                Next
            End If
        Next
        Return names
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Read single attachment text (with caching)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Reads text from a single attachment, using cache when available.
    ''' </summary>
    Private Async Function ReadSingleAttachmentText(att As AutoPilotAttachmentInfo, context As ToolExecutionContext) As Task(Of String)
        ' Return cache if available
        If att.CachedText IsNot Nothing Then Return att.CachedText

        If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then Return Nothing

        Dim text As String = Nothing
        Dim label As String = Nothing
        Dim extracted As Boolean = False

        Try
            extracted = TryExtractOfficeText(att.TempFilePath, text, label)
        Catch
        End Try

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

        If Not extracted Then
            Try
                extracted = TryExtractTextLike(att.TempFilePath, text, label)
            Catch
            End Try
        End If

        If Not extracted AndAlso att.Extension = ".pdf" Then
            Try
                text = Await SharedMethods.ReadPdfAsText(att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=False, AskUser:=False)
                extracted = Not String.IsNullOrWhiteSpace(text)
            Catch
            End Try
        End If

        If extracted AndAlso Not String.IsNullOrWhiteSpace(text) Then
            att.CachedText = text
            Return text
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Detects comment and tracked change counts in a .docx for hinting in read_attachment.
    ''' Result is cached in att.CachedDocxHint.
    ''' </summary>
    Private Function GetDocxMetadataHint(att As AutoPilotAttachmentInfo) As String
        If att.CachedDocxHint IsNot Nothing Then Return att.CachedDocxHint

        If att.Extension <> ".docx" OrElse att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
            att.CachedDocxHint = ""
            Return ""
        End If

        Try
            Dim commentCount As Integer = 0
            Dim revisionCount As Integer = 0
            Dim tempDir = Path.Combine(Path.GetTempPath(), "ap_hint_" & Guid.NewGuid().ToString("N"))
            ZipFile.ExtractToDirectory(att.TempFilePath, tempDir)

            Dim commentsPath = Path.Combine(tempDir, "word", "comments.xml")
            If File.Exists(commentsPath) Then
                Dim commDoc As New XmlDocument()
                commDoc.Load(commentsPath)
                Dim nsMgr As New XmlNamespaceManager(commDoc.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                commentCount = commDoc.SelectNodes("//w:comment", nsMgr).Count
            End If

            Dim docPath = Path.Combine(tempDir, "word", "document.xml")
            If File.Exists(docPath) Then
                Dim docXml As New XmlDocument()
                docXml.Load(docPath)
                Dim nsMgr As New XmlNamespaceManager(docXml.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                revisionCount = docXml.SelectNodes("//w:ins", nsMgr).Count +
                                docXml.SelectNodes("//w:del", nsMgr).Count +
                                docXml.SelectNodes("//w:rPrChange", nsMgr).Count
            End If

            Try : Directory.Delete(tempDir, True) : Catch : End Try

            Dim hint As String = ""
            If commentCount > 0 OrElse revisionCount > 0 Then
                Dim parts As New List(Of String)()
                If commentCount > 0 Then parts.Add($"{commentCount} comment(s)")
                If revisionCount > 0 Then parts.Add($"{revisionCount} tracked change(s)")
                hint = $"(This document contains {String.Join(" and ", parts)}. Use {AP_Tool_ReadWordDocDetails} to inspect them.)"
            End If

            att.CachedDocxHint = hint
            Return hint
        Catch
            att.CachedDocxHint = ""
            Return ""
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: comment_pdf_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCommentPdfTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim author = GetArgString(toolCall.Arguments, "author")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(
                        Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)
                    ) AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable PDF attachments found."
                Return response
            End If

            Dim effectiveAuthor = If(String.IsNullOrWhiteSpace(author), AN6, author.Trim())
            Dim authorNote = If(effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase), "", $" (author: {effectiveAuthor})")
            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Adding PDF comments to: {att.OriginalFileName} with instruction: {instruction}{authorNote}")
                ApDashboardLog($"💬 Adding PDF comments to: {att.OriginalFileName}{authorNote}", "step")

                If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                    resultMessages.Add($"✗ {att.OriginalFileName}: Only PDF files are supported for PDF comment insertion.")
                    Continue For
                End If

                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_commented.pdf"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                ' Prevent filename collision
                Dim counter = 1
                While File.Exists(outputPath)
                    outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_commented_{counter}.pdf"
                    outputPath = Path.Combine(_apCurrentTempDir, outputName)
                    counter += 1
                End While

                Dim success = Await CommentPdfForAutoPilot(att.TempFilePath, outputPath, instruction, ct, author)

                If success Then
                    att.OutputFiles.Add(outputPath)
                    resultMessages.Add($"✓ {att.OriginalFileName}: PDF comments added successfully. Output: {outputName}")
                    ApDashboardLog($"✓ PDF comments added to: {att.OriginalFileName}", "info")
                Else
                    resultMessages.Add($"✗ {att.OriginalFileName}: Failed to add PDF comments (document may be empty, image-only, or unsupported).")
                    ApDashboardLog($"⚠ Failed to add PDF comments to: {att.OriginalFileName}", "warn")
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
            response.Response = $"Error adding comments to PDF(s): {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_code_file
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateCodeFileTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            Dim content = GetArgString(toolCall.Arguments, "content")
            Dim description = GetArgString(toolCall.Arguments, "description")

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: file_name"
                Return response
            End If

            If String.IsNullOrWhiteSpace(content) Then
                response.Success = False
                response.Response = "Missing required parameter: content"
                Return response
            End If

            ' Sanitize filename — preserve the extension but clean invalid chars
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            fileName = fileName.Trim()

            ' Ensure the file has an extension; default to .txt if none provided
            If String.IsNullOrWhiteSpace(Path.GetExtension(fileName)) Then
                fileName &= ".txt"
            End If

            ' Guard against binary/Office extensions that should use dedicated tools
            Dim ext = Path.GetExtension(fileName).ToLowerInvariant()
            Dim blockedExtensions = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".exe", ".dll", ".zip", ".rar"}
            If blockedExtensions.Contains(ext) Then
                response.Success = False
                response.Response = $"Cannot create binary file with extension '{ext}' using this tool. " &
                    "Use the dedicated tools (create_word_document, create_excel_spreadsheet, create_powerpoint, create_pdf_from_text) instead."
                Return response
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                Dim extension = Path.GetExtension(fileName)
                fileName = baseName & $"_{counter}{extension}"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            context.Log($"Creating code file: {fileName}")
            ApDashboardLog($"💻 Creating code file: {fileName}", "step")

            ' Write the file with UTF-8 encoding (with BOM for maximum compatibility)
            Await Task.Run(Sub() File.WriteAllText(outputPath, content, Encoding.UTF8), ct)

            If File.Exists(outputPath) Then
                ' Register as output on the first attachment if available
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(outputPath)
                End If

                Dim sizeKb = New FileInfo(outputPath).Length / 1024
                Dim lineCount = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None).Length

                Dim resultMsg As New StringBuilder()
                resultMsg.Append($"Code file created: {fileName} ({lineCount} lines, {sizeKb:F0} KB)")
                If Not String.IsNullOrWhiteSpace(description) Then
                    resultMsg.Append($". {description}")
                End If
                resultMsg.Append(". The file will be attached to the reply.")

                response.Success = True
                response.Response = resultMsg.ToString()
                ApDashboardLog($"✓ Code file created: {fileName} ({lineCount} lines)", "info")
            Else
                response.Success = False
                response.Response = "Failed to create code file."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating code file: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_powerpoint
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreatePowerPointTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            ' Parse slides array
            Dim slidesArray As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("slides") Then
                Dim slidesObj = toolCall.Arguments("slides")
                If TypeOf slidesObj Is JArray Then
                    slidesArray = DirectCast(slidesObj, JArray)
                End If
            End If

            If slidesArray Is Nothing OrElse slidesArray.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: slides (must be a non-empty array of slide objects)"
                Return response
            End If

            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Presentation"

            ' Sanitize filename
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            If Not fileName.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase) Then
                fileName &= ".pptx"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}.pptx"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            Dim presTitle = GetArgString(toolCall.Arguments, "title")

            context.Log($"Creating PowerPoint presentation: {fileName} ({slidesArray.Count} slides)")
            ApDashboardLog($"📊 Creating PowerPoint: {fileName}", "step")

            ' ppLayoutText = 2, ppLayoutTitleOnly = 11, ppLayoutBlank = 12, ppLayoutTitle = 1
            Const ppLayoutTitle As Integer = 1
            Const ppLayoutText As Integer = 2
            ' ppSaveAsOpenXMLPresentation = 24
            Const ppSaveAsOpenXMLPresentation As Integer = 24

            Dim success = Await SwitchToUi(Function()
                                               Dim app As Object = Nothing
                                               Dim pres As Object = Nothing
                                               Dim weOwnApp As Boolean = False
                                               Try
                                                   ' Late binding: no PIAs required (same as ExtractPowerPointText)
                                                   ' Try to get an existing instance first
                                                   Try
                                                       app = System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application")
                                                   Catch ex As System.Runtime.InteropServices.COMException
                                                       app = Microsoft.VisualBasic.Interaction.CreateObject("PowerPoint.Application")
                                                       weOwnApp = True
                                                   End Try

                                                   pres = app.Presentations.Add(0) ' 0 = WithWindow:=False

                                                   ' Set presentation title metadata if provided
                                                   If Not String.IsNullOrWhiteSpace(presTitle) Then
                                                       Try
                                                           pres.BuiltInDocumentProperties("Title").Value = presTitle
                                                       Catch
                                                       End Try
                                                   End If

                                                   Dim slideIndex As Integer = 0
                                                   For Each slideObj As JObject In slidesArray
                                                       slideIndex += 1

                                                       Dim title = slideObj.Value(Of String)("title")
                                                       Dim body = slideObj.Value(Of String)("body")
                                                       Dim notes = slideObj.Value(Of String)("notes")

                                                       ' First slide uses title layout, rest use text layout
                                                       Dim layoutType As Integer = If(slideIndex = 1, ppLayoutTitle, ppLayoutText)

                                                       Dim sld As Object = Nothing
                                                       Try
                                                           sld = pres.Slides.Add(slideIndex, layoutType)

                                                           ' Set title
                                                           If Not String.IsNullOrWhiteSpace(title) Then
                                                               Try
                                                                   sld.Shapes(1).TextFrame.TextRange.Text = title
                                                               Catch
                                                               End Try
                                                           End If

                                                           ' Set body text (placeholder 2)
                                                           If Not String.IsNullOrWhiteSpace(body) Then
                                                               ' Strip Markdown bullet markers — PowerPoint already formats as bullets
                                                               Dim cleanedLines As New List(Of String)()
                                                               For Each bodyLine In body.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                                                                   Dim trimmed = bodyLine.TrimStart()
                                                                   If trimmed.StartsWith("- ") Then
                                                                       trimmed = trimmed.Substring(2)
                                                                   ElseIf trimmed.StartsWith("* ") OrElse trimmed.StartsWith("+ ") Then
                                                                       trimmed = trimmed.Substring(2)
                                                                   ElseIf trimmed.Length > 2 AndAlso Char.IsDigit(trimmed(0)) Then
                                                                       Dim dotIdx = trimmed.IndexOf(". ")
                                                                       If dotIdx > 0 AndAlso dotIdx <= 3 Then
                                                                           Dim prefix = trimmed.Substring(0, dotIdx)
                                                                           Dim allDigits = True
                                                                           For Each ch In prefix
                                                                               If Not Char.IsDigit(ch) Then allDigits = False : Exit For
                                                                           Next
                                                                           If allDigits Then trimmed = trimmed.Substring(dotIdx + 2)
                                                                       End If
                                                                   End If
                                                                   cleanedLines.Add(trimmed)
                                                               Next
                                                               body = String.Join(vbCrLf, cleanedLines)

                                                               Try
                                                                   sld.Shapes(2).TextFrame.TextRange.Text = body
                                                               Catch
                                                                   ' Some layouts may not have a second placeholder;
                                                                   ' try adding as a text box instead
                                                                   Try
                                                                       ' AddTextbox(Orientation, Left, Top, Width, Height)
                                                                       ' 1 = msoTextOrientationHorizontal
                                                                       Dim tb As Object = sld.Shapes.AddTextbox(1, 50, 120, 600, 300)
                                                                       tb.TextFrame.TextRange.Text = body
                                                                       tb.TextFrame.WordWrap = -1 ' msoTrue
                                                                       Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tb)
                                                                       Catch : End Try
                                                                   Catch
                                                                   End Try
                                                               End Try
                                                           End If

                                                           ' Set speaker notes
                                                           If Not String.IsNullOrWhiteSpace(notes) Then
                                                               Try
                                                                   Dim notesPage As Object = sld.NotesPage
                                                                   Dim notesShapes As Object = notesPage.Shapes
                                                                   Dim nCount As Integer = System.Convert.ToInt32(notesShapes.Count,
                                                                       Globalization.CultureInfo.InvariantCulture)
                                                                   ' Find the body placeholder in notes (type 2 = ppPlaceholderBody)
                                                                   For k As Integer = 1 To nCount
                                                                       Dim nShp As Object = notesShapes(k)
                                                                       Try
                                                                           Dim phType As Integer = System.Convert.ToInt32(
                                                                               nShp.PlaceholderFormat.Type,
                                                                               Globalization.CultureInfo.InvariantCulture)
                                                                           If phType = 2 Then ' ppPlaceholderBody
                                                                               nShp.TextFrame.TextRange.Text = notes
                                                                               Exit For
                                                                           End If
                                                                       Catch
                                                                       Finally
                                                                           Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(nShp)
                                                                           Catch : End Try
                                                                       End Try
                                                                   Next
                                                               Catch
                                                               End Try
                                                           End If
                                                       Finally
                                                           Try
                                                               If sld IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld)
                                                           Catch
                                                           End Try
                                                       End Try
                                                   Next

                                                   ' SaveAs(FileName, FileFormat)
                                                   pres.SaveAs(outputPath, ppSaveAsOpenXMLPresentation)
                                                   Return True
                                               Catch ex As Exception
                                                   Debug.WriteLine($"CreatePowerPoint error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   Try
                                                       If pres IsNot Nothing Then
                                                           Try : pres.Close() : Catch : End Try
                                                           Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pres)
                                                           Catch : End Try
                                                       End If
                                                   Catch
                                                   End Try
                                                   Try
                                                       If app IsNot Nothing Then
                                                           ' Only quit if we created the instance ourselves
                                                           If weOwnApp Then
                                                               Try : app.Quit() : Catch : End Try
                                                           End If
                                                           Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app)
                                                           Catch : End Try
                                                       End If
                                                   Catch
                                                   End Try
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(outputPath)
                End If

                response.Success = True
                response.Response = $"PowerPoint presentation created: {fileName} ({slidesArray.Count} slides, {New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply."
                ApDashboardLog($"✓ PowerPoint created: {fileName}", "info")
            Else
                response.Success = False
                response.Response = "Failed to create PowerPoint presentation."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating PowerPoint presentation: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_excel_spreadsheet
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateExcelTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            ' ── Resolve sheet definitions ──
            ' Support both: top-level "cells" (single sheet) and "sheets" array (multi-sheet)
            Dim sheetDefs As New List(Of (SheetName As String, Cells As JArray))()
            Dim hasVba As Boolean = False

            ' Check for VBA modules — determines .xlsm vs .xlsx
            Dim vbaModules As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("vba_modules") Then
                Dim vbaObj = toolCall.Arguments("vba_modules")
                If TypeOf vbaObj Is JArray AndAlso DirectCast(vbaObj, JArray).Count > 0 Then
                    vbaModules = DirectCast(vbaObj, JArray)
                    hasVba = True
                End If
            End If

            ' Multi-sheet mode
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("sheets") Then
                Dim sheetsObj = toolCall.Arguments("sheets")
                If TypeOf sheetsObj Is JArray Then
                    For Each sheetObj As JObject In DirectCast(sheetsObj, JArray)
                        Dim sName = sheetObj.Value(Of String)("name")
                        If String.IsNullOrWhiteSpace(sName) Then sName = $"Sheet{sheetDefs.Count + 1}"
                        Dim sCells As JArray = Nothing
                        Dim sCellsToken = sheetObj("cells")
                        If TypeOf sCellsToken Is JArray Then sCells = DirectCast(sCellsToken, JArray)
                        If sCells IsNot Nothing AndAlso sCells.Count > 0 Then
                            sheetDefs.Add((sName, sCells))
                        End If
                    Next
                End If
            End If

            ' Single-sheet mode (backward compatible)
            If sheetDefs.Count = 0 Then
                Dim cellsArray As JArray = Nothing
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("cells") Then
                    Dim cellsObj = toolCall.Arguments("cells")
                    If TypeOf cellsObj Is JArray Then cellsArray = DirectCast(cellsObj, JArray)
                End If

                If cellsArray Is Nothing OrElse cellsArray.Count = 0 Then
                    response.Success = False
                    response.Response = "Missing required parameter: cells or sheets (must contain at least one non-empty cell array)"
                    Return response
                End If

                Dim sheetName = GetArgString(toolCall.Arguments, "sheet_name")
                If String.IsNullOrWhiteSpace(sheetName) Then sheetName = "Sheet1"
                sheetDefs.Add((sheetName, cellsArray))
            End If

            ' ── Determine file name and extension ──
            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Spreadsheet"
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next

            Dim fileExt As String = If(hasVba, ".xlsm", ".xlsx")
            If Not fileName.EndsWith(fileExt, StringComparison.OrdinalIgnoreCase) Then
                ' Strip wrong extension if present
                If fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) OrElse
                   fileName.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) Then
                    fileName = Path.GetFileNameWithoutExtension(fileName)
                End If
                fileName &= fileExt
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}{fileExt}"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            ' ── Parse shared parameters ──
            Dim columnWidths As Dictionary(Of String, Double) = ParseColumnWidths(toolCall.Arguments)
            Dim rowHeights As Dictionary(Of Integer, Double) = ParseRowHeights(toolCall.Arguments)
            Dim mergeRanges = GetArgStringArray(toolCall.Arguments, "merge_ranges")
            Dim freezePane = GetArgString(toolCall.Arguments, "freeze_pane")
            Dim autoFilter = GetArgString(toolCall.Arguments, "auto_filter")
            Dim dataValidations = ParseJsonArray(toolCall.Arguments, "data_validations")
            Dim conditionalFormats = ParseJsonArray(toolCall.Arguments, "conditional_formats")
            Dim charts = ParseJsonArray(toolCall.Arguments, "charts")
            Dim namedRanges = ParseJsonArray(toolCall.Arguments, "named_ranges")
            Dim printSetup As JObject = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("print_setup") Then
                Dim psObj = toolCall.Arguments("print_setup")
                If TypeOf psObj Is JObject Then printSetup = DirectCast(psObj, JObject)
            End If

            Dim totalCells = sheetDefs.Sum(Function(sd) sd.Cells.Count)
            context.Log($"Creating Excel spreadsheet: {fileName} ({sheetDefs.Count} sheet(s), {totalCells} cells)")
            ApDashboardLog($"📊 Creating Excel: {fileName} ({sheetDefs.Count} sheet(s))", "step")

            ' xlOpenXMLWorkbook = 51, xlOpenXMLWorkbookMacroEnabled = 52
            Const xlOpenXMLWorkbook As Integer = 51
            Const xlOpenXMLWorkbookMacroEnabled As Integer = 52

            Dim success = Await SwitchToUi(Function()
                                               Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
                                               Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                                               Dim weOwnApp As Boolean = False
                                               Try
                                                   ' Try to reuse an existing Excel instance
                                                   Try
                                                       excelApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"),
                                                                        Microsoft.Office.Interop.Excel.Application)
                                                   Catch ex As System.Runtime.InteropServices.COMException
                                                       excelApp = New Microsoft.Office.Interop.Excel.Application()
                                                       weOwnApp = True
                                                   End Try

                                                   excelApp.Visible = False
                                                   excelApp.DisplayAlerts = False
                                                   excelApp.ScreenUpdating = False

                                                   wb = excelApp.Workbooks.Add()

                                                   ' ── Create worksheets ──
                                                   ' Excel starts with 1 sheet by default; add more as needed
                                                   While wb.Sheets.Count < sheetDefs.Count
                                                       wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
                                                   End While

                                                   ' Remove extra default sheets
                                                   While wb.Sheets.Count > sheetDefs.Count
                                                       CType(wb.Sheets(wb.Sheets.Count), Microsoft.Office.Interop.Excel.Worksheet).Delete()
                                                   End While

                                                   For sheetIdx = 0 To sheetDefs.Count - 1
                                                       Dim ws = CType(wb.Sheets(sheetIdx + 1), Microsoft.Office.Interop.Excel.Worksheet)
                                                       Dim sheetDef = sheetDefs(sheetIdx)
                                                       ws.Name = sheetDef.SheetName

                                                       ' ── Apply cells ──
                                                       ApplyExcelCells(ws, sheetDef.Cells)

                                                       ' ── Column widths (apply per-sheet for first sheet, or if multi-sheet) ──
                                                       If sheetIdx = 0 AndAlso columnWidths IsNot Nothing Then
                                                           ApplyColumnWidths(ws, columnWidths)
                                                       End If

                                                       ' ── Row heights ──
                                                       If sheetIdx = 0 AndAlso rowHeights IsNot Nothing Then
                                                           ApplyRowHeights(ws, rowHeights)
                                                       End If

                                                       ' ── Merge ranges ──
                                                       If sheetIdx = 0 AndAlso mergeRanges IsNot Nothing Then
                                                           For Each mr In mergeRanges
                                                               Try : ws.Range(mr).Merge() : Catch : End Try
                                                           Next
                                                       End If

                                                       ' ── Freeze pane ──
                                                       If sheetIdx = 0 AndAlso Not String.IsNullOrWhiteSpace(freezePane) Then
                                                           Try
                                                               ws.Activate()
                                                               ws.Range(freezePane).Select()
                                                               excelApp.ActiveWindow.FreezePanes = True
                                                           Catch
                                                           End Try
                                                       End If

                                                       ' ── Auto-filter ──
                                                       If sheetIdx = 0 AndAlso Not String.IsNullOrWhiteSpace(autoFilter) Then
                                                           Try : ws.Range(autoFilter).AutoFilter() : Catch : End Try
                                                       End If

                                                       ' ── Data validations ──
                                                       If sheetIdx = 0 AndAlso dataValidations IsNot Nothing Then
                                                           ApplyDataValidations(ws, dataValidations)
                                                       End If

                                                       ' ── Conditional formatting ──
                                                       If sheetIdx = 0 AndAlso conditionalFormats IsNot Nothing Then
                                                           ApplyConditionalFormats(ws, conditionalFormats)
                                                       End If

                                                       ' ── Print setup ──
                                                       If sheetIdx = 0 AndAlso printSetup IsNot Nothing Then
                                                           ApplyPrintSetup(ws, printSetup)
                                                       End If
                                                   Next

                                                   ' ── Charts (can target any sheet) ──
                                                   If charts IsNot Nothing Then
                                                       ApplyCharts(wb, charts, sheetDefs)
                                                   End If

                                                   ' ── Named ranges ──
                                                   If namedRanges IsNot Nothing Then
                                                       For Each nrObj As JObject In namedRanges
                                                           Try
                                                               Dim nrName = nrObj.Value(Of String)("name")
                                                               Dim nrRange = nrObj.Value(Of String)("range")
                                                               If Not String.IsNullOrWhiteSpace(nrName) AndAlso Not String.IsNullOrWhiteSpace(nrRange) Then
                                                                   wb.Names.Add(Name:=nrName, RefersTo:="=" & nrRange)
                                                               End If
                                                           Catch
                                                           End Try
                                                       Next
                                                   End If

                                                   ' ── VBA modules ──
                                                   If hasVba AndAlso vbaModules IsNot Nothing Then
                                                       ApplyVbaModules(wb, vbaModules)
                                                   End If

                                                   ' ── Save ──
                                                   Dim fmt = If(hasVba, xlOpenXMLWorkbookMacroEnabled, xlOpenXMLWorkbook)
                                                   wb.SaveAs(outputPath, fmt)
                                                   Return True

                                               Catch ex As Exception
                                                   Debug.WriteLine($"CreateExcel error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   SafeCloseExcel(wb, excelApp, weOwnApp)
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(outputPath)
                End If

                Dim featureList As New List(Of String)()
                If sheetDefs.Count > 1 Then featureList.Add($"{sheetDefs.Count} sheets")
                featureList.Add($"{totalCells} cells")
                If mergeRanges.Count > 0 Then featureList.Add($"{mergeRanges.Count} merged range(s)")
                If dataValidations IsNot Nothing AndAlso dataValidations.Count > 0 Then featureList.Add($"{dataValidations.Count} validation(s)")
                If conditionalFormats IsNot Nothing AndAlso conditionalFormats.Count > 0 Then featureList.Add($"{conditionalFormats.Count} conditional format(s)")
                If charts IsNot Nothing AndAlso charts.Count > 0 Then featureList.Add($"{charts.Count} chart(s)")
                If hasVba Then featureList.Add("VBA macros")

                response.Success = True
                response.Response = $"Excel spreadsheet created: {fileName} ({String.Join(", ", featureList)}, {New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply."
                ApDashboardLog($"✓ Excel created: {fileName}", "info")
            Else
                response.Success = False
                response.Response = "Failed to create Excel spreadsheet."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating Excel spreadsheet: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  EXCEL CREATION HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses a hex color string like "#FF0000" or "FF0000" to an OLE color integer.
    ''' Returns Nothing if parsing fails.
    ''' </summary>
    Private Shared Function ParseHexColor(hexStr As String) As Integer?
        If String.IsNullOrWhiteSpace(hexStr) Then Return Nothing
        hexStr = hexStr.TrimStart("#"c)
        If hexStr.Length <> 6 Then Return Nothing
        Try
            Dim r = System.Convert.ToInt32(hexStr.Substring(0, 2), 16)
            Dim g = System.Convert.ToInt32(hexStr.Substring(2, 2), 16)
            Dim b = System.Convert.ToInt32(hexStr.Substring(4, 2), 16)
            ' Excel uses BGR (OLE color) format
            Return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b))
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Parses column_widths from tool arguments.
    ''' </summary>
    Private Shared Function ParseColumnWidths(args As Dictionary(Of String, Object)) As Dictionary(Of String, Double)
        If args Is Nothing OrElse Not args.ContainsKey("column_widths") Then Return Nothing
        Dim cwObj = args("column_widths")
        If Not TypeOf cwObj Is JObject Then Return Nothing
        Dim result As New Dictionary(Of String, Double)(StringComparer.OrdinalIgnoreCase)
        For Each prop In DirectCast(cwObj, JObject).Properties()
            Dim w As Double
            If Double.TryParse(prop.Value.ToString(), Globalization.NumberStyles.Any,
                              Globalization.CultureInfo.InvariantCulture, w) Then
                result(prop.Name.ToUpperInvariant()) = w
            End If
        Next
        Return If(result.Count > 0, result, Nothing)
    End Function

    ''' <summary>
    ''' Parses row_heights from tool arguments.
    ''' </summary>
    Private Shared Function ParseRowHeights(args As Dictionary(Of String, Object)) As Dictionary(Of Integer, Double)
        If args Is Nothing OrElse Not args.ContainsKey("row_heights") Then Return Nothing
        Dim rhObj = args("row_heights")
        If Not TypeOf rhObj Is JObject Then Return Nothing
        Dim result As New Dictionary(Of Integer, Double)()
        For Each prop In DirectCast(rhObj, JObject).Properties()
            Dim rowNum As Integer
            Dim h As Double
            If Integer.TryParse(prop.Name, rowNum) AndAlso
               Double.TryParse(prop.Value.ToString(), Globalization.NumberStyles.Any,
                              Globalization.CultureInfo.InvariantCulture, h) Then
                result(rowNum) = h
            End If
        Next
        Return If(result.Count > 0, result, Nothing)
    End Function

    ''' <summary>
    ''' Parses a JSON array from tool arguments by key name.
    ''' </summary>
    Private Shared Function ParseJsonArray(args As Dictionary(Of String, Object), key As String) As List(Of JObject)
        If args Is Nothing OrElse Not args.ContainsKey(key) Then Return Nothing
        Dim obj = args(key)
        If Not TypeOf obj Is JArray Then Return Nothing
        Dim arr = DirectCast(obj, JArray)
        If arr.Count = 0 Then Return Nothing
        Return arr.OfType(Of JObject)().ToList()
    End Function

    ''' <summary>
    ''' Applies cell data, values, formulas, and rich formatting to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyExcelCells(ws As Microsoft.Office.Interop.Excel.Worksheet, cellsArray As JArray)
        For Each cellObj As JObject In cellsArray
            Dim addr = cellObj.Value(Of String)("cell")
            If String.IsNullOrWhiteSpace(addr) Then Continue For

            Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                cell = ws.Range(addr)
            Catch
                Continue For
            End Try

            ' ── Number format (apply before value so formatting takes effect) ──
            Dim numFmt = cellObj.Value(Of String)("number_format")
            If Not String.IsNullOrWhiteSpace(numFmt) Then
                Try : cell.NumberFormat = numFmt : Catch : End Try
            End If

            ' ── Formula or value ──
            Dim formula = cellObj.Value(Of String)("formula")
            If Not String.IsNullOrWhiteSpace(formula) Then
                Try
                    cell.Formula2 = formula
                Catch
                    Try : cell.Formula = formula
                    Catch ex2 As Exception
                        Debug.WriteLine($"Formula error at {addr}: {ex2.Message}")
                    End Try
                End Try
            Else
                Dim valToken = cellObj("value")
                If valToken IsNot Nothing Then
                    Dim valStr = valToken.ToString()
                    Dim numVal As Double
                    If Double.TryParse(valStr, Globalization.NumberStyles.Any,
                                      Globalization.CultureInfo.InvariantCulture, numVal) Then
                        cell.Value2 = numVal
                    Else
                        cell.Value2 = valStr
                    End If
                End If
            End If

            ' ── Font styles ──
            If GetJBool(cellObj, "bold") Then Try : cell.Font.Bold = True : Catch : End Try
            If GetJBool(cellObj, "italic") Then Try : cell.Font.Italic = True : Catch : End Try
            If GetJBool(cellObj, "underline") Then Try : cell.Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle : Catch : End Try
            If GetJBool(cellObj, "strikethrough") Then Try : cell.Font.Strikethrough = True : Catch : End Try

            Dim fontName = cellObj.Value(Of String)("font_name")
            If Not String.IsNullOrWhiteSpace(fontName) Then Try : cell.Font.Name = fontName : Catch : End Try

            Dim fontSizeToken = cellObj("font_size")
            If fontSizeToken IsNot Nothing Then
                Dim fs As Double
                If Double.TryParse(fontSizeToken.ToString(), Globalization.NumberStyles.Any,
                                  Globalization.CultureInfo.InvariantCulture, fs) AndAlso fs > 0 Then
                    Try : cell.Font.Size = fs : Catch : End Try
                End If
            End If

            ' ── Font color ──
            Dim fontColor = ParseHexColor(cellObj.Value(Of String)("font_color"))
            If fontColor.HasValue Then Try : cell.Font.Color = fontColor.Value : Catch : End Try

            ' ── Background color ──
            Dim bgColor = ParseHexColor(cellObj.Value(Of String)("bg_color"))
            If bgColor.HasValue Then
                Try
                    cell.Interior.Color = bgColor.Value
                    cell.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                Catch
                End Try
            End If

            ' ── Alignment ──
            Dim hAlign = cellObj.Value(Of String)("h_align")
            If Not String.IsNullOrWhiteSpace(hAlign) Then
                Try
                    Select Case hAlign.ToLowerInvariant()
                        Case "left" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                        Case "center" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        Case "right" : cell.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                    End Select
                Catch
                End Try
            End If

            Dim vAlign = cellObj.Value(Of String)("v_align")
            If Not String.IsNullOrWhiteSpace(vAlign) Then
                Try
                    Select Case vAlign.ToLowerInvariant()
                        Case "top" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop
                        Case "center" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                        Case "bottom" : cell.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom
                    End Select
                Catch
                End Try
            End If

            If GetJBool(cellObj, "wrap_text") Then Try : cell.WrapText = True : Catch : End Try

            ' ── Borders ──
            Dim borderStyle = cellObj.Value(Of String)("border")
            If Not String.IsNullOrWhiteSpace(borderStyle) Then
                Dim borderColor = ParseHexColor(cellObj.Value(Of String)("border_color"))
                ApplyBorderStyle(cell, borderStyle, borderColor)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Helper to read a boolean from a JObject token.
    ''' </summary>
    Private Shared Function GetJBool(obj As JObject, key As String) As Boolean
        Dim token = obj(key)
        If token Is Nothing Then Return False
        If token.Type = JTokenType.Boolean Then Return CBool(token)
        Dim s = token.ToString()
        Dim result As Boolean
        If Boolean.TryParse(s, result) Then Return result
        Return False
    End Function

    ''' <summary>
    ''' Applies border styles to a cell range.
    ''' </summary>
    Private Shared Sub ApplyBorderStyle(cell As Microsoft.Office.Interop.Excel.Range,
                                         borderStyle As String, borderColor As Integer?)
        ' Map style names to Excel line style and weight
        Dim lineStyle As Microsoft.Office.Interop.Excel.XlLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
        Dim weight As Microsoft.Office.Interop.Excel.XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

        Dim style = borderStyle.ToLowerInvariant()

        If style.Contains("medium") Then
            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
        ElseIf style.Contains("thick") Then
            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
        End If

        Try
            If style.StartsWith("all") OrElse style = "thin" OrElse style = "medium" OrElse style = "thick" Then
                ' All four sides
                Dim edges() As Microsoft.Office.Interop.Excel.XlBordersIndex = {
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
                }
                For Each edge In edges
                    cell.Borders(edge).LineStyle = lineStyle
                    cell.Borders(edge).Weight = weight
                    If borderColor.HasValue Then cell.Borders(edge).Color = borderColor.Value
                Next
            ElseIf style.StartsWith("bottom") Then
                cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = lineStyle
                cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Weight = weight
                If borderColor.HasValue Then cell.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).Color = borderColor.Value
            ElseIf style.StartsWith("outline") Then
                Dim edges() As Microsoft.Office.Interop.Excel.XlBordersIndex = {
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom,
                    Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight
                }
                For Each edge In edges
                    cell.Borders(edge).LineStyle = lineStyle
                    cell.Borders(edge).Weight = weight
                    If borderColor.HasValue Then cell.Borders(edge).Color = borderColor.Value
                Next
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Applies column widths to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyColumnWidths(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                          widths As Dictionary(Of String, Double))
        For Each kv In widths
            Try
                Dim colRange = ws.Columns(kv.Key & ":" & kv.Key)
                colRange.ColumnWidth = kv.Value
            Catch
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies row heights to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyRowHeights(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                        heights As Dictionary(Of Integer, Double))
        For Each kv In heights
            Try
                ws.Rows(kv.Key).RowHeight = kv.Value
            Catch
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies data validation rules to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyDataValidations(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                              validations As List(Of JObject))
        For Each dvObj In validations
            Try
                Dim rangeName = dvObj.Value(Of String)("range")
                If String.IsNullOrWhiteSpace(rangeName) Then Continue For

                Dim dvRange = ws.Range(rangeName)
                dvRange.Validation.Delete() ' Clear existing validation

                Dim dvType = If(dvObj.Value(Of String)("type"), "list").ToLowerInvariant()
                Dim formula1 = dvObj.Value(Of String)("formula1")
                Dim formula2 = dvObj.Value(Of String)("formula2")
                Dim operatorStr = If(dvObj.Value(Of String)("operator"), "between").ToLowerInvariant()

                ' Map type to Excel constant
                Dim xlType As Integer
                Select Case dvType
                    Case "list" : xlType = 3 ' xlValidateList
                    Case "whole_number" : xlType = 1 ' xlValidateWholeNumber
                    Case "decimal" : xlType = 2 ' xlValidateDecimal
                    Case "date" : xlType = 4 ' xlValidateDate
                    Case "text_length" : xlType = 6 ' xlValidateTextLength
                    Case "custom" : xlType = 7 ' xlValidateCustom
                    Case Else : xlType = 3
                End Select

                ' Map operator to Excel constant
                Dim xlOp As Integer = 1 ' xlBetween
                Select Case operatorStr
                    Case "between" : xlOp = 1
                    Case "not_between" : xlOp = 2
                    Case "equal" : xlOp = 3
                    Case "not_equal" : xlOp = 4
                    Case "greater_than" : xlOp = 5
                    Case "less_than" : xlOp = 6
                    Case "greater_than_or_equal" : xlOp = 7
                    Case "less_than_or_equal" : xlOp = 8
                End Select

                If dvType = "list" Then
                    ' For list validation, formula1 is the comma-separated list.
                    ' LLMs sometimes wrap individual items in quotes (e.g., "Yes","No","Maybe")
                    ' which causes the first dropdown item to start with " and the last to end with ".
                    ' Strip any such quoting to get clean values for Excel.
                    Dim cleanedFormula1 = formula1
                    If Not String.IsNullOrWhiteSpace(cleanedFormula1) Then
                        ' Remove quotes wrapping individual items: "Yes","No" → Yes,No
                        Dim parts = cleanedFormula1.Split(","c)
                        For i = 0 To parts.Length - 1
                            parts(i) = parts(i).Trim().Trim(""""c).Trim("'"c)
                        Next
                        cleanedFormula1 = String.Join(",", parts)
                    End If
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Formula1:=cleanedFormula1)
                ElseIf Not String.IsNullOrWhiteSpace(formula2) Then
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Operator:=xlOp,
                                           Formula1:=formula1, Formula2:=formula2)
                Else
                    dvRange.Validation.Add(Type:=xlType, AlertStyle:=1,
                                           Operator:=xlOp,
                                           Formula1:=formula1)
                End If

                ' Show dropdown for list type
                Dim showDropdown = dvObj("show_dropdown")
                If showDropdown IsNot Nothing AndAlso showDropdown.Type = JTokenType.Boolean Then
                    dvRange.Validation.InCellDropdown = CBool(showDropdown)
                End If

                ' Input message
                Dim inputTitle = dvObj.Value(Of String)("input_title")
                Dim inputMsg = dvObj.Value(Of String)("input_message")
                If Not String.IsNullOrWhiteSpace(inputTitle) OrElse Not String.IsNullOrWhiteSpace(inputMsg) Then
                    dvRange.Validation.ShowInput = True
                    If Not String.IsNullOrWhiteSpace(inputTitle) Then dvRange.Validation.InputTitle = inputTitle
                    If Not String.IsNullOrWhiteSpace(inputMsg) Then dvRange.Validation.InputMessage = inputMsg
                End If

                ' Error message
                Dim errorTitle = dvObj.Value(Of String)("error_title")
                Dim errorMsg = dvObj.Value(Of String)("error_message")
                If Not String.IsNullOrWhiteSpace(errorTitle) OrElse Not String.IsNullOrWhiteSpace(errorMsg) Then
                    dvRange.Validation.ShowError = True
                    If Not String.IsNullOrWhiteSpace(errorTitle) Then dvRange.Validation.ErrorTitle = errorTitle
                    If Not String.IsNullOrWhiteSpace(errorMsg) Then dvRange.Validation.ErrorMessage = errorMsg
                End If

            Catch ex As Exception
                Debug.WriteLine($"Data validation error: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies conditional formatting rules to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyConditionalFormats(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                formats As List(Of JObject))
        For Each cfObj In formats
            Try
                Dim rangeName = cfObj.Value(Of String)("range")
                If String.IsNullOrWhiteSpace(rangeName) Then Continue For

                Dim cfRange = ws.Range(rangeName)
                Dim cfType = If(cfObj.Value(Of String)("type"), "cell_value").ToLowerInvariant()
                Dim operatorStr = If(cfObj.Value(Of String)("operator"), "greater_than").ToLowerInvariant()
                Dim formula1 = cfObj.Value(Of String)("formula1")
                Dim formula2 = cfObj.Value(Of String)("formula2")

                ' Map operator
                Dim xlOp As Integer = 5 ' xlGreater
                Select Case operatorStr
                    Case "between" : xlOp = 1
                    Case "not_between" : xlOp = 2
                    Case "equal" : xlOp = 3
                    Case "not_equal" : xlOp = 4
                    Case "greater_than" : xlOp = 5
                    Case "less_than" : xlOp = 6
                    Case "greater_than_or_equal" : xlOp = 7
                    Case "less_than_or_equal" : xlOp = 8
                End Select

                Dim fc As Microsoft.Office.Interop.Excel.FormatCondition = Nothing

                Select Case cfType
                    Case "cell_value"
                        If Not String.IsNullOrWhiteSpace(formula2) Then
                            fc = CType(cfRange.FormatConditions.Add(
                                Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                Operator:=xlOp, Formula1:=formula1, Formula2:=formula2),
                                Microsoft.Office.Interop.Excel.FormatCondition)
                        Else
                            fc = CType(cfRange.FormatConditions.Add(
                                Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                Operator:=xlOp, Formula1:=formula1),
                                Microsoft.Office.Interop.Excel.FormatCondition)
                        End If

                    Case "text_contains"
                        fc = CType(cfRange.FormatConditions.Add(
                            Type:=Microsoft.Office.Interop.Excel.XlFormatConditionType.xlTextString,
                            TextOperator:=Microsoft.Office.Interop.Excel.XlContainsOperator.xlContains,
                            String:=formula1),
                            Microsoft.Office.Interop.Excel.FormatCondition)

                    Case "duplicate"
                        fc = CType(cfRange.FormatConditions.AddUniqueValues(),
                            Microsoft.Office.Interop.Excel.UniqueValues)
                        CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).DupeUnique = Microsoft.Office.Interop.Excel.XlDupeUnique.xlDuplicate
                        ' UniqueValues doesn't have the same format interface; apply formatting directly
                        Dim fmtBgColor = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColor.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Interior.Color = fmtBgColor.Value : Catch : End Try
                        End If
                        Dim fmtFontColor = ParseHexColor(cfObj.Value(Of String)("format_font_color"))
                        If fmtFontColor.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Font.Color = fmtFontColor.Value : Catch : End Try
                        End If
                        Continue For ' Skip standard formatting below

                    Case "unique"
                        fc = CType(cfRange.FormatConditions.AddUniqueValues(),
                            Microsoft.Office.Interop.Excel.UniqueValues)
                        CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).DupeUnique = Microsoft.Office.Interop.Excel.XlDupeUnique.xlUnique
                        Dim fmtBgColorU = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColorU.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.UniqueValues).Interior.Color = fmtBgColorU.Value : Catch : End Try
                        End If
                        Continue For

                    Case "color_scale"
                        cfRange.FormatConditions.AddColorScale(ColorScaleType:=3) ' 3-color scale
                        Continue For

                    Case "data_bar"
                        cfRange.FormatConditions.AddDatabar()
                        Continue For

                    Case "icon_set"
                        cfRange.FormatConditions.AddIconSetCondition()
                        Continue For

                    Case "top_10"
                        fc = CType(cfRange.FormatConditions.AddTop10(),
                            Microsoft.Office.Interop.Excel.Top10)
                        Dim rank As Integer = 10
                        If Not String.IsNullOrWhiteSpace(formula1) Then
                            Integer.TryParse(formula1, rank)
                        End If
                        CType(fc, Microsoft.Office.Interop.Excel.Top10).Rank = rank
                        Dim fmtBgColorT = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                        If fmtBgColorT.HasValue Then
                            Try : CType(fc, Microsoft.Office.Interop.Excel.Top10).Interior.Color = fmtBgColorT.Value : Catch : End Try
                        End If
                        Continue For

                    Case Else
                        Continue For
                End Select

                ' Apply formatting to the FormatCondition
                If fc IsNot Nothing Then
                    Dim fmtFontColor = ParseHexColor(cfObj.Value(Of String)("format_font_color"))
                    If fmtFontColor.HasValue Then Try : fc.Font.Color = fmtFontColor.Value : Catch : End Try

                    Dim fmtBgColor = ParseHexColor(cfObj.Value(Of String)("format_bg_color"))
                    If fmtBgColor.HasValue Then
                        Try
                            fc.Interior.Color = fmtBgColor.Value
                            fc.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        Catch
                        End Try
                    End If

                    If GetJBool(cfObj, "format_bold") Then Try : fc.Font.Bold = True : Catch : End Try
                End If

            Catch ex As Exception
                Debug.WriteLine($"Conditional format error: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Creates charts and places them on worksheets.
    ''' </summary>
    Private Shared Sub ApplyCharts(wb As Microsoft.Office.Interop.Excel.Workbook,
                                    charts As List(Of JObject),
                                    sheetDefs As List(Of (SheetName As String, Cells As JArray)))
        For Each chartObj In charts
            Try
                Dim chartType = If(chartObj.Value(Of String)("type"), "column").ToLowerInvariant()
                Dim dataRange = chartObj.Value(Of String)("data_range")
                Dim chartTitle = chartObj.Value(Of String)("title")
                Dim position = If(chartObj.Value(Of String)("position"), "E2")
                Dim chartSheetName = chartObj.Value(Of String)("sheet_name")

                If String.IsNullOrWhiteSpace(dataRange) Then Continue For

                ' Determine target worksheet
                Dim targetWs As Microsoft.Office.Interop.Excel.Worksheet
                If Not String.IsNullOrWhiteSpace(chartSheetName) Then
                    Try
                        targetWs = CType(wb.Sheets(chartSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch
                        targetWs = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                    End Try
                Else
                    targetWs = CType(wb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                End If

                ' Parse width/height
                Dim chartWidth As Double = 480
                Dim chartHeight As Double = 300
                Dim wToken = chartObj("width")
                If wToken IsNot Nothing Then Double.TryParse(wToken.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, chartWidth)
                Dim hToken = chartObj("height")
                If hToken IsNot Nothing Then Double.TryParse(hToken.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, chartHeight)

                ' Get position from cell
                Dim posCell = targetWs.Range(position)
                Dim posLeft As Double = CDbl(posCell.Left)
                Dim posTop As Double = CDbl(posCell.Top)

                ' Map chart type to Excel constant
                Dim xlChartType As Microsoft.Office.Interop.Excel.XlChartType
                Select Case chartType
                    Case "column" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                    Case "bar" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered
                    Case "line" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlLine
                    Case "pie" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie
                    Case "area" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlArea
                    Case "scatter" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter
                    Case "doughnut" : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut
                    Case Else : xlChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                End Select

                ' Add chart as embedded ChartObject
                Dim chartObjects = CType(targetWs.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
                Dim chartObject = chartObjects.Add(posLeft, posTop, chartWidth, chartHeight)
                Dim chart = chartObject.Chart

                chart.SetSourceData(targetWs.Range(dataRange))
                chart.ChartType = xlChartType

                If Not String.IsNullOrWhiteSpace(chartTitle) Then
                    chart.HasTitle = True
                    chart.ChartTitle.Text = chartTitle
                End If

            Catch ex As Exception
                Debug.WriteLine($"Chart creation error: {ex.Message}")
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Applies print/page setup to a worksheet.
    ''' </summary>
    Private Shared Sub ApplyPrintSetup(ws As Microsoft.Office.Interop.Excel.Worksheet, setup As JObject)
        Try
            Dim orientation = setup.Value(Of String)("orientation")
            If Not String.IsNullOrWhiteSpace(orientation) Then
                Select Case orientation.ToLowerInvariant()
                    Case "landscape" : ws.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape
                    Case "portrait" : ws.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait
                End Select
            End If

            Dim fitWideToken = setup("fit_to_pages_wide")
            If fitWideToken IsNot Nothing Then
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesWide = CInt(fitWideToken)
            End If

            Dim fitTallToken = setup("fit_to_pages_tall")
            If fitTallToken IsNot Nothing Then
                ws.PageSetup.Zoom = False
                ws.PageSetup.FitToPagesTall = CInt(fitTallToken)
            End If

            Dim headerText = setup.Value(Of String)("header_text")
            If Not String.IsNullOrWhiteSpace(headerText) Then ws.PageSetup.CenterHeader = headerText

            Dim footerText = setup.Value(Of String)("footer_text")
            If Not String.IsNullOrWhiteSpace(footerText) Then ws.PageSetup.CenterFooter = footerText
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Injects VBA code modules into the workbook using late binding to avoid
    ''' a hard reference to Microsoft.Vbe.Interop.
    ''' Requires "Trust access to the VBA project object model" to be enabled in Excel Trust Center settings.
    ''' </summary>
    Private Shared Sub ApplyVbaModules(wb As Microsoft.Office.Interop.Excel.Workbook, modules As JArray)
        For Each modObj As JObject In modules
            Try
                Dim modName = If(modObj.Value(Of String)("name"), "Module1")
                Dim modCode = modObj.Value(Of String)("code")
                Dim modType = If(modObj.Value(Of String)("type"), "module").ToLowerInvariant()

                If String.IsNullOrWhiteSpace(modCode) Then Continue For

                ' Use CallByName to fully late-bind and avoid requiring Microsoft.Vbe.Interop reference.
                ' Even with Option Strict Off, wb.VBProject resolves via the typed Workbook interface
                ' which pulls in the Vbe.Interop assembly at compile time.
                Dim vbProj As Object = Microsoft.VisualBasic.Interaction.CallByName(wb, "VBProject", CallType.Get)
                Dim vbComponents As Object = Microsoft.VisualBasic.Interaction.CallByName(vbProj, "VBComponents", CallType.Get)

                If modType = "thisworkbook" Then
                    ' Insert code into the ThisWorkbook module
                    Dim tbComponent As Object = vbComponents("ThisWorkbook")
                    Dim codeMod As Object = Microsoft.VisualBasic.Interaction.CallByName(tbComponent, "CodeModule", CallType.Get)
                    Microsoft.VisualBasic.Interaction.CallByName(codeMod, "AddFromString", CallType.Method, modCode)
                Else
                    ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
                    Dim componentType As Integer = If(modType = "class", 2, 1)
                    Dim newMod As Object = Microsoft.VisualBasic.Interaction.CallByName(vbComponents, "Add", CallType.Method, componentType)
                    Microsoft.VisualBasic.Interaction.CallByName(newMod, "Name", CallType.Let, modName)
                    Dim codeMod As Object = Microsoft.VisualBasic.Interaction.CallByName(newMod, "CodeModule", CallType.Get)
                    Microsoft.VisualBasic.Interaction.CallByName(codeMod, "AddFromString", CallType.Method, modCode)
                End If
            Catch ex As Exception
                Debug.WriteLine($"VBA module insertion error: {ex.Message}")
            End Try
        Next
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCreateWordDocTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim markdownContent = GetArgString(toolCall.Arguments, "markdown_content")
            If String.IsNullOrWhiteSpace(markdownContent) Then
                response.Success = False
                response.Response = "Missing required parameter: markdown_content"
                Return response
            End If

            Dim fileName = GetArgString(toolCall.Arguments, "file_name")
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "Document"

            ' Sanitize filename
            For Each c In Path.GetInvalidFileNameChars()
                fileName = fileName.Replace(c, "_"c)
            Next
            If Not fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) Then
                fileName &= ".docx"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, fileName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                Dim baseName = Path.GetFileNameWithoutExtension(fileName)
                fileName = baseName & $"_{counter}.docx"
                outputPath = Path.Combine(_apCurrentTempDir, fileName)
                counter += 1
            End While

            context.Log($"Creating Word document: {fileName}")
            ApDashboardLog($"📝 Creating Word document: {fileName}", "step")

            Dim success = Await SwitchToUi(Function()
                                               Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                               Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                               Dim weCreated As Boolean = False
                                               Try
                                                   Try
                                                       wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                                   Catch
                                                       wordApp = New Microsoft.Office.Interop.Word.Application()
                                                       wordApp.Visible = False
                                                       weCreated = True
                                                   End Try
                                                   wordApp.ScreenUpdating = False

                                                   doc = wordApp.Documents.Add()
                                                   doc.Activate()

                                                   ' Use the existing shared library method to insert formatted Markdown
                                                   Dim sel As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
                                                   SharedMethods.InsertTextWithMarkdown(sel, markdownContent, TrailingCR:=False)

                                                   doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                                                   Return True
                                               Catch ex As Exception
                                                   Debug.WriteLine($"CreateWordDoc error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   Try : If doc IsNot Nothing Then doc.Close(False)
                                                   Catch : End Try
                                                   Try : If wordApp IsNot Nothing Then wordApp.ScreenUpdating = True
                                                   Catch : End Try
                                                   If weCreated AndAlso wordApp IsNot Nothing Then
                                                       Try : wordApp.Quit(False) : Catch : End Try
                                                   End If
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                ' Register as output on the first attachment, or create a standalone entry
                If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                    _apCurrentAttachments(0).OutputFiles.Add(outputPath)
                End If

                response.Success = True
                response.Response = $"Word document created: {fileName} ({New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply."
                ApDashboardLog($"✓ Word document created: {fileName}", "info")
            Else
                response.Success = False
                response.Response = "Failed to create Word document."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating Word document: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: comment_word_document
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteCommentWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim author = GetArgString(toolCall.Arguments, "author")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

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
            Dim authorNote = If(effectiveAuthor.Equals(AN6, StringComparison.OrdinalIgnoreCase), "", $" (author: {effectiveAuthor})")
            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Adding comments to: {att.OriginalFileName} with instruction: {instruction}{authorNote}")
                ApDashboardLog($"💬 Adding comments to: {att.OriginalFileName}{authorNote}", "step")

                If Not att.TempFilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) Then
                    resultMessages.Add($"✗ {att.OriginalFileName}: Only .docx files are supported for comment insertion.")
                    Continue For
                End If

                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_commented.docx"
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                Dim success = Await CommentDocxForAutoPilot(att.TempFilePath, outputPath, instruction, ct, author)

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

    Private Async Function ExecuteCompareWordDocsTool(
                toolCall As ToolCall,
                context As ToolExecutionContext,
                ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName,
            .Timestamp = DateTime.UtcNow, .OriginalCallJson = toolCall.RawJson
        }

        Try
            Dim originalFilename = GetArgString(toolCall.Arguments, "original_filename")
            Dim revisedFilename = GetArgString(toolCall.Arguments, "revised_filename")

            If String.IsNullOrWhiteSpace(originalFilename) OrElse String.IsNullOrWhiteSpace(revisedFilename) Then
                response.Success = False
                response.ErrorMessage = "Both 'original_filename' and 'revised_filename' are required."
                response.Response = response.ErrorMessage
                Return response
            End If

            ' Guard: need at least some attachments or output files to compare
            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No attachments available for comparison."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim originalAtt = FindAttachment(originalFilename)
            Dim revisedAtt = FindAttachment(revisedFilename)

            ' Use GetAllAvailableFileNames for better error messages
            If originalAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Original attachment '{originalFilename}' not found. Available: {String.Join(", ", GetAllAvailableFileNames())}"
                response.Response = response.ErrorMessage
                Return response
            End If

            If revisedAtt Is Nothing Then
                response.Success = False
                response.ErrorMessage = $"Revised attachment '{revisedFilename}' not found. Available: {String.Join(", ", GetAllAvailableFileNames())}"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim origExt = Path.GetExtension(originalAtt.TempFilePath).ToLowerInvariant()
            Dim revExt = Path.GetExtension(revisedAtt.TempFilePath).ToLowerInvariant()
            Dim supportedExts = {".doc", ".docx"}

            If Not supportedExts.Contains(origExt) OrElse Not supportedExts.Contains(revExt) Then
                response.Success = False
                response.ErrorMessage = "Both documents must be Word files (.doc or .docx)."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim compareName = $"Comparison_{Path.GetFileNameWithoutExtension(originalFilename)}_vs_{Path.GetFileNameWithoutExtension(revisedFilename)}.docx"
            Dim comparePath = Path.Combine(_apCurrentTempDir, compareName)

            context.Log($"Comparing: {originalFilename} (original) vs {revisedFilename} (revised)")
            ApDashboardLog($"📊 Comparing: {originalFilename} vs {revisedFilename}", "step")

            Dim success As Boolean = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(
                originalAtt.TempFilePath, revisedAtt.TempFilePath, comparePath))

            If success AndAlso File.Exists(comparePath) Then
                ' Register on a real attachment if possible; for transient objects the
                ' fallback directory scan in CollectResultAttachments will pick it up.
                Dim registrationTarget = _apCurrentAttachments.FirstOrDefault(
                    Function(a) a.OriginalFileName.Equals(originalFilename, StringComparison.OrdinalIgnoreCase))
                If registrationTarget IsNot Nothing Then
                    registrationTarget.OutputFiles.Add(comparePath)
                Else
                    ' Fallback: register on the first original attachment
                    _apCurrentAttachments(0).OutputFiles.Add(comparePath)
                End If

                Dim summaryText As String = ""
                Try
                    summaryText = Await SwitchToUi(Function()
                                                       Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                                       Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                                       Dim weCreated As Boolean = False
                                                       Try
                                                           Try
                                                               wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                                           Catch
                                                               wordApp = New Microsoft.Office.Interop.Word.Application() With {.Visible = False}
                                                               weCreated = True
                                                           End Try
                                                           doc = wordApp.Documents.Open(comparePath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                                                           Dim revCount = doc.Revisions.Count
                                                           Dim sb As New StringBuilder()
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
                                                           Catch : End Try
                                                           If weCreated AndAlso wordApp IsNot Nothing Then
                                                               Try : wordApp.Quit(False) : Catch : End Try
                                                           End If
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

    Private Async Function ExecuteProcessWordDocTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.ErrorMessage = "Missing required parameter: instruction"
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim sheetNames = GetArgStringArray(toolCall.Arguments, "sheet_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                ' Resolve each requested name via FindAttachment (supports output files)
                toProcess = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim att = FindAttachment(name)
                    If att IsNot Nothing AndAlso Not att.IsOverSizeLimit AndAlso att.TempFilePath IsNot Nothing Then
                        toProcess.Add(att)
                    End If
                Next
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) (a.Extension = ".docx" OrElse a.Extension = ".doc" OrElse
                                 a.Extension = ".pptx" OrElse a.Extension = ".xlsx") AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing
                ).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable Word, PowerPoint, or Excel attachments found."
                Return response
            End If

            ' Guard against recursive re-processing: warn if all targets are tool outputs
            Dim allAreOutputs = toProcess.All(Function(a) a.IsToolOutput)
            If allAreOutputs Then
                ApDashboardLog($"⚠ process_word_document called on tool output file(s) — proceeding with caution", "warn")
            End If

            Dim resultMessages As New List(Of String)()

            For Each att In toProcess
                context.Log($"Processing: {att.OriginalFileName} with instruction: {instruction}")

                Dim inputPath = att.TempFilePath
                Dim ext = att.Extension.ToLowerInvariant()
                Dim isPptx As Boolean = ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
                Dim isXlsx As Boolean = ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                Dim outputExt As String = If(isPptx, ".pptx", If(isXlsx, ".xlsx", ".docx"))
                Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_processed" & outputExt
                Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

                ' Prevent filename collision when re-processing
                Dim counter = 1
                While File.Exists(outputPath)
                    outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_processed_{counter}" & outputExt
                    outputPath = Path.Combine(_apCurrentTempDir, outputName)
                    counter += 1
                End While

                ' Pass sheet filter only for Excel files
                Dim sheetFilter As List(Of String) = If(isXlsx AndAlso sheetNames.Count > 0, sheetNames, Nothing)
                Dim success = Await ProcessDocumentForAutoPilot(inputPath, outputPath, instruction, ct, sheetFilter)

                If success Then
                    ' Register output on the original attachment (not on a transient object)
                    Dim registrationTarget = If(att.IsToolOutput,
                        _apCurrentAttachments.FirstOrDefault(Function(a) a.OutputFiles IsNot Nothing AndAlso
                            a.OutputFiles.Any(Function(p) Path.GetFileName(p).Equals(att.OriginalFileName, StringComparison.OrdinalIgnoreCase))),
                        att)
                    If registrationTarget Is Nothing Then registrationTarget = _apCurrentAttachments(0)

                    registrationTarget.OutputFiles.Add(outputPath)

                    ' Compare document only for Word files (not PPTX or XLSX)
                    If Not isPptx AndAlso Not isXlsx Then
                        Dim comparePath = Path.Combine(_apCurrentTempDir,
                            Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_compare.docx")
                        ' Prevent compare filename collision too
                        Dim cmpCounter = 1
                        While File.Exists(comparePath)
                            comparePath = Path.Combine(_apCurrentTempDir,
                                Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_compare_{cmpCounter}.docx")
                            cmpCounter += 1
                        End While

                        Dim compareSuccess = Await SwitchToUi(Function() CreateWordCompareDocumentForAutoPilot(inputPath, outputPath, comparePath))
                        If compareSuccess Then
                            registrationTarget.OutputFiles.Add(comparePath)
                            resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} + compare document.")
                        Else
                            resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName} (compare document creation failed).")
                        End If
                    Else
                        resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Output: {outputName}")
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
            response.Response = $"Error processing document(s): {ex.Message}"
        End Try

        Return response
    End Function


    ''' <summary>
    ''' Creates a Word tracked-changes comparison document from an original and revised file.
    ''' </summary>
    ''' <param name="originalPath">Path to the baseline/original Word file.</param>
    ''' <param name="processedPath">Path to the revised/processed Word file.</param>
    ''' <param name="comparePath">Destination path for the generated comparison document.</param>
    ''' <returns><c>True</c> if comparison output is created successfully; otherwise <c>False</c>.</returns>
    Private Function CreateWordCompareDocumentForAutoPilot(originalPath As String, processedPath As String, comparePath As String) As Boolean
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
        Dim originalDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim processedDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim weCreatedWordApp As Boolean = False

        Try
            Try
                wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
            Catch
                wordApp = New Microsoft.Office.Interop.Word.Application()
                wordApp.Visible = False
                weCreatedWordApp = True
            End Try

            Dim wasScreenUpdating = wordApp.ScreenUpdating
            wordApp.ScreenUpdating = False

            originalDoc = wordApp.Documents.Open(originalPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            processedDoc = wordApp.Documents.Open(processedPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc, RevisedDocument:=processedDoc,
                Destination:=Microsoft.Office.Interop.Word.WdCompareDestination.wdCompareDestinationNew,
                Granularity:=Microsoft.Office.Interop.Word.WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:=True,
                CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True,
                CompareTextboxes:=True, CompareFields:=True, CompareComments:=True,
                RevisedAuthor:=AN6, IgnoreAllComparisonWarnings:=True)

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
            If weCreatedWordApp AndAlso wordApp IsNot Nothing Then Try : wordApp.Quit(False) : Catch : End Try
        End Try
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_pdf_text
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteExtractPdfTextTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)) AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            Else
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No PDF attachments found to extract."
                Return response
            End If

            Dim sb As New StringBuilder()
            For Each att In toProcess
                context.Log($"Extracting text from: {att.OriginalFileName}")
                ApDashboardLog($"📄 Extracting text from: {att.OriginalFileName}", "step")

                Dim pdfResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                    att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=False, AskUser:=False, context:=_context)

                Dim text As String = If(pdfResult IsNot Nothing, pdfResult.Content, "")
                Dim usedOcr As Boolean = False

                Dim ocrAvailable As Boolean = SharedMethods.IsOcrAvailable(_context)
                Dim needsOcr As Boolean = String.IsNullOrWhiteSpace(text) OrElse
                    (pdfResult IsNot Nothing AndAlso pdfResult.OcrWasSkippedDueToHeuristics)

                If needsOcr AndAlso ocrAvailable Then
                    ApDashboardLog($"🔍 Running OCR on: {att.OriginalFileName}", "step")
                    context.Log($"OCR: {att.OriginalFileName}")
                    Dim ocrResult As PdfReadResult = Await SharedMethods.ReadPdfAsTextEx(
                        att.TempFilePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=True, AskUser:=False, context:=_context)
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
                    sb.AppendLine(If(Not ocrAvailable,
                        "(no extractable text; OCR is not available in the current configuration)",
                        "(no extractable text)"))
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

    Private Async Function ExecuteDescribeBinaryTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim prompt = If(GetArgString(toolCall.Arguments, "prompt"), "Describe or transcribe this file.")

            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response
            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False : response.Response = $"Attachment '{fileName}' could not be read." : Return response
            End If

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

            Dim useSecond As Boolean = (_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi)
            Dim llmResult As String = Await SharedMethods.LLM(
                _context, prompt, "", UseSecondAPI:=useSecond, Hidesplash:=True,
                FileObject:=att.TempFilePath, cancellationToken:=ct)

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

    Private Async Function ExecuteMergePdfsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "merged.pdf")

            Dim toMerge As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toMerge = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim found = _apCurrentAttachments?.FirstOrDefault(
                        Function(a) a.OriginalFileName.Equals(name, StringComparison.OrdinalIgnoreCase) AndAlso
                                    Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing)
                    If found IsNot Nothing Then toMerge.Add(found)
                Next
            Else
                toMerge = _apCurrentAttachments?.Where(
                    Function(a) a.Extension = ".pdf" AndAlso Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            End If

            If toMerge Is Nothing OrElse toMerge.Count < 2 Then
                response.Success = False
                response.Response = "Need at least 2 PDF attachments to merge."
                Return response
            End If

            context.Log($"Merging {toMerge.Count} PDFs into {outputName}")
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

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

    Private Async Function ExecuteReadAttachmentTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            ' Support both single and batch mode
            Dim fileNames As New List(Of String)()
            Dim singleName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim batchNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            If batchNames.Count > 0 Then
                fileNames.AddRange(batchNames)
            ElseIf Not String.IsNullOrWhiteSpace(singleName) Then
                fileNames.Add(singleName)
            End If

            If fileNames.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name or attachment_names"
                Return response
            End If

            Dim sb As New StringBuilder()
            Dim anySuccess As Boolean = False

            For Each fileName In fileNames
                Dim att = FindAttachment(fileName)

                If att Is Nothing Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' not found.")
                    sb.AppendLine()
                    Continue For
                End If

                If att.IsOverSizeLimit Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' exceeds the size limit and cannot be processed.")
                    sb.AppendLine()
                    Continue For
                End If

                If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Attachment '{fileName}' could not be read.")
                    sb.AppendLine()
                    Continue For
                End If

                context.Log($"Reading attachment: {fileName}")
                Dim text = Await ReadSingleAttachmentText(att, context)

                If Not String.IsNullOrWhiteSpace(text) Then
                    If text.Length > 50000 Then
                        text = text.Substring(0, 50000) & vbCrLf & "[... content truncated at 50,000 characters ...]"
                    End If

                    If fileNames.Count > 1 Then sb.AppendLine($"[{fileName}]")
                    sb.AppendLine(text)

                    ' Append docx metadata hint if applicable
                    Dim hint = GetDocxMetadataHint(att)
                    If Not String.IsNullOrWhiteSpace(hint) Then sb.AppendLine(hint)

                    sb.AppendLine()
                    anySuccess = True
                Else
                    sb.AppendLine($"[{fileName}]")
                    sb.AppendLine($"Could not extract text from '{fileName}'. The file format may not be supported.")
                    sb.AppendLine()
                End If
            Next

            response.Success = anySuccess
            response.Response = sb.ToString().TrimEnd()
            If Not anySuccess Then
                response.Response = If(fileNames.Count = 1,
                    $"Could not extract text from '{fileNames(0)}'. The file format may not be supported.",
                    response.Response)
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

    Private Function ExecuteListAttachmentsTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            If _apCurrentAttachments Is Nothing OrElse _apCurrentAttachments.Count = 0 Then
                response.Success = True
                response.Response = "No attachments in this email."
                Return response
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine($"Attachments ({_apCurrentAttachments.Count}):")
            For i As Integer = 0 To _apCurrentAttachments.Count - 1
                Dim att = _apCurrentAttachments(i)
                Dim sizeStr = If(att.SizeBytes > 0, $"{att.SizeBytes / 1024:F0} KB", "unknown size")
                Dim statusStr = If(att.IsOverSizeLimit, " [OVER SIZE LIMIT]",
                               If(att.TempFilePath IsNot Nothing, " [available for processing]", " [not available]"))
                Dim pdfStr As String = ""
                If att.PageCount > 0 Then
                    pdfStr = $", {att.PageCount} page(s)"
                    If Not String.IsNullOrWhiteSpace(att.PageOrientation) Then
                        pdfStr &= $", {att.PageOrientation}"
                    End If
                    If Not String.IsNullOrWhiteSpace(att.PageSize) Then
                        pdfStr &= $", {att.PageSize}"
                    End If
                End If
                sb.AppendLine($"  {i + 1}. {att.OriginalFileName} ({att.Extension}, {sizeStr}{pdfStr}){statusStr}")
            Next

            ' List output files produced by earlier tool calls
            Dim outputFileCount = 0
            For Each att In _apCurrentAttachments
                If att.OutputFiles IsNot Nothing Then
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            outputFileCount += 1
                        End If
                    Next
                End If
            Next

            If outputFileCount > 0 Then
                sb.AppendLine()
                sb.AppendLine($"Tool output files ({outputFileCount}):")
                Dim outIdx = 1
                For Each att In _apCurrentAttachments
                    If att.OutputFiles Is Nothing Then Continue For
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            Dim outName = Path.GetFileName(outputPath)
                            Dim outExt = Path.GetExtension(outputPath).ToLowerInvariant()
                            Dim outSize = New FileInfo(outputPath).Length
                            sb.AppendLine($"  {outIdx}. {outName} ({outExt}, {outSize / 1024:F0} KB) [tool output — available for processing]")
                            outIdx += 1
                        End If
                    Next
                Next
            End If

            response.Success = True
            response.Response = sb.ToString().TrimEnd()

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error listing attachments: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: read_word_document_details (OpenXML deep reader)
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteReadWordDocDetailsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response
            If att.TempFilePath Is Nothing OrElse Not File.Exists(att.TempFilePath) Then
                response.Success = False : response.Response = $"Attachment '{fileName}' could not be read." : Return response
            End If
            If att.Extension <> ".docx" Then
                response.Success = False : response.Response = $"Only .docx files are supported. '{fileName}' is {att.Extension}." : Return response
            End If

            Dim includeComments = GetArgBool(toolCall.Arguments, "include_comments", True)
            Dim includeHeadersFooters = GetArgBool(toolCall.Arguments, "include_headers_footers", False)
            Dim includeFootnotesEndnotes = GetArgBool(toolCall.Arguments, "include_footnotes_endnotes", False)
            Dim includeTrackedChanges = GetArgBool(toolCall.Arguments, "include_tracked_changes", True)
            Dim filterAuthor = GetArgString(toolCall.Arguments, "tracked_changes_author")
            Dim filterSinceStr = GetArgString(toolCall.Arguments, "tracked_changes_since")

            Dim filterSince As DateTime? = Nothing
            If Not String.IsNullOrWhiteSpace(filterSinceStr) Then
                Dim parsed As DateTime
                If DateTime.TryParse(filterSinceStr, Globalization.CultureInfo.InvariantCulture,
                                     Globalization.DateTimeStyles.None, parsed) Then
                    filterSince = parsed
                End If
            End If

            context.Log($"Deep-reading Word document: {fileName}")
            ApDashboardLog($"📖 Deep-reading: {fileName}", "step")

            Dim result = Await Task.Run(Function() ExtractWordDocumentDetails(
                att.TempFilePath, includeComments, includeHeadersFooters,
                includeFootnotesEndnotes, includeTrackedChanges, filterAuthor, filterSince))

            If result.Length > 100000 Then
                result = result.Substring(0, 100000) & vbCrLf & "[... content truncated at 100,000 characters ...]"
            End If

            response.Success = True
            response.Response = result
            ApDashboardLog($"✓ Deep-read complete: {fileName} ({result.Length:N0} chars)", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error reading Word document details: {ex.Message}"
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Extracts detailed content from a .docx file using OpenXML, including body text
    ''' with inline tracked change markers, comments, headers/footers, and footnotes/endnotes.
    ''' </summary>
    Private Function ExtractWordDocumentDetails(
            filePath As String,
            includeComments As Boolean,
            includeHeadersFooters As Boolean,
            includeFootnotesEndnotes As Boolean,
            includeTrackedChanges As Boolean,
            filterAuthor As String,
            filterSince As DateTime?) As String

        Dim tempDir = Path.Combine(Path.GetTempPath(), "ap_detail_" & Guid.NewGuid().ToString("N"))
        Try
            ZipFile.ExtractToDirectory(filePath, tempDir)

            Dim nsMgr As XmlNamespaceManager = Nothing
            Dim docXml As XmlDocument = Nothing
            Dim docPath = Path.Combine(tempDir, "word", "document.xml")

            If File.Exists(docPath) Then
                docXml = New XmlDocument()
                docXml.Load(docPath)
                nsMgr = New XmlNamespaceManager(docXml.NameTable)
                nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                nsMgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
            End If

            Dim sb As New StringBuilder()

            ' ── BODY TEXT (with optional inline tracked changes) ──
            If docXml IsNot Nothing Then
                Dim bodyNode = docXml.SelectSingleNode("//w:body", nsMgr)
                If bodyNode IsNot Nothing Then
                    Dim headerLabel = If(includeTrackedChanges, "═══ DOCUMENT BODY (with tracked changes) ═══", "═══ DOCUMENT BODY ═══")
                    sb.AppendLine(headerLabel)
                    sb.AppendLine()

                    Dim revInsCount = 0
                    Dim revDelCount = 0
                    Dim revFmtCount = 0
                    Dim authorCounts As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                    For Each paraNode As XmlNode In bodyNode.SelectNodes("w:p", nsMgr)
                        Dim paraText As New StringBuilder()

                        For Each child As XmlNode In paraNode.ChildNodes
                            ProcessDocBodyNode(child, nsMgr, paraText, includeTrackedChanges,
                                             filterAuthor, filterSince,
                                             revInsCount, revDelCount, revFmtCount, authorCounts)
                        Next

                        Dim line = paraText.ToString()
                        If Not String.IsNullOrWhiteSpace(line) Then sb.AppendLine(line)
                        sb.AppendLine()
                    Next

                    ' Summary
                    If includeTrackedChanges Then
                        Dim total = revInsCount + revDelCount + revFmtCount
                        sb.AppendLine($"═══ TRACKED CHANGES SUMMARY ═══")
                        sb.AppendLine($"Total: {total} revision(s) (Insertions: {revInsCount} | Deletions: {revDelCount} | Format changes: {revFmtCount})")
                        If authorCounts.Count > 0 Then
                            sb.AppendLine("By author: " & String.Join(", ", authorCounts.Select(Function(kv) $"{kv.Key}: {kv.Value}")))
                        End If
                        sb.AppendLine()
                    End If
                End If
            End If

            ' ── COMMENTS ──
            If includeComments Then
                Dim commentsPath = Path.Combine(tempDir, "word", "comments.xml")
                If File.Exists(commentsPath) Then
                    Dim commDoc As New XmlDocument()
                    commDoc.Load(commentsPath)
                    Dim cNsMgr As New XmlNamespaceManager(commDoc.NameTable)
                    cNsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                    Dim commentNodes = commDoc.SelectNodes("//w:comment", cNsMgr)

                    ' Build comment-to-anchor mapping from document.xml
                    Dim commentAnchors As New Dictionary(Of String, String)()
                    If docXml IsNot Nothing Then
                        BuildCommentAnchorMap(docXml, nsMgr, commentAnchors)
                    End If

                    If commentNodes.Count > 0 Then
                        sb.AppendLine($"═══ COMMENTS ({commentNodes.Count}) ═══")
                        Dim idx = 1
                        For Each cNode As XmlElement In commentNodes
                            Dim author = cNode.GetAttribute("w:author")
                            Dim dateStr = cNode.GetAttribute("w:date")
                            Dim commentId = cNode.GetAttribute("w:id")
                            Dim commentText As New StringBuilder()
                            For Each tNode As XmlNode In cNode.SelectNodes(".//w:t", cNsMgr)
                                commentText.Append(tNode.InnerText)
                            Next

                            sb.AppendLine($"[Comment #{idx}] Author: {author} | Date: {dateStr}")
                            Dim anchorText As String = Nothing
                            If commentAnchors.TryGetValue(commentId, anchorText) AndAlso Not String.IsNullOrWhiteSpace(anchorText) Then
                                If anchorText.Length > 200 Then anchorText = anchorText.Substring(0, 200) & "..."
                                sb.AppendLine($"  Anchored to: ""{anchorText}""")
                            End If
                            sb.AppendLine($"  Comment: {commentText}")
                            sb.AppendLine()
                            idx += 1
                        Next
                    End If
                End If
            End If

            ' ── HEADERS & FOOTERS ──
            If includeHeadersFooters Then
                ExtractHeadersFooters(tempDir, sb, "header", "HEADERS")
                ExtractHeadersFooters(tempDir, sb, "footer", "FOOTERS")
            End If

            ' ── FOOTNOTES & ENDNOTES ──
            If includeFootnotesEndnotes Then
                ExtractNotesSection(tempDir, sb, "footnotes.xml", "FOOTNOTES")
                ExtractNotesSection(tempDir, sb, "endnotes.xml", "ENDNOTES")
            End If

            Return sb.ToString().TrimEnd()
        Finally
            Try : Directory.Delete(tempDir, True) : Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Recursively processes a node in the document body, emitting text and inline change markers.
    ''' </summary>
    Private Sub ProcessDocBodyNode(
            node As XmlNode, nsMgr As XmlNamespaceManager, sb As StringBuilder,
            includeTrackedChanges As Boolean, filterAuthor As String, filterSince As DateTime?,
            ByRef insCount As Integer, ByRef delCount As Integer, ByRef fmtCount As Integer,
            authorCounts As Dictionary(Of String, Integer))

        If node Is Nothing Then Return

        Select Case node.LocalName
            Case "r" ' Normal run
                For Each tNode As XmlNode In node.SelectNodes("w:t", nsMgr)
                    sb.Append(tNode.InnerText)
                Next

            Case "ins" ' Insertion
                Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                Dim shortDate = If(dateStr.Length >= 10, dateStr.Substring(0, 10), dateStr)

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthor, filterSince)

                If includeTrackedChanges AndAlso passesFilter Then
                    Dim innerText As New StringBuilder()
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:t", nsMgr)
                            innerText.Append(tNode.InnerText)
                        Next
                    Next
                    sb.Append($"«INS|{author}|{shortDate}»{innerText}«/INS»")
                    insCount += 1
                    IncrementAuthorCount(authorCounts, author)
                Else
                    ' When not showing changes or filtered out: show inserted text as accepted
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:t", nsMgr)
                            sb.Append(tNode.InnerText)
                        Next
                    Next
                End If

            Case "del" ' Deletion
                Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                Dim shortDate = If(dateStr.Length >= 10, dateStr.Substring(0, 10), dateStr)

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthor, filterSince)

                If includeTrackedChanges AndAlso passesFilter Then
                    Dim innerText As New StringBuilder()
                    For Each child As XmlNode In node.ChildNodes
                        For Each tNode As XmlNode In child.SelectNodes(".//w:delText | .//w:t", nsMgr)
                            innerText.Append(tNode.InnerText)
                        Next
                    Next
                    sb.Append($"«DEL|{author}|{shortDate}»{innerText}«/DEL»")
                    delCount += 1
                    IncrementAuthorCount(authorCounts, author)
                End If
                ' When not showing changes or filtered out: omit deleted text (it was deleted)

            Case "rPrChange" ' Format change
                If includeTrackedChanges Then
                    Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                    Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                    If PassesRevisionFilter(author, dateStr, filterAuthor, filterSince) Then
                        fmtCount += 1
                        IncrementAuthorCount(authorCounts, author)
                    End If
                End If

            Case Else
                ' Recurse into child nodes for structure elements like hyperlinks, smart tags, etc.
                For Each child As XmlNode In node.ChildNodes
                    ProcessDocBodyNode(child, nsMgr, sb, includeTrackedChanges,
                                     filterAuthor, filterSince, insCount, delCount, fmtCount, authorCounts)
                Next
        End Select
    End Sub

    Private Shared Function PassesRevisionFilter(author As String, dateStr As String,
                                                  filterAuthor As String, filterSince As DateTime?) As Boolean
        If Not String.IsNullOrWhiteSpace(filterAuthor) Then
            If Not author.IndexOf(filterAuthor, StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        End If
        If filterSince.HasValue AndAlso Not String.IsNullOrWhiteSpace(dateStr) Then
            Dim revDate As DateTime
            If DateTime.TryParse(dateStr, Globalization.CultureInfo.InvariantCulture,
                                 Globalization.DateTimeStyles.None, revDate) Then
                If revDate < filterSince.Value Then Return False
            End If
        End If
        Return True
    End Function

    Private Shared Sub IncrementAuthorCount(dict As Dictionary(Of String, Integer), author As String)
        If String.IsNullOrWhiteSpace(author) Then author = "(unknown)"
        If dict.ContainsKey(author) Then dict(author) += 1 Else dict(author) = 1
    End Sub

    ''' <summary>
    ''' Builds a mapping from comment ID to the text that the comment is anchored to.
    ''' </summary>
    Private Sub BuildCommentAnchorMap(docXml As XmlDocument, nsMgr As XmlNamespaceManager,
                                      anchors As Dictionary(Of String, String))
        ' Find all commentRangeStart / commentRangeEnd pairs
        Dim starts = docXml.SelectNodes("//w:commentRangeStart", nsMgr)
        For Each startNode As XmlElement In starts
            Dim commentId = startNode.GetAttribute("w:id")
            If String.IsNullOrEmpty(commentId) Then Continue For

            ' Collect text nodes between commentRangeStart and commentRangeEnd with same id
            Dim anchorText As New StringBuilder()
            Dim current = startNode.NextSibling
            Dim found = False
            Dim maxNodes = 500 ' Safety limit

            While current IsNot Nothing AndAlso maxNodes > 0
                maxNodes -= 1
                If current.LocalName = "commentRangeEnd" Then
                    Dim endId = DirectCast(current, XmlElement).GetAttribute("w:id")
                    If endId = commentId Then found = True : Exit While
                End If

                For Each tNode As XmlNode In current.SelectNodes(".//w:t", nsMgr)
                    anchorText.Append(tNode.InnerText)
                Next

                current = current.NextSibling
            End While

            ' If not found as sibling, might be across paragraphs — still use what we got
            If anchorText.Length > 0 Then anchors(commentId) = anchorText.ToString()
        Next
    End Sub

    ''' <summary>
    ''' Extracts header or footer content from the word directory.
    ''' </summary>
    Private Sub ExtractHeadersFooters(tempDir As String, sb As StringBuilder, prefix As String, label As String)
        Dim wordDir = Path.Combine(tempDir, "word")
        If Not Directory.Exists(wordDir) Then Return

        Dim files = Directory.GetFiles(wordDir, prefix & "*.xml")
        If files.Length = 0 Then Return

        Dim anyContent = False
        Dim tempSb As New StringBuilder()

        For Each f In files
            Try
                Dim doc As New XmlDocument()
                doc.Load(f)
                Dim ns As New XmlNamespaceManager(doc.NameTable)
                ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

                Dim text As New StringBuilder()
                For Each tNode As XmlNode In doc.SelectNodes("//w:t", ns)
                    text.Append(tNode.InnerText)
                Next

                If text.Length > 0 Then
                    Dim shortName = Path.GetFileNameWithoutExtension(f)
                    tempSb.AppendLine($"[{shortName}] {text}")
                    anyContent = True
                End If
            Catch
            End Try
        Next

        If anyContent Then
            sb.AppendLine($"═══ {label} ═══")
            sb.Append(tempSb)
            sb.AppendLine()
        End If
    End Sub

    ''' <summary>
    ''' Extracts footnotes or endnotes from the corresponding XML file.
    ''' </summary>
    Private Sub ExtractNotesSection(tempDir As String, sb As StringBuilder, xmlFileName As String, label As String)
        Dim notesPath = Path.Combine(tempDir, "word", xmlFileName)
        If Not File.Exists(notesPath) Then Return

        Try
            Dim doc As New XmlDocument()
            doc.Load(notesPath)
            Dim ns As New XmlNamespaceManager(doc.NameTable)
            ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

            ' Footnotes/endnotes have w:footnote or w:endnote elements; skip type="separator"/"continuationSeparator"
            Dim nodeName = If(xmlFileName.Contains("footnote"), "w:footnote", "w:endnote")
            Dim noteNodes = doc.SelectNodes($"//{nodeName}", ns)

            Dim entries As New List(Of String)()
            For Each noteNode As XmlElement In noteNodes
                Dim noteType = noteNode.GetAttribute("w:type")
                If noteType = "separator" OrElse noteType = "continuationSeparator" Then Continue For

                Dim noteId = noteNode.GetAttribute("w:id")
                Dim noteText As New StringBuilder()
                For Each tNode As XmlNode In noteNode.SelectNodes(".//w:t", ns)
                    noteText.Append(tNode.InnerText)
                Next

                If noteText.Length > 0 Then
                    entries.Add($"[{label.TrimEnd("S"c)} {noteId}] {noteText}")
                End If
            Next

            If entries.Count > 0 Then
                sb.AppendLine($"═══ {label} ({entries.Count}) ═══")
                For Each entry In entries
                    sb.AppendLine(entry)
                Next
                sb.AppendLine()
            End If
        Catch
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: create_pdf_from_text
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteCreatePdfFromTextTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim content = GetArgString(toolCall.Arguments, "content")
            If String.IsNullOrWhiteSpace(content) Then
                response.Success = False
                response.Response = "Missing required parameter: content"
                Return response
            End If

            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "output.pdf")
            Dim title = GetArgString(toolCall.Arguments, "title")
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Creating PDF: {outputName}")
            ApDashboardLog($"📝 Creating PDF: {outputName}", "step")

            ' Ensure font resolver is configured before any XFont usage
            EnsureApPdfSharpFontResolver()

            Using doc As New PdfSharp.Pdf.PdfDocument()
                doc.Info.Title = If(title, "Generated Document")

                Dim font = New PdfSharp.Drawing.XFont("Arial", 11)
                Dim titleFont = New PdfSharp.Drawing.XFont("Arial", 16, PdfSharp.Drawing.XFontStyleEx.Bold)
                Dim margin = 50.0
                Dim pageWidth = 595.0 ' A4
                Dim pageHeight = 842.0
                Dim usableWidth = pageWidth - 2 * margin
                Dim lineHeight = 15.0
                Dim y = margin

                Dim page = doc.AddPage()
                page.Width = pageWidth
                page.Height = pageHeight
                Dim gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page)

                ' Title
                If Not String.IsNullOrWhiteSpace(title) Then
                    gfx.DrawString(title, titleFont, PdfSharp.Drawing.XBrushes.Black,
                                   New PdfSharp.Drawing.XRect(margin, y, usableWidth, 30),
                                   PdfSharp.Drawing.XStringFormats.TopLeft)
                    y += 35
                End If

                ' Content lines
                Dim lines = content.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                For Each line In lines
                    If y + lineHeight > pageHeight - margin Then
                        page = doc.AddPage()
                        page.Width = pageWidth
                        page.Height = pageHeight
                        gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page)
                        y = margin
                    End If

                    If Not String.IsNullOrEmpty(line) Then
                        gfx.DrawString(line, font, PdfSharp.Drawing.XBrushes.Black,
                                       New PdfSharp.Drawing.XRect(margin, y, usableWidth, lineHeight),
                                       PdfSharp.Drawing.XStringFormats.TopLeft)
                    End If
                    y += lineHeight
                Next

                doc.Save(outputPath)
            End Using

            ' Register as output
            If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Count > 0 Then
                _apCurrentAttachments(0).OutputFiles.Add(outputPath)
            End If

            response.Success = True
            response.Response = $"PDF created: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)"
            ApDashboardLog($"✓ PDF created: {outputName}", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error creating PDF: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_excel_data
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteExtractExcelDataTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response

            Dim sheetFilter = GetArgString(toolCall.Arguments, "sheet_name")

            context.Log($"Extracting Excel data: {fileName}")
            ApDashboardLog($"📊 Extracting Excel data: {fileName}", "step")

            ' Use the existing ExtractExcelText which handles interop
            Dim text = ExtractExcelText(att.TempFilePath)

            If String.IsNullOrWhiteSpace(text) OrElse text.StartsWith("Error") Then
                response.Success = False
                response.Response = $"Could not extract data from '{fileName}'."
                Return response
            End If

            ' Filter by sheet name if specified
            If Not String.IsNullOrWhiteSpace(sheetFilter) Then
                Dim sheetMarker = $"[Sheet: {sheetFilter}]"
                Dim idx = text.IndexOf(sheetMarker, StringComparison.OrdinalIgnoreCase)
                If idx >= 0 Then
                    ' Find the next sheet marker or end
                    Dim nextSheet = text.IndexOf("[Sheet: ", idx + sheetMarker.Length, StringComparison.OrdinalIgnoreCase)
                    text = If(nextSheet >= 0, text.Substring(idx, nextSheet - idx).TrimEnd(), text.Substring(idx).TrimEnd())
                End If
            End If

            If text.Length > 50000 Then
                text = text.Substring(0, 50000) & vbCrLf & "[... content truncated at 50,000 characters ...]"
            End If

            response.Success = True
            response.Response = text
            ApDashboardLog($"✓ Excel data extracted: {fileName} ({text.Length:N0} chars)", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting Excel data: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: split_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteSplitPdfTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim startPage = GetArgInt(toolCall.Arguments, "start_page", 0)
            Dim endPage = GetArgInt(toolCall.Arguments, "end_page", 0)
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "split.pdf")

            If String.IsNullOrWhiteSpace(fileName) OrElse startPage < 1 OrElse endPage < 1 Then
                response.Success = False
                response.Response = "Missing required parameters: attachment_name, start_page, end_page"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Splitting PDF: {fileName} pages {startPage}-{endPage}")

            Using inputDoc = PdfSharp.Pdf.IO.PdfReader.Open(att.TempFilePath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import)
                If endPage > inputDoc.PageCount Then endPage = inputDoc.PageCount
                If startPage > inputDoc.PageCount Then
                    response.Success = False
                    response.Response = $"Start page {startPage} exceeds document page count ({inputDoc.PageCount})."
                    Return response
                End If

                Using outputDoc As New PdfSharp.Pdf.PdfDocument()
                    For i As Integer = startPage - 1 To endPage - 1
                        outputDoc.AddPage(inputDoc.Pages(i))
                    Next
                    outputDoc.Save(outputPath)
                End Using
            End Using

            att.OutputFiles.Add(outputPath)
            response.Success = True
            response.Response = $"Split PDF: pages {startPage}-{endPage} extracted to {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error splitting PDF: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: add_pdf_watermark
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteAddPdfWatermarkTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            Dim watermarkText = GetArgString(toolCall.Arguments, "watermark_text")
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), "watermarked.pdf")

            If String.IsNullOrWhiteSpace(fileName) OrElse String.IsNullOrWhiteSpace(watermarkText) Then
                response.Success = False
                response.Response = "Missing required parameters: attachment_name, watermark_text"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Adding watermark to: {fileName}")
            ApDashboardLog($"💧 Adding watermark to: {fileName}", "step")

            ' Ensure font resolver is configured before any XFont usage
            EnsureApPdfSharpFontResolver()

            ' Write to a temp file first, then move to final path to avoid lock conflicts
            Dim tempOutputPath = outputPath & ".tmp_" & Guid.NewGuid().ToString("N") & ".pdf"

            Try
                ' Copy source to temp output
                File.Copy(att.TempFilePath, tempOutputPath, True)
                Using doc = PdfSharp.Pdf.IO.PdfReader.Open(tempOutputPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                    Dim wmFont = New PdfSharp.Drawing.XFont("Arial", 60, PdfSharp.Drawing.XFontStyleEx.Bold)
                    Dim wmBrush = New PdfSharp.Drawing.XSolidBrush(
                        PdfSharp.Drawing.XColor.FromArgb(80, 180, 180, 180))

                    For Each page In doc.Pages
                        Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                            Dim state = gfx.Save()

                            ' Move origin to center of page
                            gfx.TranslateTransform(page.Width.Point / 2, page.Height.Point / 2)
                            gfx.RotateTransform(-45)

                            ' Measure and draw the watermark text centered
                            Dim size = gfx.MeasureString(watermarkText, wmFont)
                            gfx.DrawString(watermarkText, wmFont, wmBrush,
                                           New PdfSharp.Drawing.XRect(-size.Width / 2, -size.Height / 2, size.Width, size.Height),
                                           PdfSharp.Drawing.XStringFormats.Center)

                            gfx.Restore(state)
                        End Using
                    Next
                    doc.Save(tempOutputPath)
                End Using

                ' All handles released — safe to move
                If File.Exists(outputPath) Then File.Delete(outputPath)
                File.Move(tempOutputPath, outputPath)
            Finally
                ' Clean up temp file on any failure
                Try : If File.Exists(tempOutputPath) Then File.Delete(tempOutputPath)
                Catch : End Try
            End Try

            att.OutputFiles.Add(outputPath)
            response.Success = True
            response.Response = $"Watermark '{watermarkText}' added to {outputName}."
            ApDashboardLog($"✓ Watermark added: {outputName}", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error adding watermark: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: word_to_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteWordToPdfTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response

            Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If ext <> ".doc" AndAlso ext <> ".docx" Then
                response.Success = False
                response.Response = $"'{fileName}' is not a Word document ({ext})."
                Return response
            End If

            Dim outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & ".pdf"
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Converting to PDF: {fileName}")
            ApDashboardLog($"📄 Converting to PDF: {fileName}", "step")

            Dim success = Await SwitchToUi(Function()
                                               Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                               Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                               Dim weCreated As Boolean = False
                                               Try
                                                   Try
                                                       wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                                   Catch
                                                       wordApp = New Microsoft.Office.Interop.Word.Application()
                                                       wordApp.Visible = False
                                                       weCreated = True
                                                   End Try
                                                   wordApp.ScreenUpdating = False
                                                   doc = wordApp.Documents.Open(att.TempFilePath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
                                                   doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)
                                                   Return True
                                               Catch ex As Exception
                                                   Debug.WriteLine($"WordToPdf error: {ex.Message}")
                                                   Return False
                                               Finally
                                                   Try : If doc IsNot Nothing Then doc.Close(False)
                                                   Catch : End Try
                                                   Try : If wordApp IsNot Nothing Then wordApp.ScreenUpdating = True
                                                   Catch : End Try
                                                   If weCreated AndAlso wordApp IsNot Nothing Then
                                                       Try : wordApp.Quit(False) : Catch : End Try
                                                   End If
                                               End Try
                                           End Function)

            If success AndAlso File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"Converted '{fileName}' to PDF: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB)."
                ApDashboardLog($"✓ Converted to PDF: {outputName}", "info")
            Else
                response.Success = False
                response.Response = $"Failed to convert '{fileName}' to PDF."
            End If

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error converting to PDF: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: pdf_to_word
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecutePdfToWordTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then response.Success = False : response.Response = $"Attachment '{fileName}' not found." : Return response
            If att.IsOverSizeLimit Then response.Success = False : response.Response = $"Attachment '{fileName}' exceeds the size limit." : Return response

            Dim ext = Path.GetExtension(att.TempFilePath).ToLowerInvariant()
            If ext <> ".pdf" Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF ({ext})."
                Return response
            End If

            Dim defaultOutput = Path.GetFileNameWithoutExtension(att.OriginalFileName) & ".docx"
            Dim outputName = If(GetArgString(toolCall.Arguments, "output_filename"), defaultOutput)
            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            context.Log($"Converting PDF to Word: {fileName}")
            ApDashboardLog($"📄 Converting PDF to Word: {fileName}", "step")

            ' Use a timeout to prevent indefinite UI thread blocking
            Dim uiTask = SwitchToUi(Function()
                                        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
                                        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                        Dim weCreated As Boolean = False
                                        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel =
                                            Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                                        Dim prevAutoSec As Microsoft.Office.Core.MsoAutomationSecurity =
                                            Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityByUI
                                        Dim prevFileConverters As Object = Nothing
                                        Try
                                            Try
                                                wordApp = DirectCast(GetObject(, "Word.Application"), Microsoft.Office.Interop.Word.Application)
                                            Catch
                                                wordApp = New Microsoft.Office.Interop.Word.Application()
                                                wordApp.Visible = False
                                                weCreated = True
                                            End Try

                                            ' Capture current state BEFORE modifying
                                            prevAlerts = wordApp.DisplayAlerts
                                            prevAutoSec = wordApp.AutomationSecurity

                                            ' Suppress all alerts and macro execution
                                            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
                                            wordApp.ScreenUpdating = False
                                            wordApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable

                                            ' Disable third-party file format converters to prevent modal dialogs
                                            ' from Adobe Acrobat, Foxit, Nuance, etc.
                                            Try
                                                prevFileConverters = wordApp.Options.ConfirmConversions
                                                wordApp.Options.ConfirmConversions = False
                                            Catch
                                            End Try

                                            ' Word can open PDFs and convert them to editable .docx
                                            ' Using Format:=wdOpenFormatAuto (0) lets Word use its BUILT-IN
                                            ' PDF reflow engine rather than deferring to a third-party converter.
                                            doc = wordApp.Documents.Open(
                                                FileName:=att.TempFilePath,
                                                [ReadOnly]:=False,
                                                Visible:=False,
                                                AddToRecentFiles:=False,
                                                ConfirmConversions:=False,
                                                OpenAndRepair:=False,
                                                Format:=0) ' wdOpenFormatAuto = 0

                                            doc.SaveAs2(outputPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                                            Return True
                                        Catch ex As Exception
                                            Debug.WriteLine($"PdfToWord error: {ex.Message}")
                                            Return False
                                        Finally
                                            Try : If doc IsNot Nothing Then doc.Close(False)
                                            Catch : End Try
                                            Try
                                                If wordApp IsNot Nothing Then
                                                    wordApp.DisplayAlerts = prevAlerts
                                                    wordApp.ScreenUpdating = True
                                                    wordApp.AutomationSecurity = prevAutoSec
                                                    Try
                                                        If prevFileConverters IsNot Nothing Then
                                                            wordApp.Options.ConfirmConversions = CBool(prevFileConverters)
                                                        End If
                                                    Catch
                                                    End Try
                                                End If
                                            Catch : End Try
                                            If weCreated AndAlso wordApp IsNot Nothing Then
                                                Try : wordApp.Quit(False) : Catch : End Try
                                            End If
                                        End Try
                                    End Function)

            ' Apply a 120-second timeout to prevent indefinite UI thread blocking
            Dim timeoutTask = Task.Delay(TimeSpan.FromSeconds(120), ct)
            Dim completedTask = Await Task.WhenAny(uiTask, timeoutTask)

            Dim success As Boolean = False
            If completedTask Is uiTask Then
                success = Await uiTask
            Else
                ' Timeout or cancellation
                response.Success = False
                response.Response = $"PDF to Word conversion timed out for '{fileName}'. The PDF may be too large, corrupted, or a third-party converter dialog may be blocking. Check if any dialog is open in Word."
                ApDashboardLog($"⚠ PdfToWord timed out: {fileName}", "warn")
                Return response
            End If

            If success AndAlso File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"Converted '{fileName}' to Word: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). " &
                    "This file can now be used with compare_word_documents. " &
                    "Note: Word does NOT perform OCR — if the PDF is a scanned image, the resulting .docx will contain images without extracted text."
                ApDashboardLog($"✓ Converted to Word: {outputName}", "info")
            Else
                response.Success = False
                response.Response = $"Failed to convert '{fileName}' to Word. The PDF may be image-only, corrupted, or a third-party PDF converter add-in may have interfered. " &
                    "Ensure no PDF add-ins (Adobe Acrobat, Foxit, etc.) are registered as Word file converters."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error converting PDF to Word: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: search_in_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteSearchInAttachmentsTool(
            toolCall As ToolCall, context As ToolExecutionContext, ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim searchTerm = GetArgString(toolCall.Arguments, "search_term")
            If String.IsNullOrWhiteSpace(searchTerm) Then
                response.Success = False
                response.Response = "Missing required parameter: search_term"
                Return response
            End If

            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")

            Dim toSearch As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toSearch = _apCurrentAttachments?.Where(
                    Function(a) targetNames.Any(Function(n) a.OriginalFileName.Equals(n, StringComparison.OrdinalIgnoreCase)) AndAlso
                                Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing).ToList()
            Else
                toSearch = _apCurrentAttachments?.Where(
                    Function(a) Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing AndAlso
                                Not AP_BinaryExtensions.Contains(a.Extension)).ToList()
            End If

            If toSearch Is Nothing OrElse toSearch.Count = 0 Then
                response.Success = False
                response.Response = "No searchable attachments found."
                Return response
            End If

            context.Log($"Searching for '{searchTerm}' in {toSearch.Count} attachment(s)")
            ApDashboardLog($"🔍 Searching for: {searchTerm}", "step")

            Dim sb As New StringBuilder()
            Dim totalMatches = 0

            For Each att In toSearch
                Dim text = Await ReadSingleAttachmentText(att, context)
                If String.IsNullOrWhiteSpace(text) Then Continue For

                Dim lines = text.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim matchLines As New List(Of String)()

                For i = 0 To lines.Length - 1
                    If lines(i).IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Dim lineNum = i + 1
                        Dim excerpt = lines(i).Trim()
                        If excerpt.Length > 200 Then excerpt = excerpt.Substring(0, 200) & "..."
                        matchLines.Add($"  Line {lineNum}: {excerpt}")
                    End If
                Next

                If matchLines.Count > 0 Then
                    sb.AppendLine($"[{att.OriginalFileName}] — {matchLines.Count} match(es)")
                    For Each ml In matchLines.Take(20)
                        sb.AppendLine(ml)
                    Next
                    If matchLines.Count > 20 Then sb.AppendLine($"  ... and {matchLines.Count - 20} more match(es)")
                    sb.AppendLine()
                    totalMatches += matchLines.Count
                End If
            Next

            If totalMatches > 0 Then
                response.Success = True
                response.Response = $"Found {totalMatches} match(es) for '{searchTerm}':" & vbCrLf & sb.ToString().TrimEnd()
            Else
                response.Success = True
                response.Response = $"No matches found for '{searchTerm}' in any attachment."
            End If

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error searching attachments: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: summarize_thread
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Function ExecuteSummarizeThreadTool(toolCall As ToolCall, context As ToolExecutionContext) As ToolResponse
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            If _apCurrentMailInfo Is Nothing Then
                response.Success = False
                response.Response = "No current mail context available."
                Return response
            End If

            context.Log("Extracting email thread")
            ApDashboardLog("📧 Extracting email thread", "step")

            ' Get the monitored mailbox address to exclude autopilot's own messages
            Dim monitoredMailbox = If(_apConfig?.MonitoredMailbox, "").Trim().ToLowerInvariant()

            Dim body = _apCurrentMailInfo.Body
            If String.IsNullOrWhiteSpace(body) Then
                response.Success = True
                response.Response = "The email has no body text."
                Return response
            End If

            ' Parse the email thread by looking for common forwarding/reply separators
            Dim sb As New StringBuilder()
            Dim threadMessages As New List(Of (Sender As String, DateStr As String, Body As String))

            ' Add the current message first
            threadMessages.Add((_apCurrentMailInfo.SenderName & " <" & _apCurrentMailInfo.SenderEmail & ">",
                               _apCurrentMailInfo.ReceivedTime.ToString("yyyy-MM-dd HH:mm"),
                               ""))

            ' Split body on common thread separators
            Dim separatorPatterns = {
                "-----Original Message-----",
                "-----Ursprüngliche Nachricht-----",
                "-----Message d'origine-----",
                "-----Messaggio originale-----",
                "________________________________",
                "From: ", "Von: ", "De : ", "Da: "
            }

            ' Simple thread extraction: split on separator patterns
            Dim currentBody As New StringBuilder()
            Dim lines = body.Split({vbCrLf, vbLf}, StringSplitOptions.None)
            Dim msgIdx = 0

            For Each line In lines
                Dim isSeparator = False
                For Each sep In separatorPatterns
                    If line.TrimStart().StartsWith(sep, StringComparison.OrdinalIgnoreCase) Then
                        ' Save current body to current message
                        If msgIdx < threadMessages.Count Then
                            Dim msg = threadMessages(msgIdx)
                            msg.Body = currentBody.ToString().Trim()
                            threadMessages(msgIdx) = msg
                        End If
                        currentBody.Clear()

                        ' Try to extract sender from "From:" line
                        Dim senderLine = line.Trim()
                        If senderLine.StartsWith("From:", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("Von:", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("De :", StringComparison.OrdinalIgnoreCase) OrElse
                           senderLine.StartsWith("Da:", StringComparison.OrdinalIgnoreCase) Then
                            Dim senderPart = senderLine.Substring(senderLine.IndexOf(":"c) + 1).Trim()
                            threadMessages.Add((senderPart, "", ""))
                            msgIdx = threadMessages.Count - 1
                        Else
                            threadMessages.Add(("(previous message)", "", ""))
                            msgIdx = threadMessages.Count - 1
                        End If

                        isSeparator = True
                        Exit For
                    End If
                Next

                If Not isSeparator Then
                    ' Check for "Sent:" / "Date:" lines to capture date
                    Dim trimmed = line.TrimStart()
                    If trimmed.StartsWith("Sent:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Gesendet:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Date:", StringComparison.OrdinalIgnoreCase) OrElse
                       trimmed.StartsWith("Datum:", StringComparison.OrdinalIgnoreCase) Then
                        If msgIdx < threadMessages.Count Then
                            Dim msg = threadMessages(msgIdx)
                            msg.DateStr = trimmed.Substring(trimmed.IndexOf(":"c) + 1).Trim()
                            threadMessages(msgIdx) = msg
                        End If
                    Else
                        currentBody.AppendLine(line)
                    End If
                End If
            Next

            ' Save last body
            If msgIdx < threadMessages.Count Then
                Dim msg = threadMessages(msgIdx)
                msg.Body = currentBody.ToString().Trim()
                threadMessages(msgIdx) = msg
            End If

            ' Build output, excluding messages from/to the monitored mailbox
            sb.AppendLine($"Email Thread ({threadMessages.Count} message(s)):")
            sb.AppendLine()

            Dim displayIdx = 1
            For Each msg In threadMessages
                ' Exclude monitored mailbox messages
                If Not String.IsNullOrWhiteSpace(monitoredMailbox) AndAlso
                   msg.Sender.ToLowerInvariant().Contains(monitoredMailbox) Then
                    Continue For
                End If

                sb.AppendLine($"── Message {displayIdx} ──")
                sb.AppendLine($"From: {msg.Sender}")
                If Not String.IsNullOrWhiteSpace(msg.DateStr) Then sb.AppendLine($"Date: {msg.DateStr}")
                sb.AppendLine()
                If Not String.IsNullOrWhiteSpace(msg.Body) Then
                    Dim bodyText = msg.Body
                    If bodyText.Length > 5000 Then
                        bodyText = bodyText.Substring(0, 5000) & vbCrLf & "[... truncated ...]"
                    End If
                    sb.AppendLine(bodyText)
                Else
                    sb.AppendLine("(no body text)")
                End If
                sb.AppendLine()
                displayIdx += 1
            Next

            response.Success = True
            response.Response = sb.ToString().TrimEnd()
            ApDashboardLog($"✓ Thread extracted: {displayIdx - 1} message(s)", "info")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting email thread: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: redact_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteRedactPdfTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF file."
                Return response
            End If

            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            Dim mode = If(GetArgString(toolCall.Arguments, "mode"), "prepare").Trim().ToLowerInvariant()
            Dim includeReasonCodes = GetArgBool(toolCall.Arguments, "include_reason_codes", False)
            Dim outputName = GetArgString(toolCall.Arguments, "output_filename")

            ' Validate mode
            If mode <> "prepare" AndAlso mode <> "finalize" AndAlso mode <> "prepare_and_finalize" Then
                mode = "prepare"
            End If

            ' Instruction required for prepare modes
            If (mode = "prepare" OrElse mode = "prepare_and_finalize") AndAlso String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.Response = "Missing required parameter: 'instruction' is required for prepare and prepare_and_finalize modes."
                Return response
            End If

            ' Determine output filename
            Dim baseName = Path.GetFileNameWithoutExtension(att.OriginalFileName)
            Dim suffix = If(mode = "finalize", "_final",
                         If(mode = "prepare_and_finalize", "_redacted_final", "_redacted"))

            If String.IsNullOrWhiteSpace(outputName) Then
                outputName = baseName & suffix & ".pdf"
            End If
            If Not outputName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                outputName &= ".pdf"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                outputName = baseName & suffix & $"_{counter}.pdf"
                outputPath = Path.Combine(_apCurrentTempDir, outputName)
                counter += 1
            End While

            context.Log($"PDF redaction ({mode}): {fileName}" &
                        If(Not String.IsNullOrWhiteSpace(instruction), $" — {instruction}", ""))
            ApDashboardLog($"🔒 PDF redaction ({mode}): {fileName}", "step")

            EnsureApPdfSharpFontResolver()

            If mode = "finalize" Then
                ' Finalize-only: burn in existing annotations
                Await Task.Run(Sub() APRedactFinalizeOnly(att.TempFilePath, outputPath, includeReasonCodes, 300))

                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"PDF finalized (annotations burned in): {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). " &
                    "All redaction boxes are now permanent black rectangles."
                ApDashboardLog($"✓ PDF finalized: {outputName}", "info")

            Else
                ' Prepare (with optional finalize)
                Dim finalize As Boolean = (mode = "prepare_and_finalize")
                Dim result = Await APRedactPdf(att.TempFilePath, outputPath, instruction,
                                                finalize, includeReasonCodes, ct)

                If result Is Nothing Then
                    response.Success = False
                    response.Response = $"Redaction failed for '{fileName}'. The PDF may contain no extractable text " &
                        "(run OCR first), the AI may have returned an empty/unparseable response, or no matching " &
                        "text was found for the identified redactions. You may want to retry."
                    ApDashboardLog($"⚠ Redaction failed for: {fileName}", "warn")
                    Return response
                End If

                If result = "no_redactions" Then
                    response.Success = True
                    response.Response = $"The AI found nothing to redact in '{fileName}' based on the instruction: '{instruction}'."
                    ApDashboardLog($"ℹ No redactions found in: {fileName}", "info")
                    Return response
                End If

                att.OutputFiles.Add(outputPath)
                response.Success = True
                Dim sizeKb = If(File.Exists(outputPath), $" ({New FileInfo(outputPath).Length / 1024:F0} KB)", "")
                response.Response = $"PDF redacted: {result} Output: {outputName}{sizeKb}."
                If Not finalize Then
                    response.Response &= " Note: redaction boxes are currently removable annotations. " &
                        "Call this tool again with mode='finalize' on the output file to make them permanent."
                End If
                ApDashboardLog($"✓ PDF redacted: {outputName}", "info")
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error during PDF redaction: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  HELPER: Ensure PdfSharp font resolver is configured
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Shared _apFontResolverConfigured As Boolean = False

    ''' <summary>
    ''' Ensures a PdfSharp IFontResolver is registered so that font-by-name lookups work.
    ''' Must be called before creating any XFont. Safe to call multiple times.
    ''' </summary>
    Private Shared Sub EnsureApPdfSharpFontResolver()
        If _apFontResolverConfigured Then Return
        Try
            If PdfSharp.Fonts.GlobalFontSettings.FontResolver Is Nothing Then
                PdfSharp.Fonts.GlobalFontSettings.FontResolver = New ApFontResolver()
            End If
            _apFontResolverConfigured = True
        Catch
            ' Already set or locked — ignore
            _apFontResolverConfigured = True
        End Try
    End Sub

    ''' <summary>
    ''' Minimal font resolver for PdfSharp that loads system fonts by reading .ttf files from the Windows Fonts folder.
    ''' </summary>
    Private Class ApFontResolver
        Implements PdfSharp.Fonts.IFontResolver

        Private Shared ReadOnly _fontCache As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)

        ''' <summary>
        ''' Resolves a requested font family/style combination to a concrete font face key.
        ''' </summary>
        Public Function ResolveTypeface(familyName As String, bold As Boolean, italic As Boolean) As PdfSharp.Fonts.FontResolverInfo _
                Implements PdfSharp.Fonts.IFontResolver.ResolveTypeface

            ' Map common family names to Windows font filenames
            Dim key As String = "arial.ttf"
            Select Case familyName.ToLowerInvariant()
                Case "arial", "helvetica"
                    If bold AndAlso italic Then
                        key = "arialbi.ttf"
                    ElseIf bold Then
                        key = "arialbd.ttf"
                    ElseIf italic Then
                        key = "ariali.ttf"
                    Else
                        key = "arial.ttf"
                    End If
                Case "times new roman", "times"
                    If bold AndAlso italic Then
                        key = "timesbi.ttf"
                    ElseIf bold Then
                        key = "timesbd.ttf"
                    ElseIf italic Then
                        key = "timesi.ttf"
                    Else
                        key = "times.ttf"
                    End If
                Case "courier new", "courier"
                    If bold AndAlso italic Then
                        key = "courbi.ttf"
                    ElseIf bold Then
                        key = "courbd.ttf"
                    ElseIf italic Then
                        key = "couri.ttf"
                    Else
                        key = "cour.ttf"
                    End If
                Case "segoe ui"
                    If bold AndAlso italic Then
                        key = "segoeuiz.ttf"
                    ElseIf bold Then
                        key = "segoeuib.ttf"
                    ElseIf italic Then
                        key = "segoeuii.ttf"
                    Else
                        key = "segoeui.ttf"
                    End If
                Case "calibri"
                    If bold AndAlso italic Then
                        key = "calibriz.ttf"
                    ElseIf bold Then
                        key = "calibrib.ttf"
                    ElseIf italic Then
                        key = "calibrii.ttf"
                    Else
                        key = "calibri.ttf"
                    End If
                Case "verdana"
                    If bold AndAlso italic Then
                        key = "verdanaz.ttf"
                    ElseIf bold Then
                        key = "verdanab.ttf"
                    ElseIf italic Then
                        key = "verdanai.ttf"
                    Else
                        key = "verdana.ttf"
                    End If
                Case "tahoma"
                    If bold Then
                        key = "tahomabd.ttf"
                    Else
                        key = "tahoma.ttf"
                    End If
                Case "georgia"
                    If bold AndAlso italic Then
                        key = "georgiaz.ttf"
                    ElseIf bold Then
                        key = "georgiab.ttf"
                    ElseIf italic Then
                        key = "georgiai.ttf"
                    Else
                        key = "georgia.ttf"
                    End If
                Case "trebuchet ms"
                    If bold AndAlso italic Then
                        key = "trebucbi.ttf"
                    ElseIf bold Then
                        key = "trebucbd.ttf"
                    ElseIf italic Then
                        key = "trebucit.ttf"
                    Else
                        key = "trebuc.ttf"
                    End If
                Case Else
                    ' Fallback to Arial
                    If bold Then
                        key = "arialbd.ttf"
                    Else
                        key = "arial.ttf"
                    End If
            End Select

            Return New PdfSharp.Fonts.FontResolverInfo(key)
        End Function

        ''' <summary>
        ''' Loads raw font bytes for a resolved face name from Windows font locations.
        ''' </summary>
        Public Function GetFont(faceName As String) As Byte() _
                Implements PdfSharp.Fonts.IFontResolver.GetFont

            SyncLock _fontCache
                If _fontCache.ContainsKey(faceName) Then Return _fontCache(faceName)
            End SyncLock

            Dim fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts)
            Dim fontPath = Path.Combine(fontsDir, faceName)

            ' Also check Windows\Fonts directly (SpecialFolder.Fonts may return user fonts folder on some systems)
            If Not File.Exists(fontPath) Then
                fontPath = Path.Combine(Environment.GetEnvironmentVariable("SystemRoot"), "Fonts", faceName)
            End If

            If File.Exists(fontPath) Then
                Dim data = File.ReadAllBytes(fontPath)
                SyncLock _fontCache
                    _fontCache(faceName) = data
                End SyncLock
                Return data
            End If

            ' Last resort: return Nothing — PdfSharp will throw a descriptive error
            Return Nothing
        End Function
    End Class

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: extract_data_from_attachments
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Extracts structured/tabular data from one or more attachments using the
    ''' <see cref="SharedLibrary.FactExtractionService.RunFactExtractionAsync"/> pipeline.
    ''' Returns schema + rows as JSON for downstream tool-chaining (create_excel, create_word, etc.).
    ''' </summary>
    Private Async Function ExecuteExtractDataFromAttachmentsTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim instruction = GetArgString(toolCall.Arguments, "instruction")
            If String.IsNullOrWhiteSpace(instruction) Then
                response.Success = False
                response.Response = "Missing required parameter: instruction"
                Return response
            End If

            Dim schemaSpec = GetArgString(toolCall.Arguments, "schema")
            Dim targetNames = GetArgStringArray(toolCall.Arguments, "attachment_names")
            Dim sortColumn = GetArgInt(toolCall.Arguments, "sort_column", 0)
            Dim sortDirection = If(GetArgString(toolCall.Arguments, "sort_direction"), "ASC").Trim().ToUpperInvariant()
            If sortDirection <> "ASC" AndAlso sortDirection <> "DESC" Then sortDirection = "ASC"
            Dim dateColumnsText = If(GetArgString(toolCall.Arguments, "date_columns"), "")

            ' ── Resolve attachments to process ──
            Dim toProcess As List(Of AutoPilotAttachmentInfo)
            If targetNames.Count > 0 Then
                toProcess = New List(Of AutoPilotAttachmentInfo)()
                For Each name In targetNames
                    Dim att = FindAttachment(name)
                    If att IsNot Nothing AndAlso Not att.IsOverSizeLimit AndAlso att.TempFilePath IsNot Nothing Then
                        toProcess.Add(att)
                    End If
                Next
            Else
                ' Process all readable non-binary attachments
                toProcess = _apCurrentAttachments?.Where(
                    Function(a) Not a.IsOverSizeLimit AndAlso a.TempFilePath IsNot Nothing AndAlso
                                Not AP_BinaryExtensions.Contains(a.Extension)).ToList()
            End If

            If toProcess Is Nothing OrElse toProcess.Count = 0 Then
                response.Success = False
                response.Response = "No processable attachments found for data extraction."
                Return response
            End If

            context.Log($"Extracting structured data from {toProcess.Count} file(s): {instruction}")
            ApDashboardLog($"📊 Fact extraction: {toProcess.Count} file(s)", "step")

            ' ── Parse optional fixed schema ──
            Dim fixedSchema As System.Collections.Generic.List(Of FactExtractionService.ExtractionSchemaColumn) = Nothing
            If Not String.IsNullOrWhiteSpace(schemaSpec) Then
                fixedSchema = FactExtractionService.ParseUserSchemaSpec(schemaSpec)
                ' Detect sort column from * marker if not explicitly set
                If (fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0) AndAlso sortColumn = 0 Then
                    sortColumn = FactExtractionService.DetectSortColumnFromSpec(schemaSpec)
                End If
            End If

            ' ── Parse date columns ──
            Dim dateCols As New System.Collections.Generic.List(Of Integer)
            For Each part In dateColumnsText.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim n As Integer
                If Integer.TryParse(part.Trim(), n) AndAlso n > 0 Then dateCols.Add(n)
            Next

            ' ── Build file path list from resolved attachments ──
            Dim filePaths As New List(Of String)()
            For Each att In toProcess
                filePaths.Add(att.TempFilePath)
            Next

            ' ── Set up the extraction instruction for InterpolateAtRuntime ──
            ' ── Set up the extraction instruction for InterpolateAtRuntime ──
            Dim savedOtherPrompt = OtherPrompt
            Dim savedOutputLanguage = OutputLanguage
            Try
                OtherPrompt = instruction

                ' Let the LLM decide the output language; fall back to the user's primary language
                Dim requestedLanguage = GetArgString(toolCall.Arguments, "output_language")
                OutputLanguage = If(Not String.IsNullOrWhiteSpace(requestedLanguage), requestedLanguage.Trim(), INI_Language1)

                Dim useSecondApi As Boolean = (_apConfig IsNot Nothing AndAlso _apConfig.UseSecondApi)
                Dim cancelled As Boolean = False

                ' ── Adapt GetFileContent for the service ──
                ' RunFactExtractionAsync expects Func(Of String, Boolean, Boolean, Boolean, Task(Of String))
                ' Parameters: (path, silent, doOcr, askUser)
                ' We bridge to the Outlook text extraction pipeline (ReadSingleAttachmentText cache + fallbacks)
                Dim getFileContentFunc As Func(Of String, Boolean, Boolean, Boolean, Task(Of String)) =
                    Async Function(filePath As String, silent As Boolean, doOcr As Boolean, askUser As Boolean) As Task(Of String)
                        ' First, try to find the file in the current attachments by path and reuse cached text
                        Dim matchedAtt = toProcess.FirstOrDefault(
                            Function(a) a.TempFilePath IsNot Nothing AndAlso
                                        a.TempFilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                        If matchedAtt IsNot Nothing Then
                            Dim cachedText = Await ReadSingleAttachmentText(matchedAtt, context)
                            If Not String.IsNullOrWhiteSpace(cachedText) Then Return cachedText
                        End If

                        ' Fallback: try extraction methods directly on path
                        Dim text As String = Nothing
                        Dim label As String = Nothing
                        Try
                            If TryExtractOfficeText(filePath, text, label) AndAlso Not String.IsNullOrWhiteSpace(text) Then Return text
                        Catch : End Try
                        Try
                            Dim ext = Path.GetExtension(filePath).ToLowerInvariant()
                            If ext = ".xlsx" OrElse ext = ".xls" Then
                                text = ExtractExcelText(filePath)
                                If Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error") Then Return text
                            ElseIf ext = ".pptx" Then
                                text = ExtractPowerPointText(filePath)
                                If Not String.IsNullOrWhiteSpace(text) AndAlso Not text.StartsWith("Error") Then Return text
                            End If
                        Catch : End Try
                        Try
                            If TryExtractTextLike(filePath, text, label) AndAlso Not String.IsNullOrWhiteSpace(text) Then Return text
                        Catch : End Try
                        If Path.GetExtension(filePath).ToLowerInvariant() = ".pdf" Then
                            Try
                                text = Await SharedMethods.ReadPdfAsText(filePath, ReturnErrorInsteadOfEmpty:=True, DoOCR:=doOcr, AskUser:=False)
                                If Not String.IsNullOrWhiteSpace(text) Then Return text
                            Catch : End Try
                        End If
                        Return ""
                    End Function

                ' ── Adapt LLM function for the service ──
                ' RunFactExtractionAsync expects Func(Of String, String, String, String, Integer, Boolean, Boolean, Task(Of String))
                Dim llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Task(Of String)) =
                    Async Function(sysPrompt As String, userText As String, model As String, temp As String,
                                   timeout As Integer, useSecond As Boolean, hideSplash As Boolean) As Task(Of String)
                        Return Await ThisAddIn.LLM(sysPrompt, userText, model, temp, timeout, useSecond, True, cancellationToken:=ct)
                    End Function

                ' ── Run the extraction ──
                Dim result = Await FactExtractionService.RunFactExtractionAsync(
                    filePaths,
                    instruction,
                    dateCols,
                    sortColumn,
                    sortDirection,
                    False,          ' doOcr — we already extracted text above
                    useSecondApi,
                    _apCurrentTempDir,
                    AddressOf InterpolateAtRuntime,
                    llmFunc,
                    getFileContentFunc,
                    _context,
                    fixedSchema,
                    Nothing,        ' clampFrom
                    Nothing,        ' clampTo
                    Sub(cur, total, label)
                        ApDashboardLog($"  📊 [{cur}/{total}] {label}", "step")
                    End Sub,
                    0,              ' mergeDateColumn
                    False,          ' mergeRowsViaLlm
                    Nothing,        ' mergeInstruction
                    Function() cancelled OrElse ct.IsCancellationRequested)

                If result Is Nothing OrElse result.Rows.Count = 0 Then
                    Dim errMsg = "No data could be extracted from the provided file(s)."
                    If result IsNot Nothing AndAlso result.Errors.Count > 0 Then
                        errMsg &= " Errors: " & String.Join("; ", result.Errors.Take(5))
                    End If
                    If result IsNot Nothing AndAlso result.FailedFileNames.Count > 0 Then
                        errMsg &= " Failed files: " & String.Join(", ", result.FailedFileNames)
                    End If
                    response.Success = False
                    response.Response = errMsg
                    ApDashboardLog($"⚠ Fact extraction returned no data", "warn")
                    Return response
                End If

                ' ── Format result as JSON for the LLM ──
                Dim jResult As New JObject()

                ' Schema array
                Dim jSchema As New JArray()
                For Each col In result.Schema
                    jSchema.Add(New JObject From {
                        {"name", col.Name},
                        {"type", col.Type}
                    })
                Next
                jResult("schema") = jSchema

                ' Rows as array of arrays
                Dim jRows As New JArray()
                For Each row In result.Rows
                    Dim jRow As New JArray()
                    For Each cellVal In row.Values
                        jRow.Add(If(cellVal Is Nothing, "", cellVal.ToString()))
                    Next
                    jRows.Add(jRow)
                Next
                jResult("rows") = jRows

                ' Metadata
                jResult("total_rows") = result.Rows.Count
                jResult("total_columns") = result.Schema.Count
                jResult("files_processed") = result.ProcessedFiles
                jResult("files_failed") = result.FailedFiles
                If result.FailedFileNames.Count > 0 Then
                    jResult("failed_files") = New JArray(result.FailedFileNames.ToArray())
                End If
                If result.Errors.Count > 0 Then
                    jResult("errors") = New JArray(result.Errors.Take(10).ToArray())
                End If

                Dim jsonString = jResult.ToString(Newtonsoft.Json.Formatting.None)

                ' Truncate if extremely large to stay within LLM context limits
                If jsonString.Length > 200000 Then
                    ' Rebuild with fewer rows
                    Dim maxRows = Math.Max(1, CInt(result.Rows.Count * (200000.0 / jsonString.Length)))
                    Dim jRowsTruncated As New JArray()
                    For i = 0 To Math.Min(maxRows - 1, result.Rows.Count - 1)
                        Dim jRow As New JArray()
                        For Each cellVal In result.Rows(i).Values
                            jRow.Add(If(cellVal Is Nothing, "", cellVal.ToString()))
                        Next
                        jRowsTruncated.Add(jRow)
                    Next
                    jResult("rows") = jRowsTruncated
                    jResult("total_rows") = result.Rows.Count
                    jResult("rows_returned") = jRowsTruncated.Count
                    jResult("truncated") = True
                    jsonString = jResult.ToString(Newtonsoft.Json.Formatting.None)
                End If

                response.Success = True
                response.Response = jsonString

                Dim schemaNames = String.Join(", ", result.Schema.Select(Function(c) c.Name))
                ApDashboardLog($"✓ Fact extraction: {result.Rows.Count} rows, {result.Schema.Count} columns ({schemaNames}), {result.ProcessedFiles} file(s)", "info")

            Finally
                OtherPrompt = savedOtherPrompt
                OutputLanguage = savedOutputLanguage
            End Try

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error extracting data from attachments: {ex.Message}"
        End Try

        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: report_inability
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Handles the report_inability tool call. Generates helpful suggestions by:
    ''' (1) noting attachment size limits if relevant, (2) consulting the HelpMeInky
    ''' manual for Red Ink add-in features, and (3) optionally querying the
    ''' InternetResearch model for alternative online tools.
    ''' The suggestions are returned in the tool response for LLM incorporation,
    ''' with explicit instruction that suggestions may be rephrased but not omitted.
    ''' </summary>
    Private Async Function ExecuteReportInabilityTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = AP_Tool_ReportInability,
            .OriginalCallJson = toolCall.RawJson
        }

        Dim reason = GetArgString(toolCall.Arguments, "reason")
        If String.IsNullOrWhiteSpace(reason) Then reason = "unspecified"
        ApDashboardLog($"📋 Inability reported: {reason}", "warn")

        Dim sb As New StringBuilder()

        ' ── Attachment size limit detail ──
        If _apCurrentAttachments IsNot Nothing AndAlso _apCurrentAttachments.Any(Function(a) a.IsOverSizeLimit) Then
            Dim limitMB = _apConfig.MaxAttachmentBytes / 1024.0 / 1024.0
            Dim oversizedNames = _apCurrentAttachments.Where(Function(a) a.IsOverSizeLimit).
                Select(Function(a) $"{a.OriginalFileName} ({a.SizeBytes / 1024.0 / 1024.0:F1} MB)").ToList()
            sb.AppendLine($"The maximum permitted attachment size is {limitMB:F0} MB. " &
                          $"The following file(s) exceeded this limit: {String.Join(", ", oversizedNames)}. " &
                          $"Advise the sender to send smaller files or split large documents.")
            sb.AppendLine()
        End If

        ' ── Red Ink add-in suggestion (via HelpMeInky manual) ──
        Try
            Dim redInkSuggestion = Await GetRedInkSuggestionAsync(
                _apCurrentMailInfo, reason, ct)
            If Not String.IsNullOrWhiteSpace(redInkSuggestion) Then
                sb.AppendLine(redInkSuggestion.Trim())
                sb.AppendLine()
            End If
        Catch ex As System.Exception
            ApDashboardLog($"Red Ink suggestion failed: {ex.Message}", "warn")
        End Try

        ' ── Internet alternative suggestion ──
        Dim hasInternetSuggestion As Boolean = False
        Try
            Dim internetSuggestion = Await GetInternetAlternativeSuggestionAsync(
                _apCurrentMailInfo, reason, ct)
            If Not String.IsNullOrWhiteSpace(internetSuggestion) Then
                sb.AppendLine(internetSuggestion.Trim())
                sb.AppendLine()
                hasInternetSuggestion = True
            End If
        Catch ex As System.Exception
            ApDashboardLog($"Internet suggestion failed: {ex.Message}", "warn")
        End Try

        If sb.Length = 0 Then
            sb.AppendLine("No specific suggestions available. Advise the sender to try again, rephrase their request, or contact the operator for assistance.")
        End If

        ' ── Instruction to the LLM on how to use this tool output ──
        sb.AppendLine()
        sb.AppendLine("INSTRUCTIONS FOR YOUR RESPONSE:")
        sb.AppendLine("Include ALL of the above suggestions in your reply to the sender. " &
                      "You may rephrase the suggestions to fit naturally into your response, " &
                      "but do NOT omit any of them.")
        If hasInternetSuggestion Then
            sb.AppendLine("MANDATORY: Your reply MUST end with the following disclaimer paragraph, " &
                          "copied VERBATIM (do not rephrase, shorten or omit it):")
            sb.AppendLine("Please note: Third-party services and tools may only be used if permitted by your organization's policies. " &
                          "Before using any external service or tool, ensure it meets your corporate security, confidentiality, and data protection requirements.")
        End If

        response.Success = True
        response.Response = sb.ToString().TrimEnd()
        ApDashboardLog($"✓ Inability suggestions generated ({response.Response.Length} chars)", "step")
        Return response
    End Function


    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL EXECUTION: overlay_pdf
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Async Function ExecuteOverlayPdfTool(
            toolCall As ToolCall,
            context As ToolExecutionContext,
            ct As CancellationToken) As Task(Of ToolResponse)

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId, .ToolName = toolCall.ToolName, .Timestamp = DateTime.UtcNow
        }

        Try
            Dim fileName = GetArgString(toolCall.Arguments, "attachment_name")
            If String.IsNullOrWhiteSpace(fileName) Then
                response.Success = False
                response.Response = "Missing required parameter: attachment_name"
                Return response
            End If

            Dim att = FindAttachment(fileName)
            If att Is Nothing Then
                response.Success = False
                response.Response = $"Attachment '{fileName}' not found."
                Return response
            End If

            If Not att.TempFilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                response.Success = False
                response.Response = $"'{fileName}' is not a PDF file."
                Return response
            End If

            ' Parse elements array
            Dim elementsArray As JArray = Nothing
            If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("elements") Then
                Dim elemObj = toolCall.Arguments("elements")
                If TypeOf elemObj Is JArray Then elementsArray = DirectCast(elemObj, JArray)
            End If

            If elementsArray Is Nothing OrElse elementsArray.Count = 0 Then
                response.Success = False
                response.Response = "Missing required parameter: elements (must be a non-empty array)"
                Return response
            End If

            ' Determine output filename
            Dim outputName = GetArgString(toolCall.Arguments, "output_filename")
            If String.IsNullOrWhiteSpace(outputName) Then
                outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & "_overlay.pdf"
            End If
            If Not outputName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then
                outputName &= ".pdf"
            End If

            Dim outputPath = Path.Combine(_apCurrentTempDir, outputName)

            ' Prevent filename collision
            Dim counter = 1
            While File.Exists(outputPath)
                outputName = Path.GetFileNameWithoutExtension(att.OriginalFileName) & $"_overlay_{counter}.pdf"
                outputPath = Path.Combine(_apCurrentTempDir, outputName)
                counter += 1
            End While

            context.Log($"Overlaying {elementsArray.Count} element(s) on: {fileName}")
            ApDashboardLog($"🖌 Overlaying {elementsArray.Count} element(s) on: {fileName}", "step")

            EnsureApPdfSharpFontResolver()

            ' Pre-resolve all image attachments to avoid repeated lookups
            Dim imageCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each elemObj As JObject In elementsArray
                Dim elemType = If(elemObj.Value(Of String)("type"), "text").ToLowerInvariant()
                If elemType = "image" Then
                    Dim imgName = elemObj.Value(Of String)("image_attachment_name")
                    If Not String.IsNullOrWhiteSpace(imgName) AndAlso Not imageCache.ContainsKey(imgName) Then
                        Dim imgAtt = FindAttachment(imgName)
                        If imgAtt IsNot Nothing AndAlso imgAtt.TempFilePath IsNot Nothing AndAlso File.Exists(imgAtt.TempFilePath) Then
                            imageCache(imgName) = imgAtt.TempFilePath
                        End If
                    End If
                End If
            Next

            ' Work on a temp copy to avoid source file lock issues
            Dim tempWorkPath = outputPath & ".tmp_" & Guid.NewGuid().ToString("N") & ".pdf"
            Dim textCount = 0
            Dim imageCount = 0

            Try
                File.Copy(att.TempFilePath, tempWorkPath, True)

                Using doc = PdfSharp.Pdf.IO.PdfReader.Open(tempWorkPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Modify)
                    Dim totalPages = doc.PageCount

                    For Each elemObj As JObject In elementsArray
                        Dim elemType = If(elemObj.Value(Of String)("type"), "text").ToLowerInvariant()
                        Dim pagesSpec = If(elemObj.Value(Of String)("pages"), "all").Trim().ToLowerInvariant()
                        Dim x As Double = GetJDouble(elemObj, "x", 0)
                        Dim y As Double = GetJDouble(elemObj, "y", 0)
                        Dim rotation As Double = GetJDouble(elemObj, "rotation", 0)
                        Dim opacity As Double = GetJDouble(elemObj, "opacity", 1.0)

                        ' Resolve target page indices (0-based)
                        Dim pageIndices = ResolvePageIndices(pagesSpec, totalPages)
                        If pageIndices.Count = 0 Then Continue For

                        For Each pageIdx In pageIndices
                            If pageIdx < 0 OrElse pageIdx >= totalPages Then Continue For
                            Dim page = doc.Pages(pageIdx)

                            Using gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page, PdfSharp.Drawing.XGraphicsPdfPageOptions.Append)
                                If elemType = "text" Then
                                    ' ── TEXT ELEMENT ──
                                    Dim text = If(elemObj.Value(Of String)("text"), "")
                                    If String.IsNullOrEmpty(text) Then Continue For

                                    ' Handle \n escape sequences for multi-line text
                                    text = text.Replace("\\n", vbLf).Replace("\n", vbLf)

                                    Dim fontFamily = If(elemObj.Value(Of String)("font_family"), "Arial")
                                    Dim fontSize As Double = GetJDouble(elemObj, "font_size", 12)
                                    Dim isBold = GetJBool(elemObj, "bold")
                                    Dim isItalic = GetJBool(elemObj, "italic")
                                    Dim hAlign = If(elemObj.Value(Of String)("h_align"), "left").ToLowerInvariant()
                                    Dim maxWidth As Double = GetJDouble(elemObj, "max_width", 0)
                                    Dim fontColorHex = If(elemObj.Value(Of String)("font_color"), "#000000")

                                    ' Build font style
                                    Dim fontStyle As PdfSharp.Drawing.XFontStyleEx = PdfSharp.Drawing.XFontStyleEx.Regular
                                    If isBold AndAlso isItalic Then
                                        fontStyle = PdfSharp.Drawing.XFontStyleEx.BoldItalic
                                    ElseIf isBold Then
                                        fontStyle = PdfSharp.Drawing.XFontStyleEx.Bold
                                    ElseIf isItalic Then
                                        fontStyle = PdfSharp.Drawing.XFontStyleEx.Italic
                                    End If

                                    Dim font = New PdfSharp.Drawing.XFont(fontFamily, fontSize, fontStyle)

                                    ' Parse color
                                    Dim brush As PdfSharp.Drawing.XBrush = PdfSharp.Drawing.XBrushes.Black
                                    Try
                                        Dim colorHex = fontColorHex.TrimStart("#"c)
                                        If colorHex.Length = 6 Then
                                            Dim r = System.Convert.ToInt32(colorHex.Substring(0, 2), 16)
                                            Dim g = System.Convert.ToInt32(colorHex.Substring(2, 2), 16)
                                            Dim b = System.Convert.ToInt32(colorHex.Substring(4, 2), 16)
                                            Dim alphaInt = CInt(Math.Round(Math.Max(0, Math.Min(1, opacity)) * 255))
                                            brush = New PdfSharp.Drawing.XSolidBrush(
                                                PdfSharp.Drawing.XColor.FromArgb(alphaInt, r, g, b))
                                        End If
                                    Catch
                                    End Try

                                    ' Determine string format for alignment
                                    Dim xFormat As New PdfSharp.Drawing.XStringFormat()
                                    xFormat.LineAlignment = PdfSharp.Drawing.XLineAlignment.Near
                                    Select Case hAlign
                                        Case "center" : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Center
                                        Case "right" : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Far
                                        Case Else : xFormat.Alignment = PdfSharp.Drawing.XStringAlignment.Near
                                    End Select

                                    ' Apply rotation if specified
                                    Dim state As PdfSharp.Drawing.XGraphicsState = Nothing
                                    If rotation <> 0 Then
                                        state = gfx.Save()
                                        gfx.TranslateTransform(x, y)
                                        gfx.RotateTransform(rotation)
                                        gfx.TranslateTransform(-x, -y)
                                    End If

                                    ' Handle multi-line text
                                    Dim lines = text.Split({vbLf}, StringSplitOptions.None)
                                    Dim lineHeight = fontSize * 1.25
                                    Dim currentY = y

                                    For Each line In lines
                                        If maxWidth > 0 Then
                                            Dim rect As New PdfSharp.Drawing.XRect(x, currentY, maxWidth, lineHeight)
                                            gfx.DrawString(line, font, brush, rect, xFormat)
                                        Else
                                            Dim drawPoint As New PdfSharp.Drawing.XPoint(x, currentY)
                                            gfx.DrawString(line, font, brush, drawPoint, xFormat)
                                        End If
                                        currentY += lineHeight
                                    Next

                                    If state IsNot Nothing Then gfx.Restore(state)
                                    textCount += 1

                                ElseIf elemType = "image" Then
                                    ' ── IMAGE ELEMENT ──
                                    Dim imgName = elemObj.Value(Of String)("image_attachment_name")
                                    If String.IsNullOrWhiteSpace(imgName) Then Continue For

                                    Dim imgPath As String = Nothing
                                    If Not imageCache.TryGetValue(imgName, imgPath) OrElse
                                       String.IsNullOrEmpty(imgPath) OrElse Not File.Exists(imgPath) Then
                                        context.Log($"Image attachment not found: {imgName}")
                                        Continue For
                                    End If

                                    Dim imgWidth As Double = GetJDouble(elemObj, "width", 0)
                                    Dim imgHeight As Double = GetJDouble(elemObj, "height", 0)

                                    ' Load image via stream to support all formats
                                    Using imgStream As New FileStream(imgPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                                        Using xImg = PdfSharp.Drawing.XImage.FromStream(imgStream)
                                            ' Default to native size if not specified
                                            If imgWidth <= 0 AndAlso imgHeight <= 0 Then
                                                imgWidth = xImg.PointWidth
                                                imgHeight = xImg.PointHeight
                                            ElseIf imgWidth > 0 AndAlso imgHeight <= 0 Then
                                                ' Scale proportionally
                                                imgHeight = xImg.PointHeight * (imgWidth / xImg.PointWidth)
                                            ElseIf imgHeight > 0 AndAlso imgWidth <= 0 Then
                                                imgWidth = xImg.PointWidth * (imgHeight / xImg.PointHeight)
                                            End If

                                            ' Apply rotation if specified
                                            Dim state As PdfSharp.Drawing.XGraphicsState = Nothing
                                            If rotation <> 0 Then
                                                state = gfx.Save()
                                                Dim cx = x + imgWidth / 2
                                                Dim cy = y + imgHeight / 2
                                                gfx.TranslateTransform(cx, cy)
                                                gfx.RotateTransform(rotation)
                                                gfx.TranslateTransform(-cx, -cy)
                                            End If

                                            ' Apply opacity for images by drawing on a separate layer
                                            ' PdfSharp does not directly support image opacity, but we can
                                            ' use a workaround with XGraphics state if needed.
                                            ' For now, draw directly (opacity < 1 is best-effort).
                                            gfx.DrawImage(xImg, x, y, imgWidth, imgHeight)

                                            If state IsNot Nothing Then gfx.Restore(state)
                                            imageCount += 1
                                        End Using
                                    End Using
                                End If
                            End Using
                        Next
                    Next

                    doc.Save(tempWorkPath)
                End Using

                ' Move temp to final
                If File.Exists(outputPath) Then File.Delete(outputPath)
                File.Move(tempWorkPath, outputPath)

            Finally
                Try : If File.Exists(tempWorkPath) Then File.Delete(tempWorkPath)
                Catch : End Try
            End Try

            If File.Exists(outputPath) Then
                att.OutputFiles.Add(outputPath)
                response.Success = True
                response.Response = $"PDF overlay complete: {textCount} text element(s) and {imageCount} image element(s) placed. " &
                    $"Output: {outputName} ({New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply."
                ApDashboardLog($"✓ PDF overlay: {outputName} ({textCount} text, {imageCount} image)", "info")
            Else
                response.Success = False
                response.Response = "Failed to create overlaid PDF."
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled."
            response.Response = response.ErrorMessage
        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            response.Response = $"Error overlaying PDF: {ex.Message}"
        End Try

        Return response
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  OVERLAY PDF HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Parses a page specification string into a list of 0-based page indices.
    ''' Supports: "all", "1", "1,3,5", "2-5", "1,3-5,8".
    ''' </summary>
    Private Shared Function ResolvePageIndices(pagesSpec As String, totalPages As Integer) As List(Of Integer)
        Dim result As New List(Of Integer)()
        If String.IsNullOrWhiteSpace(pagesSpec) OrElse pagesSpec = "all" Then
            For i = 0 To totalPages - 1
                result.Add(i)
            Next
            Return result
        End If

        ' Split on commas, then handle each token
        For Each token In pagesSpec.Split(","c)
            Dim trimmed = token.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For

            Dim dashIdx = trimmed.IndexOf("-"c)
            If dashIdx > 0 Then
                ' Range: "2-5"
                Dim startStr = trimmed.Substring(0, dashIdx).Trim()
                Dim endStr = trimmed.Substring(dashIdx + 1).Trim()
                Dim startPage As Integer
                Dim endPage As Integer
                If Integer.TryParse(startStr, startPage) AndAlso Integer.TryParse(endStr, endPage) Then
                    startPage = Math.Max(1, startPage)
                    endPage = Math.Min(totalPages, endPage)
                    For i = startPage To endPage
                        If Not result.Contains(i - 1) Then result.Add(i - 1)
                    Next
                End If
            Else
                ' Single page: "3"
                Dim pageNum As Integer
                If Integer.TryParse(trimmed, pageNum) AndAlso pageNum >= 1 AndAlso pageNum <= totalPages Then
                    If Not result.Contains(pageNum - 1) Then result.Add(pageNum - 1)
                End If
            End If
        Next

        Return result
    End Function

    ''' <summary>
    ''' Reads a Double value from a JObject token with a fallback default.
    ''' </summary>
    Private Shared Function GetJDouble(obj As JObject, key As String, defaultVal As Double) As Double
        Dim token = obj(key)
        If token Is Nothing Then Return defaultVal
        Dim result As Double
        If Double.TryParse(token.ToString(), Globalization.NumberStyles.Any,
                          Globalization.CultureInfo.InvariantCulture, result) Then
            Return result
        End If
        Return defaultVal
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL IDENTIFICATION
    ' ═══════════════════════════════════════════════════════════════════════════

    Friend Function IsAutoPilotInternalTool(toolName As String) As Boolean
        Select Case toolName
            Case AP_Tool_ProcessWordDoc,
                 AP_Tool_CommentWordDoc,
                 AP_Tool_ExtractPdfText,
                 AP_Tool_MergePdfs,
                 AP_Tool_ReadAttachment,
                 AP_Tool_ListAttachments,
                 AP_Tool_DescribeBinary,
                 AP_Tool_CompareWordDocs,
                 AP_Tool_ReadWordDocDetails,
                 AP_Tool_CreatePdfFromText,
                 AP_Tool_ExtractExcelData,
                 AP_Tool_SplitPdf,
                 AP_Tool_AddPdfWatermark,
                 AP_Tool_WordToPdf,
                 AP_Tool_SearchInAttachments,
                 AP_Tool_SummarizeThread,
                 AP_Tool_PdfToWord,
                 AP_Tool_CreateWordDoc,
                 AP_Tool_CreateExcel,
                 AP_Tool_CreatePowerPoint,
                 AP_Tool_CreateCodeFile,
                 AP_Tool_CommentPdf,
                 AP_Tool_ExtractDataFromAttachments,
                 AP_Tool_RedactPdf,
                 AP_Tool_OverlayPdf,
                 AP_Tool_ReportInability
                Return True
            Case Else
                Return False
        End Select
    End Function

End Class