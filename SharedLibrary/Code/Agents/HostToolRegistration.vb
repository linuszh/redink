' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: HostToolRegistration.vb
' Purpose: Declares which tools are internal (registered on the host),
'          which are deliverable-producing (used for user-facing output),
'          and which are memory-management tools used across the agentic layer.
'
' Tool Categories:
'  - OutlookAutoPilotToolNames: Full set of Outlook autopilot capabilities.
'  - WordHostInternalToolNames: Word document interop via IWordDocumentHost.
'  - OutlookDeliverableToolNames: Tools that produce user-facing deliverables.
'  - WordDeliverableToolNames: Word-specific deliverable tools.
'  - Common internal tools: memory, text, workspace, and common utilities.
' =============================================================================


Option Explicit On
Option Strict On

Namespace Agents

    Public Module HostToolRegistration

        Private ReadOnly OutlookAutoPilotToolNames As String() = New String() {
            "process_word_document",
            "comment_word_document",
            "extract_pdf_text",
            "merge_pdfs",
            "read_attachment",
            "list_attachments",
            "describe_binary_attachment",
            "compare_word_documents",
            "read_word_document_details",
            "create_pdf_from_text",
            "extract_excel_data",
            "split_pdf",
            "add_pdf_watermark",
            "word_to_pdf",
            "search_in_attachments",
            "summarize_thread",
            "pdf_to_word",
            "create_word_document",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "comment_pdf_document",
            "extract_data_from_attachments",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image",
            "web_grounding",
            "manage_scheduled_tasks",
            "manage_user_memory",
            "manage_user_files",
            "report_inability"
        }

        Private ReadOnly WordHostInternalToolNames As String() = New String() {
            "word_doc_read",
            "word_doc_edit",
            "word_doc_create",
            "word_doc_export_pdf"
        }

        Private ReadOnly OutlookDeliverableToolNames As String() = New String() {
            "download_web_files",
            WorkspaceTools.ToolWrite,
            TextTools.ToolWrite,
            "merge_pdfs",
            "create_pdf_from_text",
            "split_pdf",
            "add_pdf_watermark",
            "word_to_pdf",
            "pdf_to_word",
            "create_word_document",
            "create_excel_spreadsheet",
            "create_powerpoint",
            "create_code_file",
            "redact_pdf",
            "overlay_pdf",
            "create_audio_file",
            "generate_image"
        }

        Private ReadOnly WordDeliverableToolNames As String() = New String() {
            "download_web_files",
            WorkspaceTools.ToolWrite,
            TextTools.ToolWrite,
            WordTools.ToolWrite,
            WordTools.ToolMarkup,
            WordTools.ToolApplyTemplate,
            WordTools.ToolSaveAs,
            "word_doc_create",
            "word_doc_edit",
            "word_doc_export_pdf"
        }

        Private Iterator Function EnumerateCommonInternalToolNames() As IEnumerable(Of String)
            Yield "retrieve_web_content"
            Yield "download_web_files"
            Yield "internet_search"
            Yield "knowledge_search"
            Yield ToolLoaderTool.LoaderToolName

            Yield MemoryTools.ToolPut
            Yield MemoryTools.ToolGet
            Yield MemoryTools.ToolList
            Yield MemoryTools.ToolDelete

            Yield TextTools.ToolRead
            Yield TextTools.ToolWrite
            Yield TextTools.ToolSearch

            Yield WorkspaceTools.ToolGet
            Yield WorkspaceTools.ToolInventory
            Yield WorkspaceTools.ToolRead
            Yield WorkspaceTools.ToolReadMany
            Yield WorkspaceTools.ToolWrite
            Yield WorkspaceTools.ToolSearch
            Yield WorkspaceTools.ToolCopy
            Yield WorkspaceTools.ToolMove
            Yield WorkspaceTools.ToolRename
            Yield WorkspaceTools.ToolDelete
            Yield WorkspaceTools.ToolMakeDir
            Yield WorkspaceTools.ToolExtractText
            Yield WorkspaceTools.ToolExtractTextMany

            Yield JsRunTool.ToolName
            Yield SkillInvokeTool.ToolName

            Yield SharedLibrary.M365ToolService.SearchToolName
            Yield SharedLibrary.M365ToolService.GetMailToolName
            Yield SharedLibrary.M365ToolService.GetMailThreadToolName
            Yield SharedLibrary.M365ToolService.GetFileToolName
            Yield SharedLibrary.M365ToolService.GetEventToolName
            Yield SharedLibrary.M365ToolService.GetChatThreadToolName
            Yield SharedLibrary.M365ToolService.GetOneNotePageToolName
        End Function

        Private Iterator Function EnumerateWordSharedToolNames() As IEnumerable(Of String)
            For Each name In EnumerateCommonInternalToolNames()
                Yield name
            Next

            Yield WordTools.ToolExtract
            Yield WordTools.ToolSearch
            Yield WordTools.ToolWrite
            Yield WordTools.ToolMarkup
            Yield WordTools.ToolCommentAdd
            Yield WordTools.ToolCommentList
            Yield WordTools.ToolCommentRemove
            Yield WordTools.ToolFormat
            Yield WordTools.ToolApplyTemplate
            Yield WordTools.ToolSaveAs

            Yield WordDocTools.ToolListOpen
            Yield WordDocTools.ToolGetActive
            Yield WordDocTools.ToolExtract
            Yield WordDocTools.ToolSearch
            Yield WordDocTools.ToolListComments
            Yield WordDocTools.ToolInsert
            Yield WordDocTools.ToolReplace
            Yield WordDocTools.ToolCommentAdd
            Yield WordDocTools.ToolFormat
        End Function

        Private Sub RegisterSet(host As ToolingHostKind, names As IEnumerable(Of String))
            If names Is Nothing Then Return

            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each rawName As String In names
                Dim name As String = If(rawName, "").Trim()
                If name = "" Then Continue For

                If seen.Add(name) Then
                    ToolExecutorRegistry.RegisterInternal(host, name)
                End If
            Next
        End Sub

        Public Sub RegisterAll(host As ToolingHostKind)
            Select Case host
                Case ToolingHostKind.Outlook
                    RegisterOutlookInternals()
                Case ToolingHostKind.Word
                    RegisterWordInternals()
            End Select
        End Sub

        Public Sub RegisterOutlookInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Outlook
            ToolExecutorRegistry.Reset(host)
            RegisterSet(host, EnumerateCommonInternalToolNames())
            RegisterSet(host, OutlookAutoPilotToolNames)
        End Sub

        Public Sub RegisterWordInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Word
            ToolExecutorRegistry.Reset(host)
            RegisterSet(host, EnumerateWordSharedToolNames())
            RegisterSet(host, WordHostInternalToolNames)
        End Sub

        Public Sub RegisterResolvedInternalTools(host As ToolingHostKind, tools As IEnumerable(Of SharedLibrary.ModelConfig))
            If tools Is Nothing Then Return

            For Each tool As SharedLibrary.ModelConfig In tools
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For
                ToolExecutorRegistry.RegisterInternal(host, tool.ToolName.Trim())
            Next
        End Sub

        Public Function GetDeliverableCapableToolNames(host As ToolingHostKind) As IReadOnlyCollection(Of String)
            Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Select Case host
                Case ToolingHostKind.Word
                    For Each name As String In WordDeliverableToolNames
                        result.Add(name)
                    Next
                Case ToolingHostKind.Outlook
                    For Each name As String In OutlookDeliverableToolNames
                        result.Add(name)
                    Next
            End Select

            Return result.ToList().AsReadOnly()
        End Function

    End Module

End Namespace