' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedLibrary/Code/Agents/HostToolRegistration.vb
' Purpose: Idempotent per-host registration of internal tool names into
'          Agents.ToolExecutorRegistry. Called from LoadToolingServices on each
'          host. External (INI / MCP / HTTP) tools are registered by the hosts
'          themselves inside LoadToolingServices once they have been confirmed
'          to carry an APICall template.
'
' Design notes:
'   - Per-host. Outlook gets AutoPilot's "Advanced Tools" registered as Internal;
'     Word does NOT — they are not implemented there.
'   - Workspace tools and the shared Word file-edit tools are registered on both
'     hosts because their implementations live in SharedLibrary.
'   - This module deliberately uses hardcoded name literals (rather than reflecting
'     into each host) so the registry is deterministic and reviewable. If a host
'     adds a new internal tool, add its literal name here in the appropriate list.
' =============================================================================

Option Explicit On
Option Strict On

Namespace Agents

    Public Module HostToolRegistration

        ' --- Outlook AutoPilot "Advanced Tools" (host-internal, Outlook only) -----
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

        ' --- Built-in internal tools (host implementations) -----------------------
        Private ReadOnly OutlookBuiltInInternalToolNames As String() = New String() {
            "retrieve_web_content",
            "download_web_files",
            "internet_search",
            "knowledge_search"
        }

        Private ReadOnly WordBuiltInInternalToolNames As String() = New String() {
            "retrieve_web_content",
            "internet_search",
            "knowledge_search"
        }

        ' --- Workspace tools (shared; available on both hosts) --------------------
        ' These mirror the static set used by SharedLibrary.Agents.WorkspaceTools.
        ' If a new workspace tool is added there, add its literal here too.
        Private ReadOnly WorkspaceToolNames As String() = New String() {
            "workspace_list",
            "workspace_read",
            "workspace_read_many",
            "workspace_write",
            "workspace_delete",
            "workspace_move",
            "workspace_extract_text",
            "workspace_extract_text_many",
            "agent_workspace_read",
            "agent_workspace_write",
            "agent_workspace_list"
        }

        ' --- Word file-edit (Word host only) --------------------------------------
        Private ReadOnly WordHostInternalToolNames As String() = New String() {
            "word_doc_read",
            "word_doc_edit",
            "word_doc_create",
            "word_doc_export_pdf"
        }

        Public Sub RegisterAll(host As ToolingHostKind)
            Select Case host
                Case ToolingHostKind.Outlook
                    RegisterOutlookInternals()
                Case ToolingHostKind.Word
                    RegisterWordInternals()
                Case Else
                    ' No-op for unknown hosts.
            End Select
        End Sub

        ''' <summary>Registers Outlook's host-internal tools. Idempotent.</summary>
        Public Sub RegisterOutlookInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Outlook
            ToolExecutorRegistry.Reset(host)
            For Each name As String In OutlookAutoPilotToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
            For Each name As String In OutlookBuiltInInternalToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
            For Each name As String In WorkspaceToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
        End Sub

        ''' <summary>Registers Word's host-internal tools. Idempotent.</summary>
        Public Sub RegisterWordInternals()
            Dim host As ToolingHostKind = ToolingHostKind.Word
            ToolExecutorRegistry.Reset(host)
            For Each name As String In WordBuiltInInternalToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
            For Each name As String In WorkspaceToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
            For Each name As String In WordHostInternalToolNames
                ToolExecutorRegistry.RegisterInternal(host, name)
            Next
        End Sub

    End Module

End Namespace