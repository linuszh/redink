' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365ToolService.vb
' Purpose: Host-neutral implementation of the LLM tool family that exposes
'          Microsoft 365 search & retrieval to any add-in's tooling loop.
'
'   Tool family (chainable):
'      m365_search            — cross-source search (mail, OneDrive, SharePoint,
'                                Teams, calendar, OneNote, people)
'      m365_get_mail          — message body + recursive attachment text
'      m365_get_mail_thread   — entire conversation as one transcript
'      m365_get_file          — DriveItem download + plain-text extraction
'      m365_get_event         — calendar event details
'      m365_get_chat_thread   — 1:1/group chat OR channel post+replies
'      m365_get_onenote_page  — OneNote page (HTML stripped)
'
'   Public surface (host-agnostic):
'      M365ToolService.IsM365ToolName(name)               → Boolean
'      M365ToolService.GetTools(context, suffix)          → List(Of ModelConfig)
'      M365ToolService.ExecuteAsync(context, toolName,
'                                    arguments, log?, ct?) → Task(Of M365ToolExecutionResult)
'
'   Each Office add-in adds a small wrapper that translates its host-local
'   ToolCall / ToolResponse types to the dictionary + result struct below.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Result of a single M365 tool invocation. Hosts copy these fields onto
    ''' their own host-local ToolResponse type.
    ''' </summary>
    Public Class M365ToolExecutionResult
        Public Property Success As Boolean = True
        Public Property Response As String = ""
        Public Property ErrorMessage As String = ""
    End Class

    ''' <summary>
    ''' Static facade that defines, registers and executes the M365 LLM tools.
    ''' Depends only on <see cref="ISharedContext"/>, <see cref="ModelConfig"/>,
    ''' plain dictionaries and an optional logger callback.
    ''' </summary>
    Public Module M365ToolService

        ' ── Public tool name constants ─────────────────────────────────────
        Public Const ToolNamePrefix As String = "m365_"
        Public Const SearchToolName As String = "m365_search"
        Public Const GetMailToolName As String = "m365_get_mail"
        Public Const GetMailThreadToolName As String = "m365_get_mail_thread"
        Public Const GetFileToolName As String = "m365_get_file"
        Public Const GetEventToolName As String = "m365_get_event"
        Public Const GetChatThreadToolName As String = "m365_get_chat_thread"
        Public Const GetOneNotePageToolName As String = "m365_get_onenote_page"

        Public Const DefaultMaxChars As Integer = 200000
        Public Const HardMaxChars As Integer = 1000000

        Private Const DefaultSuffix As String = " (internal)"

        ' ════════════════════════════════════════════════════════════════════
        '  PUBLIC API
        ' ════════════════════════════════════════════════════════════════════

        ''' <summary>Returns True for any tool whose name starts with "m365_".</summary>
        Public Function IsM365ToolName(toolName As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(toolName) AndAlso
                   toolName.StartsWith(ToolNamePrefix, StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Returns the M365 tools as ready-to-register <see cref="ModelConfig"/> items.
        ''' Returns an empty list when <paramref name="context"/>.INI_M365ClientId is empty.
        ''' </summary>
        ''' <param name="suffix">Display-name suffix (each host passes its own InternalToolSuffix).</param>
        Public Function GetTools(context As ISharedContext, Optional suffix As String = DefaultSuffix) As List(Of ModelConfig)
            Dim tools As New List(Of ModelConfig)()
            If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.INI_M365ClientId) Then Return tools

            Dim sfx = If(suffix, DefaultSuffix)
            tools.Add(BuildSearchTool(sfx))
            tools.Add(BuildGetMailTool(sfx))
            tools.Add(BuildGetMailThreadTool(sfx))
            tools.Add(BuildGetFileTool(sfx))
            tools.Add(BuildGetEventTool(sfx))
            tools.Add(BuildGetChatThreadTool(sfx))
            tools.Add(BuildGetOneNotePageTool(sfx))
            Return tools
        End Function

        ''' <summary>
        ''' Executes any m365_* tool. Hosts call this from their tool dispatcher.
        ''' </summary>
        Public Async Function ExecuteAsync(context As ISharedContext,
                                           toolName As String,
                                           arguments As Dictionary(Of String, Object),
                                           Optional log As Action(Of String) = Nothing,
                                           Optional ct As CancellationToken = Nothing) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Try
                If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.INI_M365ClientId) Then
                    r.Success = False
                    r.ErrorMessage = "M365 integration is not configured (INI_M365ClientId is empty)."
                    Return r
                End If

                Dim safeArgs = If(arguments, New Dictionary(Of String, Object)())
                Dim safeLog As Action(Of String) = If(log, Sub(s) System.Diagnostics.Debug.WriteLine("[M365Tool] " & s))
                Dim name = If(toolName, "").Trim().ToLowerInvariant()

                Select Case name
                    Case SearchToolName
                        Return Await Execute_Search(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetMailToolName
                        Return Await Execute_GetMail(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetMailThreadToolName
                        Return Await Execute_GetMailThread(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetFileToolName
                        Return Await Execute_GetFile(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetEventToolName
                        Return Await Execute_GetEvent(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetChatThreadToolName
                        Return Await Execute_GetChatThread(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case GetOneNotePageToolName
                        Return Await Execute_GetOneNotePage(context, safeArgs, safeLog, ct).ConfigureAwait(False)
                    Case Else
                        r.Success = False
                        r.ErrorMessage = "Unknown M365 tool: " & toolName
                End Select
            Catch ex As Exception
                r.Success = False
                r.ErrorMessage = ex.Message
                System.Diagnostics.Debug.WriteLine("[M365Tool] Error: " & ex.ToString())
            End Try
            Return r
        End Function

        ' ════════════════════════════════════════════════════════════════════
        '  TOOL DEFINITIONS
        ' ════════════════════════════════════════════════════════════════════

        Private Function BuildSearchTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", SearchToolName),
                New JProperty("description",
                    "Searches the user's Microsoft 365 estate (their own emails, OneDrive/SharePoint files, " &
                    "SharePoint sites and list items, Teams chats, calendar events, OneNote pages, and people). " &
                    "Use this whenever the user refers to their own materials — e.g. 'my last email to Peter', " &
                    "'the contract I drafted with Acme', 'minutes from Tuesday's meeting'. Returns hits with " &
                    "stable ids you can pass to m365_get_mail, m365_get_mail_thread, m365_get_file, " &
                    "m365_get_event, m365_get_chat_thread or m365_get_onenote_page to fetch full content. " &
                    "PRIVACY: only the signed-in user's authenticated content is queried via Microsoft Graph."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("query",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description",
                                                    "KQL-compatible search expression. Plain words plus operators " &
                                                    "like from:""peter"" subject:""acme"" received>=2025-01-01."))),
                                New JProperty("sources",
                                    New JObject(New JProperty("type", "array"),
                                                New JProperty("items", New JObject(New JProperty("type", "string"))),
                                                New JProperty("description",
                                                    "Optional. Any combination of: 'mail', 'onedrive', 'sharepoint', " &
                                                    "'sharepoint_sites', 'sharepoint_listitems', 'teams', 'calendar', " &
                                                    "'onenote', 'people'. Convenience: 'all_files', 'all_sharepoint', 'all'. " &
                                                    "Defaults to 'all' when omitted."))),
                                New JProperty("max_per_source",
                                    New JObject(New JProperty("type", "integer"),
                                                New JProperty("description", "Default 25, server cap 500."))),
                                New JProperty("from_date",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Optional ISO date (YYYY-MM-DD)."))),
                                New JProperty("to_date",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Optional ISO date (YYYY-MM-DD)."))),
                                New JProperty("kql_extra",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Optional extra KQL appended verbatim."))))),
                        New JProperty("required", New JArray("query"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = SearchToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt =
                    "m365_search: Cross-source search of the signed-in user's Microsoft 365 content. " &
                    "USE THIS whenever the user refers to their own materials — do NOT use internet_search " &
                    "or web_content_retriever for the user's own content. Provide query (required). " &
                    "Optionally narrow with sources, max_per_source, from_date, to_date, kql_extra. " &
                    "Each hit has 'n', 'id', 'source', 'title', 'summary', 'date', 'web_url' and source-specific " &
                    "ids ('conversation_id' for mail, 'drive_id' for files, 'chat_id'/'team_id'/'channel_id' for Teams). " &
                    "Pass those ids to m365_get_mail, m365_get_mail_thread, m365_get_file, m365_get_event, " &
                    "m365_get_chat_thread or m365_get_onenote_page to fetch full content.",
                .ModelDescription = "M365 Search (mail/files/sites/Teams/calendar/notes)" & suffix,
                .Tool = True,
                .ToolPriority = 996,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetMailTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetMailToolName),
                New JProperty("description",
                    "Retrieves the full body of one of the user's emails (and, by default, recursively extracts " &
                    "the plain text of any attachments — Word/PDF/Excel/PowerPoint/text). Pass the 'id' returned " &
                    "by m365_search for a mail hit."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("message_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Graph message id from m365_search."))),
                                New JProperty("include_attachments",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default true."))),
                                New JProperty("ocr_pdf",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default false."))),
                                New JProperty("max_chars",
                                    New JObject(New JProperty("type", "integer"),
                                                New JProperty("description", $"Default {DefaultMaxChars}."))))),
                        New JProperty("required", New JArray("message_id"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetMailToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt =
                    "m365_get_mail: Returns headers, body and (by default) extracted attachment text. " &
                    "Provide message_id (required). Optional: include_attachments, ocr_pdf, max_chars.",
                .ModelDescription = "M365: Read e-mail" & suffix,
                .Tool = True,
                .ToolPriority = 995,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetMailThreadTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetMailThreadToolName),
                New JProperty("description",
                    "Retrieves an entire mail conversation flattened into one chronological transcript with " &
                    "attachments extracted. Pass 'conversation_id' from an m365_search mail hit."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("conversation_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Graph conversationId."))),
                                New JProperty("max_messages",
                                    New JObject(New JProperty("type", "integer"),
                                                New JProperty("description", "Default 200."))),
                                New JProperty("ascending",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default true (oldest first)."))),
                                New JProperty("include_attachments",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default true."))),
                                New JProperty("ocr_pdf",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default false."))),
                                New JProperty("max_chars",
                                    New JObject(New JProperty("type", "integer"),
                                                New JProperty("description", $"Default {DefaultMaxChars}."))))),
                        New JProperty("required", New JArray("conversation_id"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetMailThreadToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt =
                    "m365_get_mail_thread: Returns every message in a mail conversation as one transcript. " &
                    "Provide conversation_id. Optional: max_messages, ascending, include_attachments, ocr_pdf, max_chars.",
                .ModelDescription = "M365: Read mail thread" & suffix,
                .Tool = True,
                .ToolPriority = 994,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetFileTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetFileToolName),
                New JProperty("description",
                    "Downloads a OneDrive or SharePoint file (DriveItem) and returns its plain-text content. " &
                    "Supported: .docx, .xlsx, .pptx, .pdf, .eml, .rtf, .txt/.csv/.md/.json/.xml/.html. " &
                    "Pass 'drive_item_id' from m365_search (and 'drive_id' for files in shared drives)."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("drive_item_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Graph driveItem id."))),
                                New JProperty("drive_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Optional drive id for shared drives."))),
                                New JProperty("ocr_pdf",
                                    New JObject(New JProperty("type", "boolean"),
                                                New JProperty("description", "Default false."))),
                                New JProperty("max_chars",
                                    New JObject(New JProperty("type", "integer"),
                                                New JProperty("description", $"Default {DefaultMaxChars}."))))),
                        New JProperty("required", New JArray("drive_item_id"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetFileToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt =
                    "m365_get_file: Returns plain text for a OneDrive/SharePoint file. " &
                    "Provide drive_item_id (required). Optional: drive_id, ocr_pdf, max_chars.",
                .ModelDescription = "M365: Read OneDrive/SharePoint file" & suffix,
                .Tool = True,
                .ToolPriority = 993,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetEventTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetEventToolName),
                New JProperty("description",
                    "Returns details of a calendar event (subject, organiser, start/end, location, attendees, body). " &
                    "Pass 'event_id' from an m365_search calendar hit."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("event_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Graph event id."))))),
                        New JProperty("required", New JArray("event_id"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetEventToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt = "m365_get_event: Returns calendar event details. Provide event_id.",
                .ModelDescription = "M365: Read calendar event" & suffix,
                .Tool = True,
                .ToolPriority = 992,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetChatThreadTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetChatThreadToolName),
                New JProperty("description",
                    "Returns a Teams conversation as transcript. For 1:1/group chats pass 'chat_id'. " &
                    "For channel posts pass 'team_id' + 'channel_id' + 'root_message_id' (post + replies). " &
                    "All ids are returned by m365_search on Teams hits."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("chat_id", New JObject(New JProperty("type", "string"),
                                                                       New JProperty("description", "1:1 / group chat id."))),
                                New JProperty("team_id", New JObject(New JProperty("type", "string"),
                                                                       New JProperty("description", "Team id (channel threads)."))),
                                New JProperty("channel_id", New JObject(New JProperty("type", "string"),
                                                                          New JProperty("description", "Channel id (channel threads)."))),
                                New JProperty("root_message_id", New JObject(New JProperty("type", "string"),
                                                                                New JProperty("description", "Root channel message id."))),
                                New JProperty("max_messages", New JObject(New JProperty("type", "integer"),
                                                                            New JProperty("description", "Default 200."))),
                                New JProperty("ascending", New JObject(New JProperty("type", "boolean"),
                                                                         New JProperty("description", "Default true."))),
                                New JProperty("max_chars", New JObject(New JProperty("type", "integer"),
                                                                         New JProperty("description", $"Default {DefaultMaxChars}.")))))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetChatThreadToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt =
                    "m365_get_chat_thread: Returns a Teams conversation as a transcript. Provide either chat_id " &
                    "OR team_id + channel_id + root_message_id. Optional: max_messages, ascending, max_chars.",
                .ModelDescription = "M365: Read Teams thread" & suffix,
                .Tool = True,
                .ToolPriority = 991,
                .ToolErrorHandling = "skip"
            }
        End Function

        Private Function BuildGetOneNotePageTool(suffix As String) As ModelConfig
            Dim def As New JObject(
                New JProperty("name", GetOneNotePageToolName),
                New JProperty("description",
                    "Returns the plain-text content of a OneNote page. Pass 'page_id' from m365_search."),
                New JProperty("parameters",
                    New JObject(
                        New JProperty("type", "object"),
                        New JProperty("properties",
                            New JObject(
                                New JProperty("page_id",
                                    New JObject(New JProperty("type", "string"),
                                                New JProperty("description", "Graph OneNote page id."))))),
                        New JProperty("required", New JArray("page_id"))
                    ))
            )

            Return New ModelConfig() With {
                .ToolName = GetOneNotePageToolName,
                .ToolDefinition = def.ToString(Formatting.None),
                .ToolInstructionsPrompt = "m365_get_onenote_page: Returns OneNote page text. Provide page_id.",
                .ModelDescription = "M365: Read OneNote page" & suffix,
                .Tool = True,
                .ToolPriority = 990,
                .ToolErrorHandling = "skip"
            }
        End Function

        ' ════════════════════════════════════════════════════════════════════
        '  EXECUTORS
        ' ════════════════════════════════════════════════════════════════════

        Private Async Function Execute_Search(context As ISharedContext,
                                              args As Dictionary(Of String, Object),
                                              log As Action(Of String),
                                              ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim query = GetArgString(args, "query").Trim()
            If String.IsNullOrWhiteSpace(query) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: query" : Return r
            End If

            Dim sources = ParseSources(GetArg(args, "sources"))
            If sources = M365SearchSources.None Then sources = M365SearchSources.All

            Dim opts As New M365SearchOptions() With {
                .MaxPerSource = Math.Max(1, Math.Min(GetArgInt(args, "max_per_source", 25), 500)),
                .From = GetArgDate(args, "from_date"),
                .To = GetArgDate(args, "to_date"),
                .KqlExtra = GetArgString(args, "kql_extra"),
                .Parallel = True
            }

            log($"  Searching M365 (sources={sources}) for: {query}")

            Dim result As M365SearchResult
            Try
                result = Await M365Service.SearchAsync(context, query, sources, opts, Nothing, ct).ConfigureAwait(False)
            Catch ex As Exception
                r.Success = False : r.ErrorMessage = "M365 search failed: " & ex.Message : Return r
            End Try

            Dim hitsJson As New JArray()
            For i As Integer = 0 To result.Hits.Count - 1
                hitsJson.Add(HitToJson(result.Hits(i), i + 1))
            Next
            Dim errorsJson As New JObject()
            For Each kv In result.ErrorsBySource
                errorsJson(kv.Key.ToString()) = kv.Value
            Next

            Dim envelope As New JObject(
                New JProperty("query", query),
                New JProperty("requested_sources", sources.ToString()),
                New JProperty("total", result.Hits.Count),
                New JProperty("hits", hitsJson),
                New JProperty("errors", errorsJson)
            )
            r.Response = envelope.ToString(Formatting.None)
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetMail(context As ISharedContext,
                                               args As Dictionary(Of String, Object),
                                               log As Action(Of String),
                                               ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim msgId = GetArgString(args, "message_id").Trim()
            If String.IsNullOrEmpty(msgId) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: message_id" : Return r
            End If

            Dim opts As New M365TextOptions() With {
                .IncludeAttachments = GetArgBool(args, "include_attachments", True),
                .OcrPdf = GetArgBool(args, "ocr_pdf", False),
                .AskUserForOcr = False,
                .MaxChars = ClampMaxChars(GetArgInt(args, "max_chars", DefaultMaxChars))
            }
            log($"  Reading M365 mail: {msgId}")
            Dim text = Await M365Service.GetMessageAsTextAsync(context, msgId, opts, ct).ConfigureAwait(False)
            r.Response = WrapContent("MAIL", msgId, text)
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetMailThread(context As ISharedContext,
                                                     args As Dictionary(Of String, Object),
                                                     log As Action(Of String),
                                                     ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim convId = GetArgString(args, "conversation_id").Trim()
            If String.IsNullOrEmpty(convId) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: conversation_id" : Return r
            End If

            Dim textOpts As New M365TextOptions() With {
                .IncludeAttachments = GetArgBool(args, "include_attachments", True),
                .OcrPdf = GetArgBool(args, "ocr_pdf", False),
                .AskUserForOcr = False,
                .MaxChars = ClampMaxChars(GetArgInt(args, "max_chars", DefaultMaxChars))
            }
            Dim threadOpts As New M365ThreadOptions() With {
                .MaxMessages = GetArgInt(args, "max_messages", 200),
                .Ascending = GetArgBool(args, "ascending", True),
                .IncludeMailBody = True
            }
            log($"  Reading M365 mail thread: {convId}")
            Dim text = Await M365Service.GetMailThreadAsTextAsync(context, convId, textOpts, threadOpts, ct).ConfigureAwait(False)
            r.Response = WrapContent("MAIL_THREAD", convId, text)
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetFile(context As ISharedContext,
                                               args As Dictionary(Of String, Object),
                                               log As Action(Of String),
                                               ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim itemId = GetArgString(args, "drive_item_id").Trim()
            If String.IsNullOrEmpty(itemId) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: drive_item_id" : Return r
            End If
            Dim driveId = GetArgString(args, "drive_id").Trim()

            Dim opts As New M365TextOptions() With {
                .OcrPdf = GetArgBool(args, "ocr_pdf", False),
                .AskUserForOcr = False,
                .MaxChars = ClampMaxChars(GetArgInt(args, "max_chars", DefaultMaxChars))
            }
            log($"  Reading M365 file: {itemId}")
            Dim text = Await M365Service.GetDriveItemAsTextAsync(
                context, itemId, If(String.IsNullOrEmpty(driveId), Nothing, driveId), opts, ct).ConfigureAwait(False)
            r.Response = WrapContent("FILE", itemId, text)
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetEvent(context As ISharedContext,
                                                args As Dictionary(Of String, Object),
                                                log As Action(Of String),
                                                ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim id = GetArgString(args, "event_id").Trim()
            If String.IsNullOrEmpty(id) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: event_id" : Return r
            End If
            log($"  Reading M365 event: {id}")
            Dim text = Await M365Service.GetEventAsTextAsync(context, id, ct).ConfigureAwait(False)
            r.Response = WrapContent("EVENT", id, text)
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetChatThread(context As ISharedContext,
                                                     args As Dictionary(Of String, Object),
                                                     log As Action(Of String),
                                                     ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim chatId = GetArgString(args, "chat_id").Trim()
            Dim teamId = GetArgString(args, "team_id").Trim()
            Dim channelId = GetArgString(args, "channel_id").Trim()
            Dim rootMsgId = GetArgString(args, "root_message_id").Trim()

            Dim hasChat = Not String.IsNullOrEmpty(chatId)
            Dim hasChannel = Not String.IsNullOrEmpty(teamId) AndAlso
                             Not String.IsNullOrEmpty(channelId) AndAlso
                             Not String.IsNullOrEmpty(rootMsgId)
            If Not hasChat AndAlso Not hasChannel Then
                r.Success = False
                r.ErrorMessage = "Provide either chat_id, or all of team_id + channel_id + root_message_id."
                Return r
            End If

            Dim textOpts As New M365TextOptions() With {
                .MaxChars = ClampMaxChars(GetArgInt(args, "max_chars", DefaultMaxChars))
            }
            Dim threadOpts As New M365ThreadOptions() With {
                .MaxMessages = GetArgInt(args, "max_messages", 200),
                .Ascending = GetArgBool(args, "ascending", True)
            }

            If hasChat Then
                log($"  Reading Teams chat thread: {chatId}")
                Dim t = Await M365Service.GetChatThreadAsTextAsync(context, chatId, textOpts, threadOpts, ct).ConfigureAwait(False)
                r.Response = WrapContent("CHAT_THREAD", chatId, t)
            Else
                log($"  Reading Teams channel thread: {teamId}/{channelId}/{rootMsgId}")
                Dim t = Await M365Service.GetChannelThreadAsTextAsync(context, teamId, channelId, rootMsgId,
                                                                       textOpts, threadOpts, ct).ConfigureAwait(False)
                r.Response = WrapContent("CHANNEL_THREAD", $"{teamId}/{channelId}/{rootMsgId}", t)
            End If
            r.Success = True
            Return r
        End Function

        Private Async Function Execute_GetOneNotePage(context As ISharedContext,
                                                      args As Dictionary(Of String, Object),
                                                      log As Action(Of String),
                                                      ct As CancellationToken) As Task(Of M365ToolExecutionResult)
            Dim r As New M365ToolExecutionResult()
            Dim id = GetArgString(args, "page_id").Trim()
            If String.IsNullOrEmpty(id) Then
                r.Success = False : r.ErrorMessage = "Missing required parameter: page_id" : Return r
            End If
            log($"  Reading OneNote page: {id}")
            Dim text = Await M365Service.GetOneNotePageAsTextAsync(context, id, ct).ConfigureAwait(False)
            r.Response = WrapContent("ONENOTE_PAGE", id, text)
            r.Success = True
            Return r
        End Function

        ' ════════════════════════════════════════════════════════════════════
        '  HELPERS
        ' ════════════════════════════════════════════════════════════════════

        Private Function HitToJson(h As M365SearchHit, n As Integer) As JObject
            Dim o As New JObject()
            o("n") = n
            o("source") = h.Source.ToString().ToLowerInvariant()
            o("id") = If(h.Id, "")
            o("title") = If(h.Title, "")
            If Not String.IsNullOrWhiteSpace(h.Summary) Then o("summary") = h.Summary
            If Not String.IsNullOrEmpty(h.Author) Then o("author") = h.Author
            If h.LastModifiedUtc.HasValue Then o("date") = h.LastModifiedUtc.Value.ToString("u")
            If Not String.IsNullOrWhiteSpace(h.WebUrl) Then o("web_url") = h.WebUrl

            Dim resource As JObject = Nothing
            If h.RawJson IsNot Nothing Then resource = TryCast(h.RawJson("resource"), JObject)

            Select Case h.Source
                Case M365SearchSources.Mail
                    If resource IsNot Nothing Then
                        Dim conv = If(resource("conversationId")?.ToString(), "")
                        If Not String.IsNullOrEmpty(conv) Then o("conversation_id") = conv
                        Dim imid = If(resource("internetMessageId")?.ToString(), "")
                        If Not String.IsNullOrEmpty(imid) Then o("internet_message_id") = imid
                    End If
                Case M365SearchSources.OneDrive, M365SearchSources.SharePoint
                    If Not String.IsNullOrEmpty(h.ParentId) Then o("drive_id") = h.ParentId
                Case M365SearchSources.SharePointListItems
                    If Not String.IsNullOrEmpty(h.ParentId) Then o("site_id") = h.ParentId
                Case M365SearchSources.Teams
                    If Not String.IsNullOrEmpty(h.ParentId) Then o("chat_id_or_channel_id") = h.ParentId
                    If resource IsNot Nothing Then
                        Dim chanIdent = TryCast(resource("channelIdentity"), JObject)
                        If chanIdent IsNot Nothing Then
                            Dim teamId = If(chanIdent("teamId")?.ToString(), "")
                            Dim chanId = If(chanIdent("channelId")?.ToString(), "")
                            If Not String.IsNullOrEmpty(teamId) Then o("team_id") = teamId
                            If Not String.IsNullOrEmpty(chanId) Then o("channel_id") = chanId
                        End If
                        Dim cid = If(resource("chatId")?.ToString(), "")
                        If Not String.IsNullOrEmpty(cid) Then o("chat_id") = cid
                    End If
            End Select
            Return o
        End Function

        Private Function WrapContent(kind As String, id As String, r As M365TextResult) As String
            Dim sb As New StringBuilder()
            sb.AppendLine($"<{kind} id=""{id}"" title=""{EscapeAttr(If(r?.Title, ""))}"">")
            If r IsNot Nothing AndAlso Not String.IsNullOrEmpty(r.Text) Then sb.AppendLine(r.Text)
            If r IsNot Nothing AndAlso r.Truncated Then
                sb.AppendLine("[content was truncated to fit the requested max_chars budget]")
            End If
            If r IsNot Nothing AndAlso r.Errors IsNot Nothing AndAlso r.Errors.Count > 0 Then
                sb.AppendLine($"<NOTES>{String.Join(" | ", r.Errors)}</NOTES>")
            End If
            sb.AppendLine($"</{kind}>")
            Return sb.ToString()
        End Function

        Private Function EscapeAttr(s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            Return s.Replace("""", "'").Replace(vbCr, " ").Replace(vbLf, " ")
        End Function

        Private Function ClampMaxChars(value As Integer) As Integer
            If value <= 0 Then Return DefaultMaxChars
            Return Math.Min(value, HardMaxChars)
        End Function

        ' ── Argument parsers (case-insensitive, defensive) ─────────────────

        Private Function GetArg(args As Dictionary(Of String, Object), key As String) As Object
            If args Is Nothing Then Return Nothing
            For Each kv In args
                If String.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) Then Return kv.Value
            Next
            Return Nothing
        End Function

        Private Function GetArgString(args As Dictionary(Of String, Object), key As String, Optional dflt As String = "") As String
            Dim v = GetArg(args, key)
            If v Is Nothing Then Return dflt
            Return v.ToString()
        End Function

        Private Function GetArgInt(args As Dictionary(Of String, Object), key As String, Optional dflt As Integer = 0) As Integer
            Dim s = GetArgString(args, key, "")
            If String.IsNullOrWhiteSpace(s) Then Return dflt
            Dim i As Integer
            If Integer.TryParse(s, i) Then Return i
            Return dflt
        End Function

        Private Function GetArgBool(args As Dictionary(Of String, Object), key As String, Optional dflt As Boolean = False) As Boolean
            Dim s = GetArgString(args, key, "").Trim().ToLowerInvariant()
            If String.IsNullOrEmpty(s) Then Return dflt
            Return s = "true" OrElse s = "1" OrElse s = "yes" OrElse s = "y"
        End Function

        Private Function GetArgDate(args As Dictionary(Of String, Object), key As String) As Date?
            Dim s = GetArgString(args, key, "")
            If String.IsNullOrWhiteSpace(s) Then Return Nothing
            Dim d As Date
            If Date.TryParse(s, Globalization.CultureInfo.InvariantCulture,
                             Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal,
                             d) Then Return d
            Return Nothing
        End Function

        Private Function ParseSources(token As Object) As M365SearchSources
            Dim result As M365SearchSources = M365SearchSources.None
            If token Is Nothing Then Return result

            Dim names As New List(Of String)()
            If TypeOf token Is JArray Then
                For Each x In DirectCast(token, JArray)
                    names.Add(If(x?.ToString(), ""))
                Next
            ElseIf TypeOf token Is IEnumerable(Of Object) AndAlso Not TypeOf token Is String Then
                For Each x In DirectCast(token, IEnumerable(Of Object))
                    names.Add(If(x?.ToString(), ""))
                Next
            Else
                Dim s = token.ToString()
                For Each part In s.Split({","c, ";"c, " "c}, StringSplitOptions.RemoveEmptyEntries)
                    names.Add(part)
                Next
            End If

            For Each raw In names
                Dim n = raw.Trim().Trim(""""c, "'"c).ToLowerInvariant()
                Select Case n
                    Case "mail", "email", "emails", "messages", "outlook" : result = result Or M365SearchSources.Mail
                    Case "onedrive", "files", "personal_files" : result = result Or M365SearchSources.OneDrive
                    Case "sharepoint", "sp", "sp_files" : result = result Or M365SearchSources.SharePoint
                    Case "sharepoint_sites", "sites" : result = result Or M365SearchSources.SharePointSites
                    Case "sharepoint_listitems", "listitems", "list_items" : result = result Or M365SearchSources.SharePointListItems
                    Case "teams", "chat", "chats", "teams_messages" : result = result Or M365SearchSources.Teams
                    Case "calendar", "events", "meetings" : result = result Or M365SearchSources.Calendar
                    Case "onenote", "notes", "notebook" : result = result Or M365SearchSources.OneNote
                    Case "people", "contacts" : result = result Or M365SearchSources.People
                    Case "all_files" : result = result Or M365SearchSources.AllFiles
                    Case "all_sharepoint" : result = result Or M365SearchSources.AllSharePoint
                    Case "all", "*" : result = result Or M365SearchSources.All
                End Select
            Next
            Return result
        End Function

    End Module

End Namespace