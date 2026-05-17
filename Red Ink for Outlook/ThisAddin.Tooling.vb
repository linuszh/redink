' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddin.Tooling.vb
' Purpose: Implements a model-agnostic "tooling loop" for LLM tool/function calling, including
'          tool selection, tool call detection/extraction, per-tool execution, and response
'          injection back into subsequent LLM iterations.
'
' Architecture:
'  - Tooling execution loop (`ExecuteToolingLoop`):
'      - Builds system prompt augmentation via `BuildToolInstructionsPrompt`.
'      - Injects model-specific tool definitions into `INI_APICall_ToolInstructions_2`.
'      - Calls `LLM(...)` iteratively until no tool calls are detected or `MaxIterations` is reached.
'      - When tool calls are present:
'          - Detects tool calls using `ContainsToolCalls` (regex).
'          - Extracts tool calls using `ExtractToolCalls` (JSON + extraction map).
'          - Executes each tool via `ExecuteToolCall`, collecting `ToolResponse` objects.
'          - Builds the next-iteration response payload via `BuildToolResponsesForModel` and assigns it to
'            `INI_APICall_ToolResponses_2`.
'  - Tool execution:
'      - Internal tool: `retrieve_web_content` retrieves web content for caller-provided URLs.
'      - External tools: `ExecuteExternalTool` applies the selected tool `ModelConfig`, prepares the
'        API call payload, forces JSON response mode, invokes `LLM(...)`, and restores prior settings.
'      - Tool errors are handled according to `ModelConfig.ToolErrorHandling` (abort/retry/skip).
'  - Tool selection and persistence:
'      - Loads tools from `INI_SpecialServicePath` with `LoadToolingServices`.
'      - Adds the internal web retrieval tool via `GetInternalWebTool`.
'      - Persists selections through `My.Settings.SelectedToolNames` and restores them with
'        `LoadPersistedToolSelection`.
'  - Diagnostics:
'      - `ToolingFileLogger` writes a single per-run log file when `INI_APIDebug` is enabled.
'      - Optional UI logging uses a `LogWindow` instance when enabled.
'
' External Dependencies:
'  - SharedLibrary.SharedMethods: `LLM`, `InterpolateAtRuntime`, `LoadAlternativeModels`, `GetCurrentConfig`,
'    `ApplyModelConfig`, `RestoreDefaults`, `ShowCustomMessageBox`, `ShowCustomYesNoBox`.
'  - Newtonsoft.Json: used for parsing/formatting tool calls and tool responses.
'  - WebView2: used for JavaScript-capable web retrieval when enabled.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides tooling support helpers for model-agnostic tool/function calling in LLM interactions.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>User-friendly name for the tooling feature.</summary>
    Public Const ToolFriendlyName As String = "Sources"

    ''' <summary>Auto-close delay for tooling log window.</summary>
    Public Shared Property ToolingLog_AutoCloseDefaultSeconds As Integer = 30

    ''' <summary>Selected tool names for persistence.</summary>
    Public Shared Property SelectedToolNames As List(Of String) = New List(Of String)()

    ''' <summary>
    ''' Enables the WebView2-based implementation for internal web content retrieval.
    ''' </summary>
    Private Const UseWebView2 As Boolean = True

    ''' <summary>
    ''' Internal tool name used in model tool calls for web content retrieval.
    ''' </summary>
    Private Const InternalWebToolName As String = "retrieve_web_content"

    ''' <summary>
    ''' Human-readable tool instructions shown to the model describing how to use the internal web tool.
    ''' </summary>
    Private Const InternalWebToolInstructionsPrompt As String =
        "retrieve_web_content: Retrieves text content from one or more URLs. Use this to fetch information from websites when needed. " &
        "SHAREPOINT/ONEDRIVE LIMITATION: This tool CANNOT access SharePoint, OneDrive, Microsoft Teams, or any other " &
        "authenticated cloud storage URLs. URLs containing 'sharepoint.com', 'onedrive.com', '1drv.ms', " &
        "'teams.microsoft.com', or ':f:/' point to resources that require authentication and will NOT return " &
        "useful content. UNC paths (e.g. \\server\share\file.doc) that resolve to SharePoint will also fail. " &
        "If the user asks you to retrieve content from such a link, do NOT call this tool. Instead, explain " &
        "that you cannot remotely log into authenticated cloud storage and ask the user to download the file(s) " &
        "and provide them as direct attachments."

    ''' <summary>
    ''' Canonical tool definition JSON used for the internal web tool.
    ''' </summary>
    Private Const InternalWebToolDefinition As String =
        "{""name"":""retrieve_web_content"",""description"":""Retrieves the text content from one or more web URLs. " &
        "IMPORTANT: Cannot access SharePoint, OneDrive, Teams, or other authenticated cloud storage URLs " &
        "(sharepoint.com, onedrive.com, 1drv.ms, teams.microsoft.com, :f:/). " &
        "Do NOT call this tool for such URLs — ask the user to download and attach the file(s) instead."",""parameters"":{""type"":""object"",""properties"":{""urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Array of URLs to retrieve content from""}},""required"":[""urls""]}}"

    ' Internet Search Tooling (available only when INI_ISearch is enabled and INI_ISearch_URL is configured)

    ''' <summary>
    ''' Internal tool name used in model tool calls for internet search.
    ''' </summary>
    Private Const InternalSearchToolName As String = "internet_search"

    ''' <summary>
    ''' Canonical tool definition JSON used for the internal search tool.
    ''' </summary>
    Private Const InternalSearchToolDefinition As String =
        "{""name"":""internet_search""," &
        """description"":""Searches the internet via the configured search engine, retrieves the top result pages, and returns their readable text content. Use this when you need up-to-date or factual information you are not confident about. PRIVACY: The query is sent to an external search engine. Never include personal data, confidential information, private names, case details, contract terms, internal identifiers, or any non-public information in the query. Only public figures, public institutions, published legislation, and other clearly public information may appear. If a useful query cannot be formed without non-public data, do not call this tool.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""The search query. MUST NOT contain personal data, confidential details, or any non-public information. Use only generic, anonymized, or publicly known terms.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of search result pages to retrieve (default: 4, server-capped).""}," &
        """max_depth"":{""type"":""integer"",""description"":""Maximum crawl depth per result page. 0 = top-level only (default: 0, server-capped).""}},""required"":[""query""]}}"

    ''' <summary>
    ''' Human-readable tool instructions shown to the model describing how to use the internal search tool.
    ''' </summary>
    Private Const InternalSearchToolInstructionsPrompt As String =
        "internet_search: Searches the internet and returns readable text from the top result pages. " &
        "Call this tool when you need current or factual information you are not confident about. " &
        "Provide query (required string). Optionally provide max_results (integer, default 4) and max_depth (integer, default 0). " &
        "Return value includes the search query used, the URLs visited, and the page content for each qualifying result. " &
        "IMPORTANT PRIVACY CONSTRAINT: The search query is sent to an external search engine. " &
        "You MUST NOT include any personal data, confidential information, private names, " &
        "case details, contract terms, internal identifiers, email addresses, phone numbers, " &
        "account numbers, or any other non-public information in the query. " &
        "Only well-known public figures, public institutions, published legislation, " &
        "publicly available case law references, and other clearly public information may appear in queries. " &
        "If you cannot formulate a useful query without disclosing non-public information, " &
        "do NOT call this tool — instead respond based on your existing knowledge and state your uncertainty."

    ' Knowledge Store Tooling (available only when KnowledgeStorePath or KnowledgeStorePathLocal is configured)

    Private Const InternalKnowledgeToolName As String = "knowledge_search"

    Private Const InternalKnowledgeToolDefinition As String =
        "{""name"":""knowledge_search""," &
        """description"":""Searches the user's local knowledge store (a curated collection of documents such as contracts, policies, legal briefs, " &
        "manuals, and reference material) and returns the most relevant document content. Use this tool when the user's question " &
        "relates to their own documents, internal policies, past work, or reference material that would not be found on the public internet. " &
        "Do NOT use this tool for general knowledge questions or publicly available information — use your training data or internet_search instead.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""A natural language search query describing what information is needed from the knowledge store. " &
        "Supports optional prefixes: 'tag:tagname' to filter by tag, 'store:storename' to restrict to a specific store, " &
        "or both 'tag:tagname store:storename'. Without prefixes, all stores are searched by keyword relevance.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of documents to retrieve (default: 5, max: 10).""}},""required"":[""query""]}}"

    Private Const InternalKnowledgeToolInstructionsPrompt As String =
        "knowledge_search: Searches the user's local knowledge store — a curated library of the user's own documents " &
        "(contracts, policies, briefs, manuals, templates, reference material, etc.). " &
        "Call this tool when the user's question relates to their own documents, internal policies, past work, or reference material. " &
        "Provide query (required string) describing the information needed. Optionally provide max_results (integer, default 5). " &
        "The query supports optional prefixes: 'tag:tagname' filters by document tag, 'store:storename' restricts to a specific knowledge store. " &
        "Return value is the text content of the most relevant documents, each tagged with the document name and store. " &
        "IMPORTANT: Do NOT use this tool for general knowledge or publicly available information — only for the user's own document library. " &
        "When citing information from the results, mention the document name so the user can locate the source."

    ''' <summary>Minimum characters for a search hit to be considered relevant.</summary>
    Private Const ISearch_MinChars As Integer = 500

    ''' <summary>Maximum characters per search result page (applied as a cap even with WebView2).</summary>
    Private Const ISearch_MaxChars As Integer = 4000

    ''' <summary>
    ''' Suffix appended to internal tool descriptions when shown to the user.
    ''' </summary>
    Private Const InternalToolSuffix As String = " (built-in)"

    ''' <summary>
    ''' Maximum tool loop iterations used by prompt-building helpers.
    ''' </summary>
    Public MaxToolIterations As Integer = 10

    ''' <summary>
    ''' Holds the active ToolExecutionContext during an ExecuteToolingLoop run.
    ''' Used by ApDashboardLog to route messages to the Chat Agent's LogWindow.
    ''' </summary>
    Private _activeToolingContext As ToolExecutionContext = Nothing

#Region "WebView2 Web Content Retrieval"

    ''' <summary>
    ''' Retrieves website content using WebView2 (Edge) to include JavaScript-rendered content.
    ''' A dedicated STA thread is created to host WebView2 and run a minimal message loop.
    ''' </summary>
    ''' <param name="baseUrl">Absolute HTTP/HTTPS URL to retrieve.</param>
    ''' <param name="maxChars">Maximum characters to return. Values less than or equal to 0 return full content.</param>
    ''' <returns>A task producing the extracted plain text content (possibly truncated) or an empty string.</returns>

    Private Function RetrieveWebsiteContent_WebView2(baseUrl As String, Optional maxChars As Integer = 0) As Task(Of String)
        Dim tcs As New TaskCompletionSource(Of String)()

        ' Create a new STA thread for WebView2 (it requires its own message loop)
        Dim thread As New System.Threading.Thread(
            Sub()
                Dim result As String = ""
                Dim form As Form = Nothing
                Dim webView As Microsoft.Web.WebView2.WinForms.WebView2 = Nothing
                Dim userDataFolder As String = Nothing

                Try
                    Debug.WriteLine($"[WebView2] Fetching: {baseUrl} (maxChars: {If(maxChars <= 0, "unlimited", maxChars.ToString())})")

                    ' Create a temporary user data folder
                    Dim uniqueID As String = Guid.NewGuid().ToString()
                    userDataFolder = Path.Combine(Path.GetTempPath(), "RedInkWebView2_" & uniqueID)
                    Directory.CreateDirectory(userDataFolder)

                    ' Create the form - larger size to trigger more content loading
                    form = New Form() With {
                        .Width = 1920,
                        .Height = 4000,
                        .ShowInTaskbar = False,
                        .FormBorderStyle = FormBorderStyle.None,
                        .StartPosition = FormStartPosition.Manual,
                        .Location = New System.Drawing.Point(-5000, -5000),
                        .Opacity = 0
                    }

                    webView = New Microsoft.Web.WebView2.WinForms.WebView2() With {
                        .Dock = DockStyle.Fill
                    }

                    form.Controls.Add(webView)

                    Dim navigationCompleted As Boolean = False
                    Dim navigationSuccess As Boolean = False
                    Dim contentExtracted As Boolean = False

                    ' Set up event handlers
                    AddHandler webView.CoreWebView2InitializationCompleted,
                            Sub(s, e)
                                If e.IsSuccess Then
                                    Debug.WriteLine("[WebView2] CoreWebView2 initialized")

                                    ' 1. Validate Input URL before navigating
                                    If Not IsSafeWebUrl(baseUrl) Then
                                        Debug.WriteLine($"[WebView2] Blocked unsafe URL: {baseUrl}")
                                        navigationCompleted = True
                                        Return
                                    End If

                                    ' 2. Lockdown Settings
                                    With webView.CoreWebView2.Settings
                                        .AreDefaultScriptDialogsEnabled = False
                                        .AreDefaultContextMenusEnabled = False
                                        .AreDevToolsEnabled = False
                                        .IsStatusBarEnabled = False
                                        .IsScriptEnabled = True
                                        .IsBuiltInErrorPageEnabled = False
                                        .IsWebMessageEnabled = False
                                    End With

                                    ' 3. Block New Windows / Popups
                                    AddHandler webView.CoreWebView2.NewWindowRequested,
                                        Sub(sender, args)
                                            args.Handled = True
                                        End Sub

                                    ' 4. Block Permission Requests
                                    AddHandler webView.CoreWebView2.PermissionRequested,
                                        Sub(sender, args)
                                            args.State = CoreWebView2PermissionState.Deny
                                        End Sub

                                    ' 5. Block Navigation to non-http schemes
                                    AddHandler webView.CoreWebView2.NavigationStarting,
                                        Sub(sender, args)
                                            Dim uriStart As Uri = Nothing
                                            If Uri.TryCreate(args.Uri, UriKind.Absolute, uriStart) Then
                                                If uriStart.Scheme <> Uri.UriSchemeHttp AndAlso uriStart.Scheme <> Uri.UriSchemeHttps Then
                                                    args.Cancel = True
                                                End If
                                            End If
                                        End Sub

                                    Debug.WriteLine("[WebView2] Navigating...")
                                    webView.CoreWebView2.Navigate(baseUrl)
                                Else
                                    Debug.WriteLine($"[WebView2] Initialization failed: {e.InitializationException?.Message}")
                                    navigationCompleted = True
                                End If
                            End Sub

                    AddHandler webView.NavigationCompleted,
                        Sub(s, e)
                            navigationSuccess = e.IsSuccess
                            Debug.WriteLine($"[WebView2] Navigation completed. Success: {e.IsSuccess}, Status: {e.WebErrorStatus}")

                            If e.IsSuccess Then
                                ' Use a timer to wait for JS rendering - 5 seconds
                                Dim timer As New System.Windows.Forms.Timer() With {.Interval = 5000}
                                AddHandler timer.Tick,
                                    Sub(ts, te)
                                        timer.Stop()
                                        timer.Dispose()

                                        Try
                                            ' Scroll to trigger lazy loading
                                            Dim scrollScript As String = "
                                                (async function() {
                                                    var totalHeight = document.body.scrollHeight;
                                                    var viewportHeight = window.innerHeight || 1000;
                                                    var currentPosition = 0;

                                                    while (currentPosition < totalHeight) {
                                                        window.scrollTo(0, currentPosition);
                                                        await new Promise(r => setTimeout(r, 200));
                                                        currentPosition += viewportHeight;
                                                        totalHeight = document.body.scrollHeight;
                                                    }

                                                    window.scrollTo(0, 0);
                                                    await new Promise(r => setTimeout(r, 300));
                                                    return 'done';
                                                })();
                                            "
                                            webView.CoreWebView2.ExecuteScriptAsync(scrollScript).ContinueWith(
                                                Sub(scrollTask)
                                                    ' Wait after scrolling
                                                    System.Threading.Thread.Sleep(2000)

                                                    form.BeginInvoke(
                                                        Sub()
                                                            Try

                                                                ' Extract body text with inline hyperlinks as Markdown [text](url)
                                                                Dim extractScript As String = "
                                                (function() {
                                                    // Remove noise elements
                                                    var toRemove = document.querySelectorAll('script, style, noscript, nav, footer, header');
                                                    toRemove.forEach(function(el) { try { el.remove(); } catch(e) {} });

                                                    // Recursively extract text, inlining <a> hrefs as Markdown links
                                                    function walk(node) {
                                                        if (!node) return '';
                                                        if (node.nodeType === 3) return node.textContent || '';
                                                        if (node.nodeType !== 1) return '';

                                                        var tag = node.tagName ? node.tagName.toUpperCase() : '';

                                                        // Skip hidden elements
                                                        var style = window.getComputedStyle(node);
                                                        if (style && (style.display === 'none' || style.visibility === 'hidden')) return '';

                                                        // Collect child text first
                                                        var parts = [];
                                                        for (var i = 0; i < node.childNodes.length; i++) {
                                                            parts.push(walk(node.childNodes[i]));
                                                        }
                                                        var inner = parts.join('');

                                                        // Anchor: emit Markdown link if href is meaningful
                                                        if (tag === 'A') {
                                                            var href = (node.getAttribute('href') || '').trim();
                                                            var text = inner.trim();
                                                            if (href && text && href !== '#' && !href.startsWith('javascript:')) {
                                                                // Resolve relative URLs to absolute
                                                                try { href = new URL(href, document.baseURI).href; } catch(e) {}
                                                                // Avoid duplicating link text that IS the URL
                                                                if (text === href || text === decodeURIComponent(href)) return text;
                                                                return '[' + text + '](' + href + ')';
                                                            }
                                                            return text || '';
                                                        }

                                                        // Block-level elements get newlines
                                                        if (/^(DIV|P|BR|H[1-6]|LI|TR|BLOCKQUOTE|SECTION|ARTICLE|ASIDE|MAIN|DT|DD|FIGCAPTION|PRE)$/.test(tag)) {
                                                            if (tag === 'BR') return '\n';
                                                            return '\n' + inner + '\n';
                                                        }

                                                        return inner;
                                                    }

                                                    var text = document.body ? walk(document.body) : '';

                                                    // Clean up whitespace
                                                    text = text.replace(/\n{3,}/g, '\n\n').replace(/[ \t]+/g, ' ').trim();

                                                    return text;
                                                })();
                                                "

                                                                webView.CoreWebView2.ExecuteScriptAsync(extractScript).ContinueWith(
                                                                    Sub(t)
                                                                        form.BeginInvoke(
                                                                            Sub()
                                                                                Try
                                                                                    If t.IsCompleted AndAlso Not t.IsFaulted Then
                                                                                        result = UnescapeJsonString(t.Result)
                                                                                        Debug.WriteLine($"[WebView2] Extracted {result.Length} chars (full content)")
                                                                                    End If
                                                                                Catch ex As Exception
                                                                                    Debug.WriteLine($"[WebView2] Extract error: {ex.Message}")
                                                                                End Try
                                                                                contentExtracted = True
                                                                            End Sub)
                                                                    End Sub)
                                                            Catch ex As Exception
                                                                Debug.WriteLine($"[WebView2] Script error: {ex.Message}")
                                                                contentExtracted = True
                                                            End Try
                                                        End Sub)
                                                End Sub)
                                        Catch ex As Exception
                                            Debug.WriteLine($"[WebView2] Timer error: {ex.Message}")
                                            contentExtracted = True
                                        End Try
                                    End Sub
                                timer.Start()
                            Else
                                navigationCompleted = True
                            End If
                        End Sub

                    ' Show form and start async initialization
                    form.Show()

                    ' Initialize WebView2 asynchronously
                    Dim env = CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)
                    env.ContinueWith(
                            Sub(t)
                                form.BeginInvoke(
                                    Sub()
                                        If t.IsCompleted AndAlso Not t.IsFaulted Then
                                            webView.EnsureCoreWebView2Async(t.Result)
                                            ' Don't do anything else here - wait for CoreWebView2InitializationCompleted
                                        Else
                                            Debug.WriteLine($"[WebView2] Environment creation failed")
                                            navigationCompleted = True
                                        End If
                                    End Sub)
                            End Sub)

                    ' Run message loop with timeout - 90 seconds
                    Dim startTime As DateTime = DateTime.Now
                    Dim timeout As TimeSpan = TimeSpan.FromSeconds(90)

                    While Not contentExtracted AndAlso Not navigationCompleted AndAlso (DateTime.Now - startTime) < timeout
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(50)
                    End While

                    Debug.WriteLine($"[WebView2] Loop ended. Content: {result.Length} chars, elapsed: {(DateTime.Now - startTime).TotalSeconds:F1}s")

                Catch ex As Exception
                    Debug.WriteLine($"[WebView2] Error: {ex.Message}")
                Finally
                    Try
                        If userDataFolder IsNot Nothing Then Directory.Delete(userDataFolder, True)
                    Catch
                    End Try
                    Try
                        webView?.Dispose()
                    Catch
                    End Try
                    Try
                        form?.Close()
                        form?.Dispose()
                    Catch
                    End Try
                End Try

                ' Apply character limit only if explicitly requested (maxChars > 0)
                Dim finalResult As String = result.Trim()
                If maxChars > 0 AndAlso finalResult.Length > maxChars Then
                    ' Try to cut at a sentence boundary
                    Dim cutPoint As Integer = maxChars
                    Dim lastPeriod As Integer = finalResult.LastIndexOf("."c, maxChars - 1)
                    Dim lastNewline As Integer = finalResult.LastIndexOf(vbLf, maxChars - 1)

                    If lastPeriod > maxChars * 0.8 Then
                        cutPoint = lastPeriod + 1
                    ElseIf lastNewline > maxChars * 0.8 Then
                        cutPoint = lastNewline
                    End If

                    finalResult = finalResult.Substring(0, cutPoint).Trim()
                    Debug.WriteLine($"[WebView2] Trimmed to {finalResult.Length} chars (limit was {maxChars})")
                End If

                tcs.TrySetResult(finalResult)
            End Sub)

        thread.SetApartmentState(System.Threading.ApartmentState.STA)
        thread.Start()

        Return tcs.Task
    End Function

    ''' <summary>
    ''' Determines whether a URL is allowed to be retrieved by the internal web tool.
    ''' Only absolute HTTP/HTTPS URLs are permitted; loopback/localhost is denied.
    ''' </summary>
    ''' <param name="url">Candidate URL.</param>
    ''' <returns><c>True</c> if the URL is considered safe; otherwise <c>False</c>.</returns>

    Private Function IsSafeWebUrl(url As String) As Boolean
        Try
            Dim uriResult As Uri = Nothing
            If Not Uri.TryCreate(url, UriKind.Absolute, uriResult) Then Return False

            ' 1. Only allow HTTP and HTTPS
            If uriResult.Scheme <> Uri.UriSchemeHttp AndAlso uriResult.Scheme <> Uri.UriSchemeHttps Then
                Return False
            End If

            ' 2. Basic SSRF check: Block localhost / loopback
            If uriResult.IsLoopback Then Return False
            If uriResult.Host.ToLower().Equals("localhost") Then Return False

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Unescapes the JSON-encoded string returned by WebView2 <c>ExecuteScriptAsync</c>.
    ''' </summary>
    ''' <param name="jsonString">JSON string value returned by script execution.</param>
    ''' <returns>Unescaped string content.</returns>
    Private Function UnescapeJsonString(jsonString As String) As String
        If String.IsNullOrEmpty(jsonString) Then Return ""

        ' Remove surrounding quotes
        If jsonString.StartsWith("""") AndAlso jsonString.EndsWith("""") Then
            jsonString = jsonString.Substring(1, jsonString.Length - 2)
        End If

        ' Handle null
        If jsonString = "null" Then Return ""

        ' Unescape common sequences
        Try
            jsonString = jsonString.Replace("\n", vbLf)
            jsonString = jsonString.Replace("\r", vbCr)
            jsonString = jsonString.Replace("\t", vbTab)
            jsonString = jsonString.Replace("\\", "\")
            jsonString = jsonString.Replace("\""", """")

            ' Handle unicode escapes like \u0027
            Dim unicodePattern As String = "\\u([0-9A-Fa-f]{4})"
            jsonString = Regex.Replace(jsonString, unicodePattern,
                Function(m) ChrW(System.Convert.ToInt32(m.Groups(1).Value, 16)).ToString())

        Catch ex As Exception
            Debug.WriteLine($"[WebView2] Unescape error: {ex.Message}")
        End Try

        Return jsonString
    End Function

    ''' <summary>
    ''' Retrieves website content using an HTTP client (non-WebView2 fallback).
    ''' </summary>
    ''' <param name="url">Absolute URL to retrieve.</param>
    ''' <param name="maxDepth">Unused by this implementation; retained for caller compatibility.</param>
    ''' <param name="client">HTTP client used for retrieval.</param>
    ''' <returns>Response body as a string, or an error string on failure.</returns>
    Private Async Function RetrieveWebsiteContent(url As String, maxDepth As Integer, client As System.Net.Http.HttpClient) As Task(Of String)
        Try
            Return Await client.GetStringAsync(url)
        Catch ex As Exception
            Return $"Error retrieving {url}: {ex.Message}"
        End Try
    End Function

#End Region




#Region "Tooling File Logger (Reduced, Single File)"

    ''' <summary>
    ''' Reduced file-based logger for tooling operations.
    ''' - Single file per run (overwrites).
    ''' - Writes: LogWindow steps, warnings, errors, and pre-LLM call snapshots.
    ''' - Writes full raw LLM/tool responses with two empty lines before and after.
    ''' </summary>
    Public Class ToolingFileLogger

        ''' <summary>Absolute filesystem path to the current tooling log file (if enabled).</summary>
        Private Shared _logPath As String = Nothing

        ''' <summary>Whether file logging is enabled (controlled by <c>INI_APIDebug</c> in <see cref="StartSession"/>).</summary>
        Private Shared _isEnabled As Boolean = False

        ''' <summary>Whether the session header has already been written in this run.</summary>
        Private Shared _started As Boolean = False

        ''' <summary>Synchronizes file writes to avoid interleaving.</summary>
        Private Shared ReadOnly _lock As New Object()

        ''' <summary>Stable log filename used for the tooling session log (overwritten each run).</summary>
        Private Shared ReadOnly StableLogFileName As String = $"{AN5}_Tooling_Log.txt"

        ''' <summary>
        ''' Starts a tooling log session and writes the log header.
        ''' Logging is enabled only when <c>INI_APIDebug</c> is <c>True</c>.
        ''' </summary>
        Public Shared Sub StartSession()
            _isEnabled = INI_APIDebug
            If Not _isEnabled Then Return

            If _started AndAlso Not String.IsNullOrWhiteSpace(_logPath) Then Return

            Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            _logPath = Path.Combine(desktopPath, StableLogFileName)

            Try
                WriteHeader()
                _started = True
            Catch ex As Exception
                _isEnabled = False
                _started = False
                _logPath = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' Writes the session header and overwrites any previous log file with the same name.
        ''' </summary>
        Private Shared Sub WriteHeader()
            Dim header As New StringBuilder()
            header.AppendLine("=" & New String("="c, 78))
            header.AppendLine($"{AN} - Tooling Log")
            header.AppendLine($"Started: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}")
            header.AppendLine($"Version: {Version}")
            header.AppendLine($"File: {StableLogFileName} (overwritten each run)")
            header.AppendLine("=" & New String("="c, 78))
            header.AppendLine()

            File.WriteAllText(_logPath, header.ToString())
        End Sub

        ''' <summary>
        ''' Logs a step message to the tooling log (if enabled).
        ''' </summary>
        ''' <param name="message">Message text.</param>
        Public Shared Sub LogStep(message As String)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return
            WriteLine("STEP", message)
        End Sub

        ''' <summary>
        ''' Logs a warning message and optional details/exception to the tooling log (if enabled).
        ''' </summary>
        ''' <param name="message">Primary message text.</param>
        ''' <param name="details">Optional detail text written as a separate log line.</param>
        ''' <param name="ex">Optional exception whose type/message/stack is written.</param>
        Public Shared Sub LogWarn(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return
            WriteLine("WARN", message)
            If Not String.IsNullOrWhiteSpace(details) Then
                WriteLine("WARN", $"Details: {details}")
            End If
            If ex IsNot Nothing Then
                WriteException("WARN", ex)
            End If
        End Sub

        ''' <summary>
        ''' Logs an error message and optional details/exception to the tooling log (if enabled).
        ''' </summary>
        ''' <param name="message">Primary message text.</param>
        ''' <param name="details">Optional detail text written as a separate log line.</param>
        ''' <param name="ex">Optional exception whose type/message/stack is written.</param>
        Public Shared Sub LogError(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            WriteLine("ERR", message)
            If Not String.IsNullOrWhiteSpace(details) Then
                WriteLine("ERR", $"Details: {details}")
            End If
            If ex IsNot Nothing Then
                WriteException("ERR", ex)
            End If
        End Sub

        ''' <summary>
        ''' Writes exception details to the log under the specified category.
        ''' </summary>
        ''' <param name="category">Log category identifier (e.g., WARN/ERR/END).</param>
        ''' <param name="ex">Exception instance to serialize.</param>
        Private Shared Sub WriteException(category As String, ex As Exception)
            If ex Is Nothing Then Return
            WriteLine(category, $"Exception Type: {ex.GetType().FullName}")
            WriteLine(category, $"Exception Message: {ex.Message}")
            If Not String.IsNullOrWhiteSpace(ex.StackTrace) Then
                WriteLine(category, "Stack Trace:")
                WriteRaw(category, ex.StackTrace)
            End If
            If ex.InnerException IsNot Nothing Then
                WriteLine(category, $"Inner Exception Type: {ex.InnerException.GetType().FullName}")
                WriteLine(category, $"Inner Exception Message: {ex.InnerException.Message}")
            End If
        End Sub

        ''' <summary>
        ''' Logs every public instance property of <see cref="ModelConfig"/> (unmodified) for diagnostics,
        ''' skipping a fixed list of sensitive/high-volume properties.
        ''' </summary>
        ''' <param name="config">Configuration instance to log.</param>
        ''' <param name="label">Label written before the config dump.</param>
        Public Shared Sub LogModelConfigOnce(config As ModelConfig, label As String)
            If Not _isEnabled OrElse config Is Nothing Then Return

            WriteLine("CONF", $"{label}:")

            Try
                Dim excludedExact As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "APIEncrypted",
                    "APIKey",
                    "APIKeyBack",
                    "APIKeyPrefix",
                    "DecodedAPI",
                    "TokenCount",
                    "MaxOutputToken",
                    "MergePrompt",
                    "QueryPrompt",
                    "TokenExpiry"
                }

                Dim excludedPrefixes As String() = {
                    "OAuth2",
                    "Parameter"
                }

                Dim props = GetType(ModelConfig).GetProperties(BindingFlags.Instance Or BindingFlags.Public).
                    Where(Function(p)
                              If excludedExact.Contains(p.Name) Then Return False
                              For Each prefix In excludedPrefixes
                                  If p.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return False
                              Next
                              Return True
                          End Function).
                    OrderBy(Function(p) p.Name).
                    ToList()

                For Each p In props
                    Dim v As Object = Nothing
                    Try
                        v = p.GetValue(config, Nothing)
                    Catch ex As Exception
                        WriteLine("CONF", $"  {p.Name}: <error reading>")
                        WriteLine("CONF", $"    {ex.GetType().Name}: {ex.Message}")
                        Continue For
                    End Try

                    Dim textValue As String
                    If v Is Nothing Then
                        textValue = ""
                    ElseIf TypeOf v Is DateTime Then
                        textValue = DirectCast(v, DateTime).ToString("yyyy-MM-dd HH:mm:ss.fff")
                    Else
                        textValue = v.ToString()
                    End If

                    WriteLine("CONF", $"  {p.Name}: {textValue}")
                Next
            Catch ex As Exception
                LogError("Failed to log ModelConfig.", ex:=ex)
            End Try
        End Sub

        ''' <summary>
        ''' Logs a snapshot of selected tooling-related INI variables prior to calling the main tool-enabled LLM.
        ''' Tool instructions and tool responses are logged as length stubs only (full content is already
        ''' recorded via <see cref="LogModelConfigOnce"/> and <see cref="LogRawResponse"/>).
        ''' </summary>
        Public Shared Sub LogPreMainLlmCallSnapshot()
            If Not _isEnabled Then Return
            WriteLine("LLM", "Pre LLM() snapshot (main tooling LLM):")
            WriteLine("LLM", $"  INI_Model_2: {SafeStr(INI_Model_2)}")
            WriteLine("LLM", $"  INI_APICall_2: {SafeStr(INI_APICall_2)}")

            Dim toolInstr = SafeStr(INI_APICall_ToolInstructions_2)
            WriteLine("LLM", $"  INI_APICall_ToolInstructions_2: ({toolInstr.Length} chars)")

            Dim toolResp = SafeStr(INI_APICall_ToolResponses_2)
            If toolResp.Length <= 500 Then
                WriteLine("LLM", $"  INI_APICall_ToolResponses_2: {toolResp}")
            Else
                Dim excerpt = toolResp.Substring(0, 500) & "..."
                WriteLine("LLM", $"  INI_APICall_ToolResponses_2: ({toolResp.Length} chars) {excerpt}")
            End If

            WriteLine("LLM", $"  INI_Response_2: {SafeStr(INI_Response_2)}")
        End Sub

        ''' <summary>
        ''' Logs a snapshot of selected variables prior to calling an external tool/service via <c>LLM</c>.
        ''' </summary>
        ''' <param name="ctx">Object expected to expose <c>INI_Model_2</c> and <c>INI_APICall_2</c> members.</param>
        Public Shared Sub LogPreToolLlmCallSnapshot(ctx As Object)
            If Not _isEnabled Then Return

            WriteLine("LLM", "Pre LLM() snapshot (tool/service call):")
            Try
                WriteLine("LLM", $"  INI_Model_2: {SafeGetMemberString(ctx, "INI_Model_2")}")
                WriteLine("LLM", $"  INI_APICall_2: {SafeGetMemberString(ctx, "INI_APICall_2")}")
            Catch ex As Exception
                LogError("Failed to capture tool LLM snapshot.", ex:=ex)
            End Try
        End Sub

        ''' <summary>
        ''' Logs raw response content (unmodified) with two blank lines before and after.
        ''' </summary>
        ''' <param name="source">Source label (e.g. "Main LLM()" or tool name).</param>
        ''' <param name="rawResponse">Raw response text.</param>
        Public Shared Sub LogRawResponse(source As String, rawResponse As String)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            WriteLine("RESP", $"Raw response ({source}) begins:")
            WriteRaw("RESP", vbCrLf & vbCrLf & SafeStr(rawResponse) & vbCrLf & vbCrLf)
            WriteLine("RESP", $"Raw response ({source}) ends.")
        End Sub

        ''' <summary>
        ''' Logs a brief stub of a raw response (length + short excerpt) without the full content.
        ''' Used for main LLM responses to keep the log file focused on tool calls.
        ''' </summary>
        ''' <param name="source">Source label (e.g. "Main LLM()").</param>
        ''' <param name="rawResponse">Raw response text.</param>
        ''' <param name="excerptLength">Maximum number of characters to include in the excerpt.</param>
        Public Shared Sub LogRawResponseStub(source As String, rawResponse As String, Optional excerptLength As Integer = 200)
            If Not _isEnabled OrElse String.IsNullOrWhiteSpace(_logPath) Then Return

            Dim safe As String = SafeStr(rawResponse)
            Dim charCount As Integer = safe.Length
            Dim excerpt As String = If(charCount <= excerptLength,
                safe,
                safe.Substring(0, excerptLength) & "...")

            WriteLine("RESP", $"Raw response ({source}): {charCount} chars")
            If charCount > 0 Then
                WriteLine("RESP", $"Excerpt: {excerpt}")
            End If
        End Sub

        ''' <summary>
        ''' Ends a tooling log session and writes summary/exception details (if provided).
        ''' </summary>
        ''' <param name="success">Whether the session completed successfully.</param>
        ''' <param name="summary">Optional summary string written to the log.</param>
        ''' <param name="ex">Optional exception written to the log.</param>
        Public Shared Sub EndSession(Optional success As Boolean = True, Optional summary As String = "", Optional ex As Exception = Nothing)
            If Not _isEnabled Then Return
            WriteLine("END", $"Success: {success}")
            If Not String.IsNullOrWhiteSpace(summary) Then
                WriteLine("END", $"Summary: {summary}")
            End If
            If ex IsNot Nothing Then
                WriteException("END", ex)
            End If
            WriteLine("END", $"Ended: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}")

            _isEnabled = False
            _started = False
            _logPath = Nothing
        End Sub

        ''' <summary>
        ''' Writes a single timestamped log line.
        ''' </summary>
        ''' <param name="category">Category identifier (STEP/WARN/ERR/etc.).</param>
        ''' <param name="message">Message text.</param>
        Private Shared Sub WriteLine(category As String, message As String)
            SyncLock _lock
                Try
                    Dim entry = $"[{DateTime.Now:HH:mm:ss.fff}] [{category}] {message}"
                    File.AppendAllText(_logPath, entry & Environment.NewLine)
                Catch
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Writes raw text to the log file without adding timestamps.
        ''' </summary>
        ''' <param name="category">Unused category label; retained to match caller signature.</param>
        ''' <param name="raw">Raw text to append.</param>
        Private Shared Sub WriteRaw(category As String, raw As String)
            SyncLock _lock
                Try
                    ' Raw is written as-is; only prefixed by one timestamp line already.
                    File.AppendAllText(_logPath, raw)
                Catch
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' Converts a potentially <c>Nothing</c> string into a non-null string for logging.
        ''' </summary>
        Private Shared Function SafeStr(value As String) As String
            If value Is Nothing Then Return ""
            Return value
        End Function

        ''' <summary>
        ''' Reads a string representation of a named property or field value via reflection.
        ''' </summary>
        ''' <param name="obj">Target object.</param>
        ''' <param name="memberName">Property or field name.</param>
        ''' <returns>Member value converted to string, or an empty string if missing.</returns>
        Private Shared Function SafeGetMemberString(obj As Object, memberName As String) As String
            If obj Is Nothing Then Return ""
            Dim t = obj.GetType()
            Dim p = t.GetProperty(memberName)
            If p IsNot Nothing Then
                Dim v = p.GetValue(obj, Nothing)
                Return If(v IsNot Nothing, v.ToString(), "")
            End If
            Dim f = t.GetField(memberName)
            If f IsNot Nothing Then
                Dim v = f.GetValue(obj)
                Return If(v IsNot Nothing, v.ToString(), "")
            End If
            Return ""
        End Function

        ''' <summary>
        ''' Returns whether file logging is currently enabled.
        ''' </summary>
        Public Shared ReadOnly Property IsEnabled As Boolean
            Get
                Return _isEnabled
            End Get
        End Property

        ''' <summary>
        ''' Returns the currently active log file path (empty/Nothing if not enabled).
        ''' </summary>
        Public Shared ReadOnly Property LogFilePath As String
            Get
                Return _logPath
            End Get
        End Property
    End Class

#End Region

#Region "Tooling Data Classes"

    ''' <summary>
    ''' Represents a single tool call extracted from an LLM response.
    ''' </summary>
    Public Class ToolCall

        ''' <summary>Tool call identifier used to correlate call and response objects.</summary>
        Public Property CallId As String

        ''' <summary>Name of the tool to execute.</summary>
        Public Property ToolName As String

        ''' <summary>Arguments passed to the tool.</summary>
        Public Property Arguments As Dictionary(Of String, Object)

        ''' <summary>Raw JSON representation of the tool call token.</summary>
        Public Property RawJson As String

        ''' <summary>
        ''' Initializes a new tool call instance with an empty arguments dictionary.
        ''' </summary>
        Public Sub New()
            Arguments = New Dictionary(Of String, Object)()
        End Sub
    End Class

    ''' <summary>
    ''' Represents the outcome of executing a single tool call.
    ''' </summary>
    Public Class ToolResponse

        ''' <summary>Tool call identifier used to correlate call and response objects.</summary>
        Public Property CallId As String

        ''' <summary>Name of the tool that was executed.</summary>
        Public Property ToolName As String

        ''' <summary>Raw response returned by the tool execution.</summary>
        Public Property Response As String

        ''' <summary>True if the tool execution completed successfully; otherwise False.</summary>
        Public Property Success As Boolean

        ''' <summary>Error message populated when <see cref="Success"/> is False.</summary>
        Public Property ErrorMessage As String

        ''' <summary>Timestamp captured at response creation time.</summary>
        Public Property Timestamp As DateTime

        ''' <summary>Original tool call JSON as extracted from the LLM response.</summary>
        Public Property OriginalCallJson As String

        ''' <summary>
        ''' Initializes a new tool response instance with default success state.
        ''' </summary>
        Public Sub New()
            Timestamp = DateTime.Now
            Success = True
        End Sub
    End Class

    ''' <summary>
    ''' Holds per-run state for the tooling loop, including selected tools, iteration counters, and logging.
    ''' </summary>
    Public Class ToolExecutionContext

        ''' <summary>Tools selected for this session.</summary>
        Public Property SelectedTools As List(Of ModelConfig)

        ''' <summary>All responses generated during this session (successful and failed).</summary>
        Public Property AllToolResponses As List(Of ToolResponse)

        ''' <summary>Current iteration counter within <see cref="ExecuteToolingLoop"/>.</summary>
        Public Property CurrentIteration As Integer

        ''' <summary>Maximum permitted number of iterations.</summary>
        Public Property MaxIterations As Integer

        ''' <summary>Cancellation flag set by UI event handler.</summary>
        Public Property IsCancelled As Boolean

        ''' <summary>In-memory log entries appended during session execution.</summary>
        Public Property LogEntries As List(Of String)

        ''' <summary>Snapshot of the LLM/tooling model config used for tool call detection/extraction formats.</summary>
        Public Property ToolingModel As ModelConfig

        ''' <summary>Optional UI log window instance used for user-visible progress logging.</summary>
        Public Property LogWindowForm As LogWindow

        Public Property FailedToolCallCounts As Dictionary(Of String, Integer)
        Public Property DuplicateFailureAbortThreshold As Integer

        ''' <summary>
        ''' Initializes a new tool execution context with default collections and limits.
        ''' </summary>
        Public Sub New()
            SelectedTools = New List(Of ModelConfig)()
            AllToolResponses = New List(Of ToolResponse)()
            LogEntries = New List(Of String)()
            FailedToolCallCounts = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            DuplicateFailureAbortThreshold = 2
            CurrentIteration = 0
            MaxIterations = INI_ToolingMaximumIterations
            IsCancelled = False
        End Sub

        ''' <summary>
        ''' Appends a message to the in-memory log, debugger output, optional UI log window, and the tooling file log.
        ''' </summary>
        ''' <param name="message">Log message.</param>
        ''' <param name="level">Log level passed to the UI log window.</param>
        Public Sub Log(message As String, Optional level As String = "step")
            Dim entry = $"[{DateTime.Now:HH:mm:ss}] {message}"
            LogEntries.Add(entry)
            Debug.WriteLine($"[Tooling] {entry}")

            ToolingFileLogger.LogStep(message)

            If LogWindowForm IsNot Nothing AndAlso Not LogWindowForm.IsDisposed Then
                Try
                    LogWindowForm.AppendLog(message, level)
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to append to LogWindow.", ex:=ex)
                End Try
            End If
        End Sub

        ''' <summary>
        ''' Logs a warning through the session logger and file logger.
        ''' </summary>
        Public Sub LogWarn(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            Log(message, "warn")
            ToolingFileLogger.LogWarn(message, details, ex)
        End Sub

        ''' <summary>
        ''' Logs an error through the session logger and file logger.
        ''' </summary>
        Public Sub LogError(message As String, Optional details As String = "", Optional ex As Exception = Nothing)
            Log(message, "error")
            ToolingFileLogger.LogError(message, details, ex)
        End Sub

        ''' <summary>
        ''' Placeholder method (intentionally empty) retained by the caller surface.
        ''' </summary>
        Public Sub WriteDebugLog()
            ' Intentionally empty (avoid multiple log files).
        End Sub
    End Class

#End Region

#Region "Execute Tooling"

    ''' <summary>
    ''' Executes an iterative tool-enabled LLM loop until either no tool calls are detected, the maximum iteration count is reached,
    ''' or the user cancels. Tool call detection/extraction and response injection are controlled by the active tooling model config.
    ''' </summary>
    ''' <param name="sysCommand">Base system command prompt text.</param>
    ''' <param name="userText">User prompt text (used only if fullPromptOverride is empty).</param>
    ''' <param name="selectedTools">Tool configurations available to the model.</param>
    ''' <param name="useSecondAPI">Whether to route LLM calls through the secondary API.</param>
    ''' <param name="fileObject">Optional file object payload passed to <c>LLM</c>.</param>
    ''' <param name="doTPMarkup">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="bubblesText">Optional bubble text appended to the user prompt.</param>
    ''' <param name="noFormatting">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="keepFormat">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="slideDeck">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="doMyStyle">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="myStyleInsert">Unused parameter in this method body (passed through by caller signature).</param>
    ''' <param name="addDocs">If True, <paramref name="insertDocs"/> is appended to the user prompt (when provided).</param>
    ''' <param name="insertDocs">Optional inserted document text appended to the user prompt when <paramref name="addDocs"/> is True.</param>
    ''' <param name="slideInsert">Optional slide text appended to the user prompt.</param>
    ''' <param name="otherPrompt">Optional additional user prompt passed to <c>LLM</c>.</param>
    ''' <param name="fullPromptOverride">When provided, uses this as the complete user prompt instead of building one internally. 
    ''' This ensures tooling calls receive the same context as non-tooling LLM calls.</param>
    ''' <param name="hideSplash">When True, suppresses the splash/progress indicator during LLM calls.</param>
    ''' <param name="hideLogWindow">When True, suppresses the tooling log window (useful for chat integration).</param>
    ''' <param name="DoChart">When True, adds charting instructions to the system prompt.</param>
    ''' <returns>The final LLM response string returned by the last iteration.</returns>
    ''' <summary>
    ''' Executes an iterative tool-enabled LLM loop until either no tool calls are detected, the maximum iteration count is reached,
    ''' or the user cancels. Tool call detection/extraction and response injection are controlled by the active tooling model config.
    ''' </summary>
    Public Async Function ExecuteToolingLoop(
        sysCommand As String,
        userText As String,
        selectedTools As List(Of ModelConfig),
        useSecondAPI As Boolean,
        Optional fileObject As String = "",
        Optional doTPMarkup As Boolean = False,
        Optional bubblesText As String = "",
        Optional noFormatting As Boolean = False,
        Optional keepFormat As Boolean = False,
        Optional slideDeck As String = "",
        Optional doMyStyle As Boolean = False,
        Optional myStyleInsert As String = "",
        Optional addDocs As Boolean = False,
        Optional insertDocs As String = "",
        Optional slideInsert As String = "",
        Optional otherPrompt As String = "",
        Optional fullPromptOverride As String = "",
        Optional hideSplash As Boolean = False,
        Optional hideLogWindow As Boolean = False,
        Optional DoChart As Boolean = False,
        Optional cancellationToken As System.Threading.CancellationToken = Nothing,
        Optional binaryOutputDirectory As String = Nothing) As Task(Of String)

        ' Check for power transition BEFORE starting (matches RunLlmAsync pattern)
        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
            Return "Operation cancelled due to power transition."
        End If

        ToolingFileLogger.StartSession()

        Dim context As New ToolExecutionContext() With {
            .MaxIterations = INI_ToolingMaximumIterations,
            .SelectedTools = selectedTools
        }

        context.ToolingModel = GetCurrentConfig(_context)

        ' Start-of-run config logging (ONCE)
        ToolingFileLogger.LogModelConfigOnce(context.ToolingModel, "Tooling LLM ModelConfig")

        Dim internalSelected As Boolean =
            (selectedTools IsNot Nothing AndAlso selectedTools.Any(Function(t) t.ToolName IsNot Nothing AndAlso t.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase)))

        ToolingFileLogger.LogStep($"Internal web tool selected: {internalSelected}")

        If selectedTools IsNot Nothing Then
            For i = 0 To selectedTools.Count - 1
                ToolingFileLogger.LogModelConfigOnce(selectedTools(i), $"Selected Tool #{i + 1} ModelConfig")
            Next
        End If

        If INI_ToolingLogWindow AndAlso Not hideLogWindow Then
            ' Create and show the log window on the UI thread to avoid threading issues
            Dim logForm As LogWindow = Nothing
            Try
                ' Use Await with SwitchToUi for proper async UI thread marshaling
                Await SwitchToUi(Sub()
                                     logForm = New LogWindow()
                                     logForm.Show()
                                 End Sub).ConfigureAwait(False)
            Catch ex As Exception
                ToolingFileLogger.LogWarn("Failed to create LogWindow.", ex:=ex)
            End Try

            context.LogWindowForm = logForm
            If logForm IsNot Nothing Then
                AddHandler logForm.CancelRequested, Sub() context.IsCancelled = True
            End If
        End If

        ' Make context available to ApDashboardLog for Chat Agent routing
        _activeToolingContext = context

        context.Log("Starting tooling session...")
        If selectedTools IsNot Nothing Then
            context.Log($"Selected tools: {String.Join(", ", selectedTools.Select(Function(t) t.ToolName))}")
        Else
            context.Log("Selected tools: (none)")
        End If

        ' Calculate effective timeout for LLM calls
        Dim effectiveTimeout As Integer = CInt(If(useSecondAPI, INI_Timeout_2, INI_Timeout))

        Try
            ' Build System Prompt (matching direct LLM() call plus Tooling Instructions)
            Dim baseSysPrompt As String = sysCommand

            If Not String.IsNullOrWhiteSpace(bubblesText) Then
                baseSysPrompt &= " " & SP_Add_BubblesExtract
            End If

            If doTPMarkup Then
                baseSysPrompt &= " " & SP_Add_Revisions
            End If

            If Not String.IsNullOrWhiteSpace(slideDeck) Then
                baseSysPrompt &= " " & SP_Add_Slides
            ElseIf Not noFormatting Then
                If keepFormat Then
                    baseSysPrompt &= " " & SP_Add_KeepHTMLIntact
                Else
                    baseSysPrompt &= " " & SP_Add_KeepInlineIntact
                End If
            End If

            If doMyStyle AndAlso Not String.IsNullOrWhiteSpace(myStyleInsert) Then
                baseSysPrompt &= " " & myStyleInsert
            End If

            If DoChart Then
                baseSysPrompt &= " " & SP_Add_Chart
            End If

            Dim enhancedSysPrompt As String = baseSysPrompt & Environment.NewLine & Environment.NewLine & BuildToolInstructionsPrompt(selectedTools)

            Dim toolDefinitions = BuildToolInstructionsForModel(selectedTools, context.ToolingModel)
            INI_APICall_ToolInstructions_2 = toolDefinitions
            INI_APICall_ToolResponses_2 = ""

            context.Log("Tool definitions prepared for model")

            If INI_ToolingDryRun Then
                Dim preview = $"The following tools will be made available to the model:{Environment.NewLine}{Environment.NewLine}"
                For Each tool In selectedTools
                    preview &= $"- {tool.ToolName}: {tool.ToolInstructionsPrompt}{Environment.NewLine}"
                Next

                Dim proceed = ShowCustomYesNoBox(preview & Environment.NewLine & "Do you want to proceed with the tool-enabled call?", "Proceed", "Abort")
                If proceed <> 1 Then
                    context.LogWarn("Dry run aborted by user")
                    ToolingFileLogger.EndSession(False, "Dry run aborted by user")
                    Return ""
                End If
            End If

            Dim currentResponse As String = ""
            Dim iteration As Integer = 0
            Dim fullUserPrompt As String = ""
            Dim abortDueToToolError As Boolean = False
            Dim abortToolErrorMessage As String = ""
            Dim abortToolName As String = ""
            Dim abortToolParamSummary As String = ""
            Dim abortToolRawCallJson As String = ""
            Dim abortFactsPrompt As String = ""

            Dim noSelectedText As Boolean = String.IsNullOrWhiteSpace(userText)
            While iteration < context.MaxIterations AndAlso Not context.IsCancelled

                ' Check for cancellation token at each iteration
                If cancellationToken.IsCancellationRequested Then
                    context.LogWarn("Cancellation requested via token")
                    Exit While
                End If

                ' Check for power transition during loop (matches RunLlmAsync pattern)
                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                    context.LogWarn("Power transition detected, aborting tooling loop")
                    ToolingFileLogger.EndSession(False, "Cancelled due to power transition")
                    Return "Operation cancelled due to power transition."
                End If

                iteration += 1
                context.CurrentIteration = iteration
                context.Log($"--- Iteration {iteration} of {context.MaxIterations} ---")

                context.Log("Calling LLM...", "llm")

                If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then
                    fullUserPrompt = fullPromptOverride
                Else
                    If noSelectedText Then
                        fullUserPrompt = If(addDocs AndAlso Not String.IsNullOrWhiteSpace(insertDocs), " " & insertDocs & " ", "")
                        If Not String.IsNullOrWhiteSpace(slideInsert) Then
                            fullUserPrompt &= slideInsert
                        End If
                    Else
                        fullUserPrompt = "<TEXTTOPROCESS>" & userText & "</TEXTTOPROCESS>"
                        If addDocs AndAlso Not String.IsNullOrWhiteSpace(insertDocs) Then
                            fullUserPrompt &= " " & insertDocs & " "
                        End If
                        If Not String.IsNullOrWhiteSpace(slideInsert) Then
                            fullUserPrompt &= slideInsert
                        End If
                        If Not String.IsNullOrWhiteSpace(bubblesText) Then
                            fullUserPrompt &= " " & bubblesText
                        End If
                    End If
                End If

                ToolingFileLogger.LogPreMainLlmCallSnapshot()

                ' Create linked cancellation with timeout (matches RunLlmAsync pattern)
                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60 ' 60 second buffer
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
                        System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            currentResponse = Await LLM(
                                enhancedSysPrompt,
                                fullUserPrompt,
                                "", "", 0,
                                useSecondAPI,
                                hideSplash,
                                otherPrompt,
                                fileObject,
                                combinedCts.Token,
                                True, True, binaryOutputDirectory:=binaryOutputDirectory)

                        Catch ex As OperationCanceledException When timeoutCts.IsCancellationRequested
                            context.LogError($"LLM call timed out after {totalTimeout}s")
                            ToolingFileLogger.EndSession(False, $"Timeout after {totalTimeout}s")
                            Return "Operation timed out. Please try again with a shorter prompt or different model."

                        Catch ex As OperationCanceledException
                            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                                context.LogWarn("Cancelled due to power transition")
                                ToolingFileLogger.EndSession(False, "Cancelled due to power transition")
                                Return "Operation cancelled due to power transition."
                            End If
                            context.LogWarn("Operation was cancelled")
                            ToolingFileLogger.EndSession(False, "Cancelled by user")
                            Return "Operation was canceled by the user."
                        End Try
                    End Using
                End Using

                ToolingFileLogger.LogRawResponseStub("Main LLM()", currentResponse)

                If String.IsNullOrWhiteSpace(currentResponse) Then
                    context.LogWarn("Empty response from LLM", details:="LLM() returned null/empty/whitespace.")
                    Exit While
                End If

                context.Log($"Response received ({currentResponse.Length} chars)")

                Dim detectionPattern = context.ToolingModel.ToolCallDetectionPattern

                If ContainsToolCalls(currentResponse, detectionPattern) Then
                    context.Log("Tool calls detected in response")

                    Dim extractionMap = context.ToolingModel.ToolCallExtractionMap
                    Dim toolCalls = ExtractToolCalls(currentResponse, extractionMap)
                    context.Log($"Extracted {toolCalls.Count} tool call(s)")

                    If toolCalls.Count = 0 Then
                        context.LogWarn(
                            "Tool calls detected but none could be parsed; treating as text response",
                            details:=$"ToolCallExtractionMap='{extractionMap}'")
                        Exit While
                    End If

                    For Each tc In toolCalls
                        If context.IsCancelled OrElse cancellationToken.IsCancellationRequested Then Exit For

                        ' Check power transition before each tool execution
                        If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                            context.LogWarn("Power transition detected during tool execution")
                            Exit For
                        End If

                        context.Log($"Executing tool: {tc.ToolName} (ID: {tc.CallId})")

                        Dim toolConfig = selectedTools.FirstOrDefault(
                            Function(t) t.ToolName.Equals(tc.ToolName, StringComparison.OrdinalIgnoreCase))

                        If toolConfig Is Nothing Then
                            If tc.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) Then
                                toolConfig = GetInternalWebTool()
                                ToolingFileLogger.LogStep("Using internal web tool.")
                            ElseIf tc.ToolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) Then
                                ' Determine privacy flag: AutoPilot config takes precedence, then INI setting
                                Dim enforcePrivacy As Boolean
                                If _apConfig IsNot Nothing Then
                                    enforcePrivacy = _apConfig.EnablePrivacyProtection
                                Else
                                    enforcePrivacy = INI_EnablePrivacyForSearch
                                End If
                                toolConfig = GetInternalSearchTool(enforcePrivacy:=enforcePrivacy)
                                ToolingFileLogger.LogStep("Using internal search tool.")
                            ElseIf IsInternalKnowledgeToolName(tc.ToolName) Then
                                toolConfig = GetInternalKnowledgeTool(tc.ToolName)
                                If toolConfig IsNot Nothing Then
                                    ToolingFileLogger.LogStep("Using store-specific internal knowledge tool.")
                                End If
                            End If

                            If toolConfig Is Nothing Then
                                context.LogError(
                                    $"Unknown tool: {tc.ToolName}",
                                    details:=$"CallId={tc.CallId}; Raw={tc.RawJson}")

                                Dim errorResp As New ToolResponse() With {
                                    .CallId = tc.CallId,
                                    .ToolName = tc.ToolName,
                                    .Success = False,
                                    .ErrorMessage = $"Unknown tool: {tc.ToolName}",
                                    .OriginalCallJson = tc.RawJson
                                }

                                context.AllToolResponses.Add(errorResp)
                                Continue For
                            End If
                        End If

                        Dim toolCallSignature = BuildToolCallSignature(tc)
                        Dim previousFailureCount As Integer = 0

                        If context.FailedToolCallCounts.TryGetValue(toolCallSignature, previousFailureCount) AndAlso
                           previousFailureCount >= context.DuplicateFailureAbortThreshold Then

                            Dim duplicateMsg =
                                        $"Aborting because the same failing tool call was repeated {previousFailureCount} time(s): {tc.ToolName}. " &
                                        "The model should revise its plan instead of retrying the identical call."

                            context.LogError(duplicateMsg, details:=$"CallId={tc.CallId}; RawCall={tc.RawJson}")
                            abortDueToToolError = True
                            abortToolName = tc.ToolName
                            abortToolParamSummary = BuildCondensedParamSummary(tc.Arguments)
                            abortToolRawCallJson = tc.RawJson
                            abortToolErrorMessage = duplicateMsg
                            Exit For
                        End If

                        Dim toolResponse = Await ExecuteToolCall(tc, toolConfig, context, cancellationToken)
                        toolResponse.OriginalCallJson = tc.RawJson
                        context.AllToolResponses.Add(toolResponse)

                        If Not toolResponse.Success Then
                            context.LogError(
                                $"Tool error ({tc.ToolName}): {toolResponse.ErrorMessage}",
                                details:=$"CallId={tc.CallId}; RawCall={tc.RawJson}")

                            If context.FailedToolCallCounts.ContainsKey(toolCallSignature) Then
                                context.FailedToolCallCounts(toolCallSignature) += 1
                            Else
                                context.FailedToolCallCounts(toolCallSignature) = 1
                            End If

                            Select Case toolConfig.ToolErrorHandling?.ToLowerInvariant()
                                Case "abort"
                                    context.LogError("Aborting due to tool error (ToolErrorHandling=abort)")

                                    If ShouldShowToolingModalDialogs() Then
                                        ShowCustomMessageBox($"Tool execution failed: {toolResponse.ErrorMessage}")
                                    End If

                                    abortDueToToolError = True
                                    abortToolName = tc.ToolName
                                    abortToolParamSummary = BuildCondensedParamSummary(tc.Arguments)
                                    abortToolRawCallJson = tc.RawJson
                                    abortToolErrorMessage = If(toolResponse.ErrorMessage, "Unknown tool error.")
                                    Exit For

                                Case "retry"
                                    context.LogWarn("Will retry on next iteration (ToolErrorHandling=retry)")
                                Case Else
                                    context.LogWarn("Skipping tool error (ToolErrorHandling=skip)")
                            End Select
                        Else
                            If context.FailedToolCallCounts.ContainsKey(toolCallSignature) Then
                                context.FailedToolCallCounts.Remove(toolCallSignature)
                            End If

                            context.Log($"Tool completed successfully ({toolResponse.Response?.Length} chars)", "success")
                        End If
                    Next

                    Dim toolResponses = BuildToolResponsesForModel(context.AllToolResponses, context.ToolingModel)
                    INI_APICall_ToolResponses_2 = toolResponses
                    context.Log("Tool responses prepared for next iteration")

                    If abortDueToToolError Then
                        context.LogWarn("Stopping tooling loop after tool error abort")
                        Exit While
                    End If

                Else
                    context.Log("Text response received (no tool calls)")
                    Exit While
                End If

            End While

            ' If we hit max iterations and the last response was a tool call, force a final text response.
            ' The tool results are already in INI_APICall_ToolResponses_2 from the last iteration.
            If iteration >= context.MaxIterations AndAlso
               Not context.IsCancelled AndAlso
               Not cancellationToken.IsCancellationRequested AndAlso
               ContainsToolCalls(currentResponse, context.ToolingModel.ToolCallDetectionPattern) Then

                context.Log("Forcing final response (max iterations reached with pending tool call)...")

                ' Disable tool definitions to prevent further tool calls
                INI_APICall_ToolInstructions_2 = ""

                ' Append instruction to force synthesis
                Dim finalSysPrompt As String = enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                    "IMPORTANT: You have reached the maximum number of tool iterations. Do NOT call any more tools. " &
                    "Based on all the information gathered from the tools so far, provide your final answer now."

                ToolingFileLogger.LogStep("Forcing final LLM call without tools")
                ToolingFileLogger.LogPreMainLlmCallSnapshot()

                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
                        System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            Dim finalResponse As String = Await LLM(
                                finalSysPrompt,
                                fullUserPrompt,
                                "", "", 0,
                                useSecondAPI,
                                hideSplash,
                                otherPrompt,
                                fileObject,
                                combinedCts.Token,
                                True, True, binaryOutputDirectory:=binaryOutputDirectory)

                            If Not String.IsNullOrWhiteSpace(finalResponse) Then
                                currentResponse = finalResponse
                                context.Log($"Final response received ({currentResponse.Length} chars)")
                                ToolingFileLogger.LogRawResponseStub("Main LLM() - Forced Final", currentResponse)
                            Else
                                context.LogWarn("Empty response from forced final LLM call")
                            End If

                        Catch ex As OperationCanceledException
                            context.LogWarn("Forced final call was cancelled")
                        Catch ex As Exception
                            context.LogError($"Error during forced final call: {ex.Message}", ex:=ex)
                        End Try
                    End Using
                End Using
            End If

            If abortDueToToolError AndAlso
                   Not context.IsCancelled AndAlso
                   Not cancellationToken.IsCancellationRequested Then

                context.Log("Forcing final response after tool error abort...")

                ' Disable further tool calls, but keep accumulated tool responses available.
                INI_APICall_ToolInstructions_2 = ""

                Dim successfulToolFacts As New System.Text.StringBuilder()
                Dim failedToolFacts As New System.Text.StringBuilder()

                For Each tr In context.AllToolResponses
                    If tr Is Nothing Then Continue For

                    If tr.Success Then
                        successfulToolFacts.AppendLine($"- Tool: {tr.ToolName}")
                        If Not String.IsNullOrWhiteSpace(tr.Response) Then
                            successfulToolFacts.AppendLine($"  Result: {BuildResultExcerpt(tr.Response, 160)}")
                        End If
                    Else
                        failedToolFacts.AppendLine($"- Tool: {tr.ToolName}")
                        If Not String.IsNullOrWhiteSpace(tr.ErrorMessage) Then
                            failedToolFacts.AppendLine($"  Error: {tr.ErrorMessage}")
                        End If
                    End If
                Next

                abortFactsPrompt =
                        "<TOOL_ABORT_FACTS>" & Environment.NewLine &
                        "Use ONLY the facts in this block when explaining the failure." & Environment.NewLine &
                        "Do NOT replace the stated failure with another cause." & Environment.NewLine &
                        $"Failed tool: {abortToolName}" & Environment.NewLine &
                        $"Failed tool parameters: {abortToolParamSummary}" & Environment.NewLine &
                        $"Failed tool raw call JSON: {abortToolRawCallJson}" & Environment.NewLine &
                        $"Exact failure message: {abortToolErrorMessage}" & Environment.NewLine &
                        Environment.NewLine &
                        "<COMPLETED_TOOL_STEPS>" & Environment.NewLine &
                        If(successfulToolFacts.Length > 0, successfulToolFacts.ToString().TrimEnd(), "(none)") & Environment.NewLine &
                        "</COMPLETED_TOOL_STEPS>" & Environment.NewLine &
                        Environment.NewLine &
                        "<FAILED_TOOL_STEPS>" & Environment.NewLine &
                        If(failedToolFacts.Length > 0, failedToolFacts.ToString().TrimEnd(), "(none)") & Environment.NewLine &
                        "</FAILED_TOOL_STEPS>" & Environment.NewLine &
                        "</TOOL_ABORT_FACTS>"

                Dim abortFinalSysPrompt As String = enhancedSysPrompt & Environment.NewLine & Environment.NewLine &
                    "IMPORTANT: A tool-assisted run has stopped because a tool call failed. Do NOT call any more tools. " &
                    "Provide a concise, user-friendly status update. " &
                    "You MUST rely only on the explicitly supplied failure facts and completed-step facts. " &
                    "Do NOT infer a different cause. Do NOT rewrite the failure into another tool problem. " &
                    "Treat the exact failure message as authoritative. " &
                    "If the failed tool parameters indicate action='delete' or action='rmdir', describe it as a delete/remove attempt, not as a move. " &
                    "If the exact failure says permission was disabled, state that plainly and do not replace it with a path-validation explanation. " &
                    "Explain clearly: (1) what was completed successfully, (2) what failed, (3) why it failed, and (4) what therefore remains incomplete. " &
                    "Do not mention internal logs, JSON, raw tool protocols, or hidden implementation details."

                Dim abortFinalUserPrompt As String = fullUserPrompt & Environment.NewLine & Environment.NewLine &
                            abortFactsPrompt

                ToolingFileLogger.LogStep("Forcing final LLM call without tools after tool error abort")
                ToolingFileLogger.LogPreMainLlmCallSnapshot()

                Using timeoutCts As New System.Threading.CancellationTokenSource()
                    Dim totalTimeout = effectiveTimeout + 60
                    timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                    Using combinedCts As System.Threading.CancellationTokenSource =
            System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token)

                        Try
                            Dim finalAbortResponse As String = Await LLM(
                    abortFinalSysPrompt,
                    abortFinalUserPrompt,
                    "", "", 0,
                    useSecondAPI,
                    hideSplash,
                    otherPrompt,
                    fileObject,
                    combinedCts.Token,
                    True, True, binaryOutputDirectory:=binaryOutputDirectory)

                            If Not String.IsNullOrWhiteSpace(finalAbortResponse) Then
                                currentResponse = finalAbortResponse
                                context.Log($"Final abort summary received ({currentResponse.Length} chars)")
                                ToolingFileLogger.LogRawResponseStub("Main LLM() - Tool Error Final", currentResponse)
                            Else
                                context.LogWarn("Empty response from final abort-summary LLM call")
                                currentResponse =
                        If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                           "The tool-assisted run stopped before completion.",
                           $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")
                            End If

                        Catch ex As OperationCanceledException
                            context.LogWarn("Final abort-summary call was cancelled")
                            currentResponse =
                    If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                       "The tool-assisted run stopped before completion.",
                       $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")

                        Catch ex As Exception
                            context.LogError($"Error during final abort-summary call: {ex.Message}", ex:=ex)
                            currentResponse =
                    If(String.IsNullOrWhiteSpace(abortToolErrorMessage),
                       "The tool-assisted run stopped before completion.",
                       $"The tool-assisted run stopped before completion. Failed tool: {abortToolName}. Reason: {abortToolErrorMessage}")
                        End Try
                    End Using
                End Using
            End If

            If context.IsCancelled OrElse cancellationToken.IsCancellationRequested Then
                context.LogWarn("Session cancelled")
                ToolingFileLogger.EndSession(False, "Cancelled")
                Return If(cancellationToken.IsCancellationRequested, "Operation was canceled by the user.", "")
            End If

            If iteration >= context.MaxIterations Then
                context.LogWarn($"Maximum iterations ({context.MaxIterations}) reached")
                If ShouldShowToolingModalDialogs() Then
                    ShowCustomMessageBox($"Maximum tool iterations ({context.MaxIterations}) reached. The response may be incomplete.")
                End If
                ToolingFileLogger.LogWarn("Maximum iterations reached.", details:=$"MaxIterations={context.MaxIterations}")
            End If

            context.Log("=== Session Summary ===")
            context.Log($"Total iterations: {iteration}")
            context.Log($"Total tool calls: {context.AllToolResponses.Count}")
            Dim successCount As Integer = context.AllToolResponses.Where(Function(r) r.Success).Count()
            Dim failedCount As Integer = context.AllToolResponses.Where(Function(r) Not r.Success).Count()
            context.Log($"Successful: {successCount}", If(failedCount = 0, "success", "step"))
            context.Log($"Failed: {failedCount}", If(failedCount = 0, "step", "warn"))

            currentResponse = AppendM365SourcesFooter(currentResponse, context.AllToolResponses)

            Dim sessionSucceeded As Boolean = Not abortDueToToolError
            Dim sessionSummary As String =
    $"Iterations: {iteration}, Tool calls: {context.AllToolResponses.Count}, Success: {successCount}, Failed: {failedCount}"

            If abortDueToToolError AndAlso Not String.IsNullOrWhiteSpace(abortToolErrorMessage) Then
                sessionSummary &= $", Aborted due to tool error: {abortToolErrorMessage}"
            End If

            ToolingFileLogger.EndSession(sessionSucceeded, sessionSummary)
            Return currentResponse

        Catch ex As OperationCanceledException
            context.LogWarn("Operation cancelled")
            ToolingFileLogger.EndSession(False, "Cancelled")
            Return "Operation was canceled by the user."

        Catch ex As Exception
            context.LogError($"Error in tooling loop: {ex.Message}", ex:=ex)
            If ShouldShowToolingModalDialogs() Then
                ShowCustomMessageBox($"Error during tool execution: {ex.Message}")
            End If

            ToolingFileLogger.EndSession(False, $"Exception: {ex.Message}", ex:=ex)
            Return $"Error during tool execution: {ex.Message}"
        Finally
            INI_APICall_ToolInstructions_2 = ""
            INI_APICall_ToolResponses_2 = ""

            If context.LogWindowForm IsNot Nothing AndAlso Not context.LogWindowForm.IsDisposed Then
                Try
                    context.LogWindowForm.MarkComplete()
                Catch ex As Exception
                    ToolingFileLogger.LogWarn("Failed to mark LogWindow complete.", ex:=ex)
                End Try

                context.Log("Session complete - close this window when ready")
                If ToolingFileLogger.IsEnabled Then
                    context.Log($"Log saved to: {ToolingFileLogger.LogFilePath}")
                End If
            End If

            Try
                _lastCompletedToolResponses = New List(Of ToolResponse)()
                If context IsNot Nothing AndAlso context.AllToolResponses IsNot Nothing Then
                    Debug.WriteLine("[AISearch] Persisting completed tool responses: " & context.AllToolResponses.Count.ToString())

                    For Each r In context.AllToolResponses
                        _lastCompletedToolResponses.Add(New ToolResponse() With {
                            .CallId = r.CallId,
                            .ToolName = r.ToolName,
                            .Response = r.Response,
                            .Success = r.Success,
                            .ErrorMessage = r.ErrorMessage,
                            .Timestamp = r.Timestamp,
                            .OriginalCallJson = r.OriginalCallJson
                        })
                    Next
                Else
                    Debug.WriteLine("[AISearch] Persisting completed tool responses: context or response list is Nothing")
                End If
            Catch ex As Exception
                Debug.WriteLine("[AISearch] Persisting completed tool responses failed: " & ex.Message)
            End Try

            ' Clear the context reference so ApDashboardLog stops routing
            _activeToolingContext = Nothing
        End Try
    End Function
#End Region

#Region "Tooling Helper Functions"

    Private Function BuildToolCallSignature(toolCall As ToolCall) As String
        If toolCall Is Nothing Then Return ""

        Dim parts As New List(Of String)()

        If toolCall.Arguments IsNot Nothing Then
            For Each kvp In toolCall.Arguments.OrderBy(Function(x) x.Key, StringComparer.OrdinalIgnoreCase)
                parts.Add(kvp.Key.ToLowerInvariant() & "=" & NormalizeToolArgumentValue(kvp.Value))
            Next
        End If

        Return toolCall.ToolName & "|" & String.Join(";", parts)
    End Function

    Private Function NormalizeToolArgumentValue(value As Object) As String
        If value Is Nothing Then Return "null"

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString(Formatting.None)
        End If

        If TypeOf value Is JToken Then
            Return DirectCast(value, JToken).ToString(Formatting.None)
        End If

        If TypeOf value Is IEnumerable(Of Object) AndAlso Not TypeOf value Is String Then
            Dim items = DirectCast(value, IEnumerable(Of Object)).
                Select(Function(v) NormalizeToolArgumentValue(v))
            Return "[" & String.Join(",", items) & "]"
        End If

        Return System.Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function ShouldShowToolingModalDialogs() As Boolean
        Return Not _chatAgentActive AndAlso Not _apActive
    End Function

    Private Const InternalKnowledgeToolNamePrefix As String = "knowledge_search_store_"

    Private Function EncodeToolToken(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        ' Use a SHA256 hash (hex, lowercase) to produce a fixed-length token
        ' that is always valid for API function names and stays well under
        ' the 128-character name limit imposed by model APIs (e.g. Gemini).
        Using sha = System.Security.Cryptography.SHA256.Create()
            Dim hashBytes = sha.ComputeHash(Encoding.UTF8.GetBytes(value))
            Dim sb As New StringBuilder(hashBytes.Length * 2)
            For Each b In hashBytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()   ' 64 hex chars, always
        End Using
    End Function

    Private Function DecodeToolToken(value As String) As String
        ' Hash-based tokens are one-way; decoding is no longer possible.
        ' Callers must use GetKnowledgeStoreForToolName which matches
        ' by recomputing hashes against known stores.
        Return ""
    End Function

    Private Function IsInternalKnowledgeToolName(toolName As String) As Boolean
        Return Not String.IsNullOrWhiteSpace(toolName) AndAlso
               toolName.StartsWith(InternalKnowledgeToolNamePrefix, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function BuildInternalKnowledgeToolName(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        If store Is Nothing Then Return ""
        Return InternalKnowledgeToolNamePrefix & EncodeToolToken(store.StoreId)
    End Function

    Private Function GetKnowledgeStoreForToolName(toolName As String) As KnowledgeStoreCatalog.KnowledgeStoreDefinition
        If Not IsInternalKnowledgeToolName(toolName) Then
            Return Nothing
        End If

        Dim encodedToken As String = toolName.Substring(InternalKnowledgeToolNamePrefix.Length)

        ' Hash-based tokens are one-way — match by recomputing the hash for
        ' each known store and comparing against the token in the tool name.
        Dim indexedStores = GetIndexedKnowledgeStores()
        If indexedStores Is Nothing OrElse indexedStores.Count = 0 Then Return Nothing

        ' Exact hash match
        For Each store In indexedStores
            Dim expectedHash = EncodeToolToken(store.StoreId)
            If String.Equals(encodedToken, expectedHash, StringComparison.OrdinalIgnoreCase) Then
                Return store
            End If
        Next

        ' If there is only one knowledge store, return it directly — no ambiguity
        If indexedStores.Count = 1 Then
            ToolingFileLogger.LogWarn(
                "Knowledge tool name did not match any store hash; " &
                "falling back to the only available store.",
                details:=$"ToolName='{toolName}', Token='{encodedToken}'")
            Return indexedStores(0)
        End If

        ' Multiple stores: fuzzy match by longest common prefix of the token
        ' (handles cases where the LLM truncates the hash)
        Dim bestMatch As KnowledgeStoreCatalog.KnowledgeStoreDefinition = Nothing
        Dim bestMatchLen As Integer = 0

        For Each store In indexedStores
            Dim expectedName = BuildInternalKnowledgeToolName(store)
            If String.IsNullOrWhiteSpace(expectedName) Then Continue For

            Dim commonLen = GetCommonPrefixLength(toolName, expectedName)
            If commonLen > InternalKnowledgeToolNamePrefix.Length AndAlso commonLen > bestMatchLen Then
                bestMatchLen = commonLen
                bestMatch = store
            End If
        Next

        If bestMatch IsNot Nothing Then
            ToolingFileLogger.LogWarn(
                $"Knowledge tool name partially matched store '{bestMatch.Name}' " &
                $"(prefix match: {bestMatchLen} of {toolName.Length} chars).",
                details:=$"ToolName='{toolName}', Token='{encodedToken}'")
        End If

        Return bestMatch
    End Function

    Private Shared Function GetCommonPrefixLength(a As String, b As String) As Integer
        If a Is Nothing OrElse b Is Nothing Then Return 0
        Dim maxLen = Math.Min(a.Length, b.Length)
        For i As Integer = 0 To maxLen - 1
            If Char.ToUpperInvariant(a(i)) <> Char.ToUpperInvariant(b(i)) Then Return i
        Next
        Return maxLen
    End Function

    Private Function GetIndexedKnowledgeStores() As List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)
        Dim result As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            result = stores.
                Where(Function(s)
                          If s Is Nothing Then Return False

                          Try
                              Dim manifest = KnowledgeStoreManifest.Load(s)
                              Return manifest IsNot Nothing AndAlso
                                     manifest.Entries IsNot Nothing AndAlso
                                     manifest.Entries.Count > 0
                          Catch
                              Return False
                          End Try
                      End Function).
                OrderBy(Function(s) If(KnowledgeStoreCatalog.GetDisplayLabel(s), "").ToLowerInvariant()).
                ToList()
        Catch
        End Try

        Return result
    End Function

    Private Function BuildInternalKnowledgeToolDefinition(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)
        Dim toolName As String = BuildInternalKnowledgeToolName(store)

        ' Load schema to get the user-authored tooling description
        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " " & schema.ToolingDescription.Trim()
        End If

        Dim definition As New JObject(
            New JProperty("name", toolName),
            New JProperty("description",
                $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} Use this for the user's own materials in that source, not for public-web lookup."),
            New JProperty("parameters",
                New JObject(
                    New JProperty("type", "object"),
                    New JProperty("properties",
                        New JObject(
                            New JProperty("query",
                                New JObject(
                                    New JProperty("type", "string"),
                                    New JProperty("description", "Optional natural-language query to search within this Knowledge Store.")
                                )
                            ),
                            New JProperty("tag",
                                New JObject(
                                    New JProperty("type", "string"),
                                    New JProperty("description", "Optional tag filter within this Knowledge Store.")
                                )
                            ),
                            New JProperty("max_results",
                                New JObject(
                                    New JProperty("type", "integer"),
                                    New JProperty("description", "Optional maximum number of results to retrieve.")
                                )
                            )
                        )
                    ),
                    New JProperty("additionalProperties", False)
                )
            )
        )

        Return definition.ToString(Formatting.None)
    End Function

    Private Function BuildInternalKnowledgeToolInstructionsPrompt(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As String
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)

        ' Load schema to get the user-authored tooling description
        Dim schema = KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)
        Dim contentHint As String = ""
        If schema IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(schema.ToolingDescription) Then
            contentHint = " Content: " & schema.ToolingDescription.Trim()
        End If

        Return $"Searches only the user's Knowledge Store '{displayLabel}'.{contentHint} " &
            "Provide query (optional), tag (optional), and max_results (optional). " &
            "If query is omitted, the tool may return the most relevant documents from that store or all documents matching the tag. " &
            "Do NOT use this tool for public information or general knowledge. " &
            $"When citing results, mention the document name And store name '{displayLabel}'."
    End Function

    Private Function GetInternalKnowledgeTool(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition) As ModelConfig
        Dim displayLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(store)

        Return New ModelConfig() With {
            .ToolName = BuildInternalKnowledgeToolName(store),
            .ToolInstructionsPrompt = BuildInternalKnowledgeToolInstructionsPrompt(store),
            .ToolDefinition = BuildInternalKnowledgeToolDefinition(store),
            .ModelDescription = $"Knowledge Store: {displayLabel}{InternalToolSuffix}",
            .Tool = True,
            .ToolPriority = 997,
            .ToolErrorHandling = "skip"
        }
    End Function

    Private Function GetInternalKnowledgeTool(toolName As String) As ModelConfig
        Dim store = GetKnowledgeStoreForToolName(toolName)
        If store Is Nothing Then Return Nothing
        Return GetInternalKnowledgeTool(store)
    End Function

    Private Function GetInternalKnowledgeTools() As List(Of ModelConfig)
        Return GetIndexedKnowledgeStores().
            Select(Function(store) GetInternalKnowledgeTool(store)).
            Where(Function(tool) tool IsNot Nothing).
            ToList()
    End Function

    ''' <summary>
    ''' Converts a canonical tool definition JSON string to a model-specific format using a template string.
    ''' </summary>
    ''' <param name="canonicalDefinition">Canonical tool definition JSON (must parse as a JSON object).</param>
    ''' <param name="template">Template used to render the model-specific definition.</param>
    ''' <returns>Rendered tool definition string, or an empty string on error.</returns>
    Public Function ConvertCanonicalToModelFormat(canonicalDefinition As String, template As String) As String
        If String.IsNullOrWhiteSpace(canonicalDefinition) OrElse String.IsNullOrWhiteSpace(template) Then
            ToolingFileLogger.LogWarn(
                "ConvertCanonicalToModelFormat: Empty input.",
                details:=$"canonicalDefinitionEmpty={String.IsNullOrWhiteSpace(canonicalDefinition)}; templateEmpty={String.IsNullOrWhiteSpace(template)}")
            Return ""
        End If

        Try
            Dim jDef As JObject = JObject.Parse(canonicalDefinition)
            Dim name As String = If(jDef("name")?.ToString(), "")
            Dim description As String = If(jDef("description")?.ToString(), "")

            ' Use Formatting.None to produce compact JSON for the parameters object.
            ' JObject.ToString() defaults to Formatting.Indented, which injects literal
            ' newlines and whitespace that bloat the payload and can break model API
            ' templates that expect single-line JSON values.
            Dim parametersToken As JToken = jDef("parameters")
            Dim parameters As String = If(parametersToken IsNot Nothing,
                parametersToken.ToString(Formatting.None), "{}")

            ' JSON-escape name and description before injecting into the template.
            ' JValue.ToString() returns the RAW unescaped string (e.g., embedded " or \
            ' are not escaped). When the template places these inside JSON string
            ' literals like "name":"{name}", unescaped characters produce invalid JSON.
            ' This is especially critical when combining multiple tools — a single
            ' malformed definition breaks the entire tools array and the API rejects
            ' the request, causing LLM() to return an empty string.
            Dim result As String = template
            result = result.Replace("{name}", EscapeJsonString(name))
            result = result.Replace("{description}", EscapeJsonString(description))
            result = result.Replace("{parameters}", parameters)

            Return result
        Catch ex As Exception
            ToolingFileLogger.LogError("ConvertCanonicalToModelFormat error.", details:=$"canonicalDefinition='{canonicalDefinition}'", ex:=ex)
            Debug.WriteLine($"ConvertCanonicalToModelFormat error: {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Builds the model-specific tool definitions block for the current tooling model by converting each selected tool definition
    ''' and injecting the resulting definition list into the model's <see cref="ModelConfig.APICall_ToolInstructions"/> template.
    ''' </summary>
    ''' <param name="selectedTools">Tools to include, sorted by <see cref="ModelConfig.ToolPriority"/>.</param>
    ''' <param name="toolingModel">The tooling model that defines the instruction template.</param>
    ''' <returns>Tool instructions string passed via <c>INI_APICall_ToolInstructions_2</c>.</returns>
    Public Function BuildToolInstructionsForModel(selectedTools As List(Of ModelConfig), toolingModel As ModelConfig) As String
        If toolingModel Is Nothing Then
            ToolingFileLogger.LogWarn("BuildToolInstructionsForModel: toolingModel is Nothing.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(toolingModel.APICall_ToolInstructions) Then
            ToolingFileLogger.LogWarn("BuildToolInstructionsForModel: toolingModel.APICall_ToolInstructions is empty.")
            Return ""
        End If

        Dim definitions As New StringBuilder()
        Dim isFirst As Boolean = True

        Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

        For Each tool In sortedTools
            If String.IsNullOrWhiteSpace(tool.ToolDefinition) Then
                ToolingFileLogger.LogWarn("Tool skipped: no ToolDefinition.", details:=$"ToolName='{tool.ToolName}'")
                Continue For
            End If

            Dim modelSpecificDef = ConvertCanonicalToModelFormat(
                tool.ToolDefinition,
                toolingModel.APICall_ToolInstructions_Template)

            If Not String.IsNullOrWhiteSpace(modelSpecificDef) Then
                If Not isFirst Then definitions.Append(",")
                definitions.Append(modelSpecificDef)
                isFirst = False
            Else
                ToolingFileLogger.LogWarn("Tool definition conversion returned empty.", details:=$"ToolName='{tool.ToolName}'")
            End If
        Next

        Dim result = toolingModel.APICall_ToolInstructions.Replace("{definitions}", definitions.ToString())
        result = result.Replace(LLM_APICall_Placeholder_ToolDefinitions.TrimStart("{"c).TrimEnd("}"c), definitions.ToString())

        Return result
    End Function

    ''' <summary>
    ''' Builds the model-specific tool response payload to inject into the next iteration of the tooling loop.
    ''' </summary>
    ''' <param name="responses">Tool execution outcomes to serialize.</param>
    ''' <param name="toolingModel">Tooling model that defines response templates and container structure.</param>
    ''' <returns>Serialized tool response payload.</returns>
    Public Function BuildToolResponsesForModel(responses As List(Of ToolResponse), toolingModel As ModelConfig) As String
        If toolingModel Is Nothing Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel is Nothing.")
            Return ""
        End If

        If String.IsNullOrWhiteSpace(toolingModel.APICall_ToolResponses) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses is empty.")
            Return ""
        End If

        Dim responsePartTemplate As String = toolingModel.APICall_ToolResponses_Template
        If String.IsNullOrWhiteSpace(responsePartTemplate) Then
            ToolingFileLogger.LogWarn("BuildToolResponsesForModel: toolingModel.APICall_ToolResponses_Template is empty.")
            Return ""
        End If

        Dim callPartTemplate As String = If(toolingModel.APICall_ToolCallPart_Template, "")
        Dim useCallParts As Boolean = Not String.IsNullOrWhiteSpace(callPartTemplate)

        Dim callParts As New StringBuilder()
        Dim responseParts As New StringBuilder()
        Dim firstCall As Boolean = True
        Dim firstResp As Boolean = True

        For Each resp In responses
            If useCallParts Then
                ' Extract the original arguments from the parsed tool call JSON
                Dim argsJson As String = "{}"
                Try
                    Dim jCall = JObject.Parse(resp.OriginalCallJson)
                    Dim argsToken = jCall("arguments")
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            argsJson = argsToken.ToString()
                        Else
                            argsJson = argsToken.ToString(Formatting.None)
                        End If
                    End If
                Catch
                    argsJson = "{}"
                End Try

                ' Determine if arguments should be escaped (template has quoted placeholder)
                Dim escapedArgsJson As String
                If callPartTemplate.Contains("""{arguments}""") Then
                    escapedArgsJson = EscapeJsonString(argsJson)
                Else
                    escapedArgsJson = argsJson
                End If

                ' Build the call part, also support {call} placeholder for raw call JSON
                Dim callPart As String = callPartTemplate _
                    .Replace("{call_id}", If(resp.CallId, "")) _
                    .Replace("{name}", If(resp.ToolName, "")) _
                    .Replace("{arguments}", escapedArgsJson) _
                    .Replace("{call}", resp.OriginalCallJson)

                If Not firstCall Then callParts.Append(",")
                callParts.Append(callPart)
                firstCall = False
            End If

            ' Build response content
            Dim responseContent As String = If(resp.Success, If(resp.Response, ""), $"Error: {resp.ErrorMessage}")

            ' Check if response should be escaped (template has quoted placeholder) or raw
            Dim finalResponseContent As String
            If responsePartTemplate.Contains("""{response}""") Then
                ' Template expects escaped string
                finalResponseContent = EscapeJsonString(responseContent)
            ElseIf responsePartTemplate.Contains("{response}") AndAlso
                   Not responsePartTemplate.Contains("""{response}") Then
                ' Template expects raw content - wrap in object if it's not valid JSON
                If IsValidJson(responseContent) Then
                    finalResponseContent = responseContent
                Else
                    ' Wrap plain text in a result object with escaped content
                    finalResponseContent = "{""result"": """ & EscapeJsonString(responseContent) & """}"
                End If
            Else
                finalResponseContent = EscapeJsonString(responseContent)
            End If

            Dim respPart As String = responsePartTemplate _
                .Replace("{call_id}", If(resp.CallId, "")) _
                .Replace("{name}", If(resp.ToolName, "")) _
                .Replace("{response}", finalResponseContent)

            If Not firstResp Then responseParts.Append(",")
            responseParts.Append(respPart)
            firstResp = False
        Next

        Dim functionCallsOutput As String = callParts.ToString()
        Dim responsesOutput As String = responseParts.ToString()

        ' Replace placeholders - NO comma manipulation by code
        ' Templates are responsible for their own structure
        Dim result As String = toolingModel.APICall_ToolResponses

        ' Simple replacement - if content exists, replace; if empty, remove placeholder
        result = result.Replace("{functioncalls}", functionCallsOutput)
        result = result.Replace("{responses}", responsesOutput)

        ' Clean up any empty structural remnants (empty arrays, double commas, etc.)
        ' This handles cases where one placeholder was empty
        result = Regex.Replace(result, "\[\s*\]", "[]")           ' Normalize empty arrays
        result = Regex.Replace(result, ",\s*,", ",")              ' Remove double commas
        result = Regex.Replace(result, "\[\s*,", "[")             ' Remove leading comma in array
        result = Regex.Replace(result, ",\s*\]", "]")             ' Remove trailing comma in array

        Return result
    End Function

    ''' <summary>
    ''' Determines whether a string represents a JSON object or array.
    ''' </summary>
    ''' <param name="str">Candidate JSON string.</param>
    ''' <returns>True if valid JSON object/array; otherwise False.</returns>
    Private Function IsValidJson(str As String) As Boolean
        If String.IsNullOrWhiteSpace(str) Then Return False
        str = str.Trim()
        If (str.StartsWith("{") AndAlso str.EndsWith("}")) OrElse
           (str.StartsWith("[") AndAlso str.EndsWith("]")) Then
            Try
                JToken.Parse(str)
                Return True
            Catch
                Return False
            End Try
        End If
        Return False
    End Function

    ''' <summary>
    ''' Escapes a string for safe embedding into a JSON string literal.
    ''' </summary>
    ''' <param name="str">Input string.</param>
    ''' <returns>Escaped string content (without surrounding quotes).</returns>
    Private Function EscapeJsonString(str As String) As String
        If String.IsNullOrEmpty(str) Then Return ""

        Dim sb As New StringBuilder()
        For Each c As Char In str
            Select Case c
                Case """"c : sb.Append("\""")
                Case "\"c : sb.Append("\\")
                Case "/"c : sb.Append("\/")
                Case ChrW(8) : sb.Append("\b")   ' Backspace
                Case ChrW(12) : sb.Append("\f")  ' Form feed
                Case vbLf(0) : sb.Append("\n")
                Case vbCr(0) : sb.Append("\r")
                Case vbTab(0) : sb.Append("\t")
                Case Else
                    If AscW(c) < 32 Then
                        ' Other control characters
                        sb.Append("\u" & AscW(c).ToString("X4"))
                    Else
                        sb.Append(c)
                    End If
            End Select
        Next
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' Determines whether a response contains tool calls by applying a detection regex pattern.
    ''' If <paramref name="detectionPattern"/> is empty, the pattern is derived from <c>INI_Response_2</c>.
    ''' </summary>
    ''' <param name="response">LLM response text.</param>
    ''' <param name="detectionPattern">Regex pattern used for detection.</param>
    ''' <returns>True if tool calls are detected; otherwise False.</returns>
    Public Function ContainsToolCalls(response As String, detectionPattern As String) As Boolean
        If String.IsNullOrWhiteSpace(response) Then Return False

        Dim pattern As String = detectionPattern
        If String.IsNullOrWhiteSpace(pattern) Then
            pattern = ExtractToolCallPatternFromResponse(INI_Response_2)
        End If

        If String.IsNullOrWhiteSpace(pattern) Then Return False

        Try
            Return Regex.IsMatch(response, pattern, RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Catch ex As Exception
            ToolingFileLogger.LogError("Regex match error.", details:=$"pattern='{pattern}'", ex:=ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Extracts a regex pattern embedded in a response key using the <c>ToolCallMatching*</c> markers.
    ''' </summary>
    ''' <param name="responseKey">Response configuration key string (e.g., <c>INI_Response_2</c>).</param>
    ''' <returns>Extracted regex pattern, or an empty string if not available/invalid.</returns>
    Private Function ExtractToolCallPatternFromResponse(responseKey As String) As String
        If String.IsNullOrEmpty(responseKey) Then
            Return String.Empty
        End If

        Dim startMarker As String = ToolCallMatchingStart
        Dim startIdx As Integer = responseKey.IndexOf(startMarker, StringComparison.OrdinalIgnoreCase)
        If startIdx < 0 Then Return String.Empty

        Dim endIdx As Integer = responseKey.IndexOf(ToolCallMatchingEnd, startIdx, StringComparison.OrdinalIgnoreCase)
        Dim triggerLen As Integer = If(endIdx >= 0,
               (endIdx - startIdx + ToolCallMatchingEnd.Length),
               (responseKey.Length - startIdx))

        Dim triggerText As String = responseKey.Substring(startIdx, triggerLen)

        Dim lt As Integer = triggerText.IndexOf("<"c)
        Dim gt As Integer = triggerText.LastIndexOf(">"c)

        Dim detectedPattern As String = String.Empty

        If lt >= 0 AndAlso gt > lt Then
            detectedPattern = triggerText.Substring(lt + 1, gt - lt - 1).Trim()
        Else
            Dim colonIdx As Integer = triggerText.IndexOf(ToolCallMatchingMiddle, StringComparison.OrdinalIgnoreCase)
            If colonIdx >= 0 Then
                Dim raw As String = triggerText.Substring(colonIdx + ToolCallMatchingMiddle.Length)
                Dim paren As Integer = raw.LastIndexOf(ToolCallMatchingEnd, StringComparison.OrdinalIgnoreCase)
                If paren >= 0 Then raw = raw.Substring(0, paren)
                detectedPattern = raw.Trim()
            End If
        End If

        If Not String.IsNullOrWhiteSpace(detectedPattern) Then
            Try
                Dim rx As New Regex(detectedPattern)
            Catch ex As ArgumentException
                ToolingFileLogger.LogError("Invalid regex pattern.", details:=$"pattern='{detectedPattern}'", ex:=ex)
                Return String.Empty
            End Try
        End If

        Return detectedPattern
    End Function

    ''' <summary>
    ''' Extracts tool calls from a JSON response according to a JSON "extraction map".
    ''' </summary>
    ''' <param name="response">Response text expected to parse as JSON.</param>
    ''' <param name="extractionMap">JSON map specifying paths for tool call array/id/name/arguments.</param>
    ''' <returns>List of extracted tool calls (may be empty).</returns>
    Public Function ExtractToolCalls(response As String, extractionMap As String) As List(Of ToolCall)
        Dim calls As New List(Of ToolCall)()

        If String.IsNullOrWhiteSpace(response) OrElse String.IsNullOrWhiteSpace(extractionMap) Then
            ToolingFileLogger.LogWarn(
                "ExtractToolCalls: Missing response or extractionMap.",
                details:=$"responseEmpty={String.IsNullOrWhiteSpace(response)}; extractionMapEmpty={String.IsNullOrWhiteSpace(extractionMap)}")
            Return calls
        End If

        Try
            Dim jResponse As JToken = JToken.Parse(response)
            Dim jMap As JObject = JObject.Parse(extractionMap)

            Dim arrayPath = If(jMap("array_path")?.ToString(), "")
            Dim callIdPath = If(jMap("call_id_path")?.ToString(), "id")
            Dim namePath = If(jMap("name_path")?.ToString(), "name")
            Dim argsPath = If(jMap("arguments_path")?.ToString(), "arguments")

            Dim toolCallTokens As IEnumerable(Of JToken)

            If Not String.IsNullOrWhiteSpace(arrayPath) Then
                toolCallTokens = jResponse.SelectTokens(arrayPath).ToList()
            Else
                toolCallTokens = {jResponse}
            End If

            For Each tcToken In toolCallTokens
                Try
                    Dim tc As New ToolCall() With {
                        .CallId = If(tcToken.SelectToken(callIdPath)?.ToString(), Guid.NewGuid().ToString("N")),
                        .ToolName = If(tcToken.SelectToken(namePath)?.ToString(), ""),
                        .RawJson = tcToken.ToString()
                    }

                    Dim argsToken = tcToken.SelectToken(argsPath)
                    If argsToken IsNot Nothing Then
                        If argsToken.Type = JTokenType.String Then
                            Try
                                tc.Arguments = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argsToken.ToString())
                            Catch ex As Exception
                                tc.Arguments = New Dictionary(Of String, Object) From {{"raw", argsToken.ToString()}}
                                ToolingFileLogger.LogWarn(
                                    "Arguments JSON string could not be deserialized; stored in 'raw'.",
                                    details:=$"ToolName='{tc.ToolName}'; CallId={tc.CallId}",
                                    ex:=ex)
                            End Try
                        Else
                            tc.Arguments = argsToken.ToObject(Of Dictionary(Of String, Object))()
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(tc.ToolName) Then
                        calls.Add(tc)
                    Else
                        ToolingFileLogger.LogWarn("ExtractToolCalls: Skipped tool call with empty ToolName.", details:=$"Raw={tc.RawJson}")
                    End If
                Catch ex As Exception
                    ToolingFileLogger.LogError("Error parsing individual tool call.", ex:=ex)
                    Debug.WriteLine($"Error parsing individual tool call: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            ToolingFileLogger.LogError("ExtractToolCalls error.", details:=$"extractionMap='{extractionMap}'", ex:=ex)
            Debug.WriteLine($"ExtractToolCalls error: {ex.Message}")
        End Try

        Return calls
    End Function

    ''' <summary>
    ''' Builds the tool instructions prompt appended to the tooling session's system prompt.
    ''' </summary>
    ''' <param name="selectedTools">Tools to include, sorted by <see cref="ModelConfig.ToolPriority"/>.</param>
    ''' <returns>System prompt fragment describing tooling usage and available tools.</returns>
    Public Function BuildToolInstructionsPrompt(selectedTools As List(Of ModelConfig)) As String
        Dim sb As New StringBuilder()

        MaxToolIterations = INI_ToolingMaximumIterations
        sb.AppendLine(InterpolateAtRuntime(SP_Add_Tooling))
        sb.AppendLine()
        sb.AppendLine("Available tools:")

        If selectedTools Is Nothing Then
            Return sb.ToString()
        End If

        Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

        For Each tool In sortedTools
            If Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                sb.AppendLine()
                sb.AppendLine($"- {tool.ToolInstructionsPrompt}")
            End If
        Next

        Return sb.ToString()
    End Function

    Private Function BuildKnowledgeToolStoreInventoryLine() As String
        Dim storeLabels As New List(Of String)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If store Is Nothing Then Continue For

                Dim label As String = KnowledgeStoreCatalog.GetDisplayLabel(store)
                If String.IsNullOrWhiteSpace(label) Then Continue For

                Try
                    Dim manifest = KnowledgeStoreManifest.Load(store)
                    If manifest IsNot Nothing AndAlso manifest.Entries IsNot Nothing AndAlso manifest.Entries.Count > 0 Then
                        label &= $" ({manifest.Entries.Count} docs)"
                    End If
                Catch
                End Try

                If Not storeLabels.Any(Function(x) String.Equals(x, label, StringComparison.OrdinalIgnoreCase)) Then
                    storeLabels.Add(label)
                End If
            Next
        Catch
        End Try

        If storeLabels.Count = 0 Then
            Return ""
        End If

        Return "Knowledge stores currently available: " & String.Join(", ", storeLabels) & "."
    End Function

    Private Function GetAvailableKnowledgeStoreNames() As List(Of String)
        Dim storeNames As New List(Of String)()

        Try
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)

            For Each store In stores
                If store Is Nothing Then Continue For

                Dim displayLabel = KnowledgeStoreCatalog.GetDisplayLabel(store)
                If Not String.IsNullOrWhiteSpace(displayLabel) Then
                    If Not storeNames.Any(Function(x) String.Equals(x, displayLabel, StringComparison.OrdinalIgnoreCase)) Then
                        storeNames.Add(displayLabel)
                    End If
                End If

                Dim plainName As String = If(store.Name, "").Trim()
                If plainName <> "" Then
                    If Not storeNames.Any(Function(x) String.Equals(x, plainName, StringComparison.OrdinalIgnoreCase)) Then
                        storeNames.Add(plainName)
                    End If
                End If
            Next
        Catch
        End Try

        Return storeNames
    End Function

    Private Function BuildInternalKnowledgeToolDefinition() As String
        Dim storeInventory As String = BuildKnowledgeToolStoreInventoryLine()
        Dim descriptionSuffix As String = If(String.IsNullOrWhiteSpace(storeInventory), "", " " & storeInventory)

        Dim definition As New JObject(
        New JProperty("name", InternalKnowledgeToolName),
        New JProperty("description",
            "Searches the user's local knowledge stores (their own curated document collections). " &
            "This tool mirrors the Freestyle knowledge trigger functionality. " &
            "You can either use structured arguments (query/store/tag/max_results) or pass the exact Freestyle trigger syntax via raw_trigger. " &
            "Use this for the user's own materials, not for public-web lookup." & descriptionSuffix),
        New JProperty("parameters",
            New JObject(
                New JProperty("type", "object"),
                New JProperty("properties",
                    New JObject(
                        New JProperty("query",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional natural-language knowledge-store query. If omitted, the tool can still search broadly or within a given store/tag scope.")
                            )
                        ),
                        New JProperty("store",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional knowledge store name. Use exactly one of the store names exposed in the tool instructions.")
                            )
                        ),
                        New JProperty("tag",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional tag filter, equivalent to Freestyle syntax 'tag:YourTag'.")
                            )
                        ),
                        New JProperty("raw_trigger",
                            New JObject(
                                New JProperty("type", "string"),
                                New JProperty("description", "Optional exact Freestyle-style trigger. Examples: '(kb)', '(kb:termination without notice)', '(kb:store:Policies confidentiality)', '(kb:tag:NDA confidentiality)'. If supplied, this takes precedence over query/store/tag.")
                            )
                        ),
                        New JProperty("max_results",
                            New JObject(
                                New JProperty("type", "integer"),
                                New JProperty("description", "Optional maximum number of results to retrieve. Best effort; the resolver may still enforce its own cap.")
                            )
                        )
                    )
                ),
                New JProperty("additionalProperties", False)
            )
        )
    )

        Return definition.ToString(Formatting.None)
    End Function

    Private Function BuildInternalKnowledgeToolInstructionsPrompt() As String
        Dim kbTrigger As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTrigger
        Dim kbTriggerPrefix As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTriggerPrefix

        Dim sb As New StringBuilder()
        sb.Append("knowledge_search: Searches the user's local knowledge stores — the user's own curated document collections such as contracts, policies, briefs, manuals, templates, emails, and reference material. ")
        sb.Append("This tool supports the same search semantics as the Freestyle knowledge trigger. ")
        sb.Append("Prefer structured arguments for normal calls: query (optional), store (optional), tag (optional), max_results (optional). ")
        sb.Append("If you need exact parity with Freestyle, pass raw_trigger using the literal syntax. ")
        sb.Append($"Valid trigger forms include '{kbTrigger}', '{kbTriggerPrefix}your query)', '{kbTriggerPrefix}store:StoreName your query)', '{kbTriggerPrefix}tag:TagName your query)', and combinations such as '{kbTriggerPrefix}store:StoreName tag:TagName your query)'. ")
        sb.Append("If store is omitted, all stores are searched. If query is omitted but store and/or tag is provided, the tool still performs a scoped retrieval. If everything is omitted, it performs a broad cross-store retrieval. ")

        Dim storeInventory As String = BuildKnowledgeToolStoreInventoryLine()
        If Not String.IsNullOrWhiteSpace(storeInventory) Then
            sb.Append(storeInventory & " ")
            sb.Append("Use the store names exactly as listed. ")
        End If

        sb.Append("Do NOT use this tool for public information or general knowledge — use your own knowledge or internet_search for that. ")
        sb.Append("When citing results, mention the document name and store name.")

        Return sb.ToString()
    End Function

    Private Function BuildKnowledgeToolTrigger(query As String, storeName As String, tagName As String, rawTrigger As String) As String
        Dim kbTrigger As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTrigger
        Dim kbTriggerPrefix As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTriggerPrefix

        ' raw_trigger takes precedence — normalize and return as-is
        If Not String.IsNullOrWhiteSpace(rawTrigger) Then
            Dim normalized As String = rawTrigger.Trim()

            If String.Equals(normalized, kbTrigger, StringComparison.OrdinalIgnoreCase) Then
                Return kbTrigger
            End If

            If normalized.StartsWith(kbTriggerPrefix, StringComparison.OrdinalIgnoreCase) Then
                If Not normalized.EndsWith(")") Then
                    normalized &= ")"
                End If
                Return normalized
            End If

            If normalized.StartsWith("(kb", StringComparison.OrdinalIgnoreCase) Then
                If Not normalized.EndsWith(")") Then
                    normalized &= ")"
                End If
                Return normalized
            End If

            Return kbTriggerPrefix & normalized.TrimEnd(")"c).Trim() & ")"
        End If

        ' Query-only (no store, no tag): use bare (kb) trigger.
        ' The search query will be handled separately via KnowledgeQueryService
        ' semantic search in ExecuteInternalKnowledgeTool, NOT embedded in the trigger.
        ' This matches Freestyle behavior where "Nudging (kb)" loads all docs
        ' and the query is used as the LLM prompt, not as a metadata filter.
        If Not String.IsNullOrWhiteSpace(query) AndAlso
       String.IsNullOrWhiteSpace(storeName) AndAlso
       String.IsNullOrWhiteSpace(tagName) Then

            Dim normalizedQuery As String = query.Trim()

            ' If the query already IS or contains a trigger, return it directly
            If String.Equals(normalizedQuery, kbTrigger, StringComparison.OrdinalIgnoreCase) Then
                Return kbTrigger
            End If

            If normalizedQuery.IndexOf(kbTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return normalizedQuery
            End If

            If normalizedQuery.IndexOf(kbTriggerPrefix, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return normalizedQuery
            End If

            If normalizedQuery.StartsWith("(kb", StringComparison.OrdinalIgnoreCase) Then
                If Not normalizedQuery.EndsWith(")") Then
                    normalizedQuery &= ")"
                End If
                Return normalizedQuery
            End If

            ' Query-only: return bare (kb) — semantic search handles the query separately
            Return kbTrigger
        End If

        ' Store and/or tag specified: build parameterized trigger
        Dim parts As New List(Of String)()

        If Not String.IsNullOrWhiteSpace(storeName) Then
            parts.Add("store:" & storeName.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(tagName) Then
            parts.Add("tag:" & tagName.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(query) Then
            parts.Add(query.Trim())
        End If

        If parts.Count = 0 Then
            Return kbTrigger
        End If

        Return kbTriggerPrefix & String.Join(" ", parts).Trim() & ")"
    End Function


    Private Function GetToolArgumentString(arguments As Dictionary(Of String, Object), key As String) As String
        If arguments Is Nothing OrElse Not arguments.ContainsKey(key) OrElse arguments(key) Is Nothing Then
            Return ""
        End If

        Dim value = arguments(key)

        If TypeOf value Is JValue Then
            Return DirectCast(value, JValue).ToString().Trim()
        End If

        Return value.ToString().Trim()
    End Function

    Private Sub TrySetLateBoundProperty(target As Object, propertyName As String, value As Object)
        If target Is Nothing Then Return

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanWrite Then Return

            Dim convertedValue As Object = value

            If value IsNot Nothing Then
                Dim targetType = If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType)
                convertedValue = System.Convert.ChangeType(value, targetType, Globalization.CultureInfo.InvariantCulture)
            End If

            prop.SetValue(target, convertedValue, Nothing)
        Catch
        End Try
    End Sub

    Private Function TryGetLateBoundString(target As Object, propertyName As String) As String
        If target Is Nothing Then
            Return ""
        End If

        Try
            Dim prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.IgnoreCase)

            If prop Is Nothing OrElse Not prop.CanRead Then
                Return ""
            End If

            Dim value = prop.GetValue(target, Nothing)
            Return If(value, "").ToString()
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Creates a built-in internal web retrieval tool configuration as a <see cref="ModelConfig"/>.
    ''' </summary>
    ''' <returns>Internal tool configuration.</returns>
    Public Function GetInternalWebTool() As ModelConfig
        Return New ModelConfig() With {
            .ToolName = InternalWebToolName,
            .ToolInstructionsPrompt = InternalWebToolInstructionsPrompt,
            .ToolDefinition = InternalWebToolDefinition,
            .ModelDescription = "Web Content Retriever" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 999,
            .ToolErrorHandling = "skip"
        }
    End Function

    ''' <summary>
    ''' Creates a built-in internal internet search tool configuration as a <see cref="ModelConfig"/>.
    ''' Only meaningful when <c>INI_ISearch</c> is enabled and <c>INI_ISearch_URL</c> is configured.
    ''' </summary>
    ''' <param name="enforcePrivacy">When True, privacy constraints are included in the tool definition and instructions.</param>
    ''' <returns>Internal search tool configuration.</returns>
    Public Function GetInternalSearchTool(Optional enforcePrivacy As Boolean = True) As ModelConfig
        Dim definition As String = InternalSearchToolDefinition
        Dim instructions As String = InternalSearchToolInstructionsPrompt

        If Not enforcePrivacy Then
            ' Strip privacy constraints from the tool definition
            definition =
                "{""name"":""internet_search""," &
                """description"":""Searches the internet via the configured search engine, retrieves the top result pages, and returns their readable text content. Use this when you need up-to-date or factual information you are not confident about.""," &
                """parameters"":{""type"":""object"",""properties"":{" &
                """query"":{""type"":""string"",""description"":""The search query.""}," &
                """max_results"":{""type"":""integer"",""description"":""Maximum number of search result pages to retrieve (default: 4, server-capped).""}," &
                """max_depth"":{""type"":""integer"",""description"":""Maximum crawl depth per result page. 0 = top-level only (default: 0, server-capped).""}},""required"":[""query""]}}"

            instructions =
                "internet_search: Searches the internet and returns readable text from the top result pages. " &
                "Call this tool when you need current or factual information you are not confident about. " &
                "Provide query (required string). Optionally provide max_results (integer, default 4) and max_depth (integer, default 0). " &
                "Return value includes the search query used, the URLs visited, and the page content for each qualifying result."
        End If

        Return New ModelConfig() With {
            .ToolName = InternalSearchToolName,
            .ToolInstructionsPrompt = instructions,
            .ToolDefinition = definition,
            .ModelDescription = "Internet Search (" & If(Not String.IsNullOrWhiteSpace(INI_ISearch_Name), INI_ISearch_Name, "Search") & ")" & InternalToolSuffix,
            .Tool = True,
            .ToolPriority = 998,
            .ToolErrorHandling = "skip"
        }
    End Function


    ''' <summary>
    ''' Creates a built-in internal knowledge store search tool configuration as a <see cref="ModelConfig"/>.
    ''' Only meaningful when <c>INI_KnowledgeStorePath</c> or <c>INI_KnowledgeStorePathLocal</c> is configured
    ''' and at least one knowledge store has an indexed manifest.
    ''' </summary>
    ''' <returns>Internal knowledge search tool configuration.</returns>
    Public Function GetInternalKnowledgeTool() As ModelConfig
        Return New ModelConfig() With {
        .ToolName = InternalKnowledgeToolName,
        .ToolInstructionsPrompt = BuildInternalKnowledgeToolInstructionsPrompt(),
        .ToolDefinition = BuildInternalKnowledgeToolDefinition(),
        .ModelDescription = "Knowledge Store Search" & InternalToolSuffix,
        .Tool = True,
        .ToolPriority = 997,
        .ToolErrorHandling = "skip"
    }
    End Function

    Private Sub LogAgentToolCallStatistic(toolName As String)
        Dim surface As String = ""

        If _apActive Then
            surface = "Outlook_AutoPilot"
        ElseIf _chatAgentActive Then
            surface = "Outlook_LocalChatAgent"
        Else
            Return
        End If

        SharedLogger.LogAgentToolCall(_context, _context.RDV, surface, toolName)
    End Sub

    ''' <summary>
    ''' Executes a single tool call using an internal tool implementation or an external tool configuration.
    ''' Internal tools: <c>retrieve_web_content</c> and <c>internet_search</c> (when search is enabled).
    ''' </summary>
    Public Async Function ExecuteToolCall(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)

        LogAgentToolCallStatistic(toolCall.ToolName)

        ' ── Local Chat Agent workspace tools; never available to AutoPilot ──
        If _chatAgentActive AndAlso Not _apActive AndAlso IsChatAgentWorkspaceTool(toolCall.ToolName) Then
            Dim workspaceResult = Await ExecuteChatAgentWorkspaceTool(toolCall, context, cancellationToken)
            Return workspaceResult
        End If

        ' ── AutoPilot / Chat Agent internal tool routing ──
        If (_apActive OrElse _chatAgentActive) AndAlso IsAutoPilotInternalTool(toolCall.ToolName) Then
            Dim apResult = Await TryExecuteAutoPilotTool(toolCall, context, cancellationToken)
            If apResult IsNot Nothing Then
                ' Record for "Sources used:" footer
                If _apCurrentToolCallLog IsNot Nothing Then
                    Dim elapsed = DateTime.Now - apResult.Timestamp
                    Dim excerpt = BuildResultExcerpt(apResult.Response, 80)
                    RecordAutoPilotToolCall(
                        toolCall.ToolName,
                        If(toolConfig?.ModelDescription, toolCall.ToolName),
                        BuildCondensedParamSummary(toolCall.Arguments),
                        isInternalTool:=True,
                        wasSuccessful:=apResult.Success,
                        resultExcerpt:=excerpt,
                        elapsed:=elapsed,
                        urls:=Nothing)
                End If

                Return apResult
            End If
        End If

        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        ' Build condensed parameter summary for log window
        Dim paramSummary As String = BuildCondensedParamSummary(toolCall.Arguments)
        context.Log($"Executing tool: {toolCall.ToolName}{paramSummary}")


        Try
            ' Check cancellation before execution
            cancellationToken.ThrowIfCancellationRequested()

            If toolCall.ToolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalWebTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf toolCall.ToolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) Then
                response = Await ExecuteInternalSearchTool(toolCall, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf IsInternalKnowledgeToolName(toolCall.ToolName) Then
                response = Await ExecuteInternalKnowledgeTool(toolCall, context)
                ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)

            ElseIf SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolCall.ToolName) Then
                ' M365 tools are allowed in interactive/local scenarios, but never in AutoPilot.
                ' They may trigger an interactive MSAL sign-in on the user's machine.
                If Not _apActive Then
                    response = Await ExecuteInternalM365Tool(toolCall, context)
                    ToolingFileLogger.LogRawResponseStub($"Internal tool ({toolCall.ToolName})", response.Response)
                Else
                    response.Success = False
                    response.ErrorMessage = "M365 tools cannot run inside AutoPilot because they may require interactive user sign-in."
                    ToolingFileLogger.LogWarn(
                        "M365 tool blocked in AutoPilot.",
                        details:=$"tool={toolCall.ToolName}; _apActive={_apActive}; _chatAgentActive={_chatAgentActive}")
                End If

            Else
                response = Await ExecuteExternalTool(toolCall, toolConfig, context, cancellationToken)
                ToolingFileLogger.LogRawResponseStub($"Tool LLM() ({toolCall.ToolName})", response.Response)
            End If

            ' Log completion with excerpt
            If response.Success Then
                Dim resultSummary As String = BuildResultExcerpt(response.Response, 80)
                context.Log($"Tool {toolCall.ToolName} completed: {resultSummary}", "success")
            Else
                context.Log($"Tool {toolCall.ToolName} failed: {response.ErrorMessage}", "error")
            End If

            ' ── AutoPilot / Chat Agent: record tool call for dashboard + sources footer ──
            If (_apActive OrElse _chatAgentActive) AndAlso _apCurrentToolCallLog IsNot Nothing Then
                Dim elapsed = DateTime.Now - response.Timestamp
                Dim excerpt = BuildResultExcerpt(response.Response, 80)

                ' Log to dashboard (ApDashboardLog routes to AutoPilot OR Chat Agent automatically)
                ApDashboardLog($"🔧 External tool: {toolCall.ToolName}{paramSummary}", "info")
                If response.Success Then
                    ApDashboardLog($"   ✓ {excerpt}", "info")
                Else
                    ApDashboardLog($"   ✗ {If(response.ErrorMessage, excerpt)}", "error")
                End If

                ' Record for "Sources used:" footer
                Dim toolUrls As List(Of String) = Nothing
                If toolCall.Arguments IsNot Nothing AndAlso toolCall.Arguments.ContainsKey("urls") Then
                    Try
                        Dim urlsObj = toolCall.Arguments("urls")
                        If TypeOf urlsObj Is JArray Then
                            toolUrls = DirectCast(urlsObj, JArray).Select(Function(t) t.ToString()).ToList()
                        ElseIf TypeOf urlsObj Is IEnumerable(Of Object) Then
                            toolUrls = DirectCast(urlsObj, IEnumerable(Of Object)).Select(Function(o) o.ToString()).ToList()
                        ElseIf TypeOf urlsObj Is String Then
                            toolUrls = New List(Of String) From {CStr(urlsObj)}
                        End If
                    Catch
                    End Try
                End If
                RecordAutoPilotToolCall(
                    toolCall.ToolName,
                    If(toolConfig?.ModelDescription, toolCall.ToolName),
                    paramSummary,
                    isInternalTool:=False,
                    wasSuccessful:=response.Success,
                    resultExcerpt:=excerpt,
                    elapsed:=elapsed,
                    urls:=toolUrls)
            End If

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            context.Log($"Tool {toolCall.ToolName} cancelled")
            ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            context.Log($"Tool {toolCall.ToolName} error: {ex.Message}")
            ToolingFileLogger.LogError($"Tool {toolCall.ToolName} execution error.", ex:=ex)
        End Try

        Return response
    End Function


    ''' <summary>
    ''' Builds a brief excerpt of the tool result for display in the log window.
    ''' </summary>
    ''' <param name="result">Full tool response text.</param>
    ''' <param name="maxExcerptLength">Maximum length for the excerpt portion.</param>
    ''' <returns>Formatted string like "12,345 chars: 'The quick brown fox...'".</returns>
    Private Function BuildResultExcerpt(result As String, Optional maxExcerptLength As Integer = 80) As String
        If String.IsNullOrEmpty(result) Then
            Return "0 chars (empty)"
        End If

        Dim charCount As Integer = result.Length
        Dim formattedCount As String = charCount.ToString("N0")

        ' Clean up the result for excerpt (remove excessive whitespace/newlines)
        Dim cleaned As String = Regex.Replace(result, "\s+", " ").Trim()

        If cleaned.Length <= maxExcerptLength Then
            Return $"{formattedCount} chars: '{cleaned}'"
        End If

        ' Truncate and add ellipsis
        Dim excerpt As String = cleaned.Substring(0, maxExcerptLength - 3) & "..."
        Return $"{formattedCount} chars: '{excerpt}'"
    End Function

    ''' <summary>
    ''' Builds a condensed parameter summary for display in the log window.
    ''' </summary>
    ''' <param name="arguments">Tool call arguments dictionary.</param>
    ''' <param name="maxLength">Maximum length for each parameter value display.</param>
    ''' <returns>Formatted parameter string like " (query: 'search term', count: 10)".</returns>
    Private Function BuildCondensedParamSummary(arguments As Dictionary(Of String, Object), Optional maxLength As Integer = 50) As String
        If arguments Is Nothing OrElse arguments.Count = 0 Then
            Return ""
        End If

        Dim parts As New List(Of String)()

        For Each kvp In arguments
            Dim valueStr As String = ""
            If kvp.Value IsNot Nothing Then
                If TypeOf kvp.Value Is JArray Then
                    Dim arr = DirectCast(kvp.Value, JArray)
                    valueStr = $"[{arr.Count} items]"
                ElseIf TypeOf kvp.Value Is IEnumerable(Of Object) AndAlso Not TypeOf kvp.Value Is String Then
                    valueStr = $"[{DirectCast(kvp.Value, IEnumerable(Of Object)).Count()} items]"
                Else
                    valueStr = kvp.Value.ToString()
                    ' Use shorter limit for long text parameters like "instruction"
                    Dim effectiveMax = If(valueStr.Length > 200, Math.Min(maxLength, 80), maxLength)
                    If valueStr.Length > effectiveMax Then
                        valueStr = valueStr.Substring(0, effectiveMax - 3) & "..."
                    End If
                End If
            End If

            parts.Add($"{kvp.Key}: '{valueStr}'")
        Next

        Return $" ({String.Join(", ", parts)})"
    End Function


    ''' <summary>
    ''' Executes the internal web retrieval tool by fetching content for one or more URLs and returning tagged content blocks.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>url</c> or <c>urls</c> arguments.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing retrieved content or an error.</returns>
    Private Async Function ExecuteInternalWebTool(toolCall As ToolCall, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            ' Check cancellation before starting
            cancellationToken.ThrowIfCancellationRequested()

            Dim urls As New List(Of String)()

            If toolCall.Arguments.ContainsKey("urls") Then
                Dim urlsArg = toolCall.Arguments("urls")
                If TypeOf urlsArg Is JArray Then
                    For Each item In DirectCast(urlsArg, JArray)
                        urls.Add(item.ToString())
                    Next
                ElseIf TypeOf urlsArg Is IEnumerable(Of Object) Then
                    For Each item In DirectCast(urlsArg, IEnumerable(Of Object))
                        urls.Add(item.ToString())
                    Next
                ElseIf TypeOf urlsArg Is String Then
                    urls.Add(urlsArg.ToString())
                End If
            ElseIf toolCall.Arguments.ContainsKey("url") Then
                urls.Add(toolCall.Arguments("url").ToString())
            End If

            If urls.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No URLs provided"
                ToolingFileLogger.LogWarn("Internal web tool: No URLs provided.", details:=$"CallId={toolCall.CallId}; Args={JsonConvert.SerializeObject(toolCall.Arguments)}")
                Return response
            End If

            ' ── SharePoint / OneDrive / Teams detection ──
            ' These authenticated cloud storage URLs require login and will not return useful content.
            Dim sharepointPatterns As String() = {"sharepoint.com", "onedrive.com", "1drv.ms", "teams.microsoft.com", ":f:/", "/:f:/"}
            Dim blockedUrls As New List(Of String)()
            For Each url In urls
                Dim lowerUrl = url.ToLowerInvariant()
                For Each pattern In sharepointPatterns
                    If lowerUrl.Contains(pattern) Then
                        blockedUrls.Add(url)
                        Exit For
                    End If
                Next
            Next

            If blockedUrls.Count > 0 Then
                Dim blockedList = String.Join(", ", blockedUrls)
                response.Success = False
                response.ErrorMessage =
                    $"Cannot retrieve content from the following URL(s) because they point to SharePoint, OneDrive, or Microsoft Teams — " &
                    $"these are authenticated cloud storage resources that require login and cannot be accessed remotely: {blockedList}. " &
                    "Please ask the user to download the file(s) and attach them directly to the e-mail."
                context.Log($"Blocked SharePoint/OneDrive URL(s): {blockedList}", "warn")
                ToolingFileLogger.LogWarn("Internal web tool: SharePoint/OneDrive URL blocked.", details:=$"urls={blockedList}")
                Return response
            End If

            context.Log($"Retrieving content from {urls.Count} URL(s)...")

            Dim results As New StringBuilder()

            If UseWebView2 Then
                For i = 0 To urls.Count - 1
                    ' Check cancellation before each URL fetch
                    cancellationToken.ThrowIfCancellationRequested()

                    Dim url = urls(i)
                    Try
                        context.Log($"  Fetching: {url}")
                        ' RetrieveWebsiteContent_WebView2 handles its own threading (runs on STA thread internally)
                        Dim content = Await RetrieveWebsiteContent_WebView2(url, 0)


                        If Not String.IsNullOrWhiteSpace(content) Then
                            results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                            results.AppendLine($"<CONTENT_{i + 1}>")
                            results.AppendLine(content)
                            results.AppendLine($"</CONTENT_{i + 1}>")
                            results.AppendLine()
                        Else
                            results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                            results.AppendLine($"<CONTENT_{i + 1}>No content retrieved</CONTENT_{i + 1}>")
                            results.AppendLine()
                            ToolingFileLogger.LogWarn("Internal web tool: No content retrieved.", details:=$"url={url}")
                        End If
                    Catch ex As OperationCanceledException
                        Throw ' Re-throw cancellation
                    Catch ex As Exception
                        results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                        results.AppendLine($"<ERROR_{i + 1}>{ex.Message}</ERROR_{i + 1}>")
                        results.AppendLine()
                        ToolingFileLogger.LogError("Internal web tool fetch error.", details:=$"url={url}", ex:=ex)
                    End Try
                Next
            Else
                Using httpClient As New System.Net.Http.HttpClient()
                    httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                    httpClient.Timeout = TimeSpan.FromSeconds(30)

                    For i = 0 To urls.Count - 1
                        ' Check cancellation before each URL fetch
                        cancellationToken.ThrowIfCancellationRequested()

                        Dim url = urls(i)
                        Try
                            context.Log($"  Fetching: {url}")
                            Dim content = Await RetrieveWebsiteContent(url, INI_ISearch_MaxDepth, httpClient)

                            If Not String.IsNullOrWhiteSpace(content) Then
                                results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                                results.AppendLine($"<CONTENT_{i + 1}>")
                                results.AppendLine(content)
                                results.AppendLine($"</CONTENT_{i + 1}>")
                                results.AppendLine()
                            Else
                                results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                                results.AppendLine($"<CONTENT_{i + 1}>No content retrieved</CONTENT_{i + 1}>")
                                results.AppendLine()
                                ToolingFileLogger.LogWarn("Internal web tool: No content retrieved.", details:=$"url={url}")
                            End If
                        Catch ex As OperationCanceledException
                            Throw ' Re-throw cancellation
                        Catch ex As Exception
                            results.AppendLine($"<URL_{i + 1}>{url}</URL_{i + 1}>")
                            results.AppendLine($"<ERROR_{i + 1}>{ex.Message}</ERROR_{i + 1}>")
                            results.AppendLine()
                            ToolingFileLogger.LogError("Internal web tool fetch error.", details:=$"url={url}", ex:=ex)
                        End Try
                    Next
                End Using
            End If

            response.Response = results.ToString()
            response.Success = True

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn("Internal web tool cancelled.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal web tool error.", ex:=ex)
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Executes the internal internet search tool by querying the configured search engine,
    ''' extracting result URLs via response masks, fetching qualifying page content, and returning
    ''' tagged result blocks including the search query and all visited URLs for transparency.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>query</c> and optional <c>max_results</c>/<c>max_depth</c>.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing search results or an error.</returns>
    Private Async Function ExecuteInternalSearchTool(toolCall As ToolCall, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            cancellationToken.ThrowIfCancellationRequested()

            ' ── Validate search configuration ────────────────────────────
            If Not INI_ISearch OrElse String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
                response.Success = False
                response.ErrorMessage = "Internet search is not configured or not enabled."
                ToolingFileLogger.LogWarn("Internal search tool: search not enabled/configured.",
                    details:=$"INI_ISearch={INI_ISearch}; INI_ISearch_URL='{INI_ISearch_URL}'")
                Return response
            End If

            ' ── Extract and validate parameters ──────────────────────────
            Dim query As String = ""
            If toolCall.Arguments.ContainsKey("query") Then
                query = If(toolCall.Arguments("query")?.ToString(), "").Trim()
            End If

            If String.IsNullOrWhiteSpace(query) Then
                response.Success = False
                response.ErrorMessage = "No search query provided."
                ToolingFileLogger.LogWarn("Internal search tool: empty query.",
                    details:=$"CallId={toolCall.CallId}; Args={JsonConvert.SerializeObject(toolCall.Arguments)}")
                Return response
            End If

            ' ── PII / confidential data safety net ───────────────────────
            ' Block queries that contain obvious personal data patterns.
            ' Only enforced when privacy protection is enabled.
            Dim enforcePrivacyForPII As Boolean = True  ' default for non-AutoPilot callers
            If _apConfig IsNot Nothing Then
                enforcePrivacyForPII = _apConfig.EnablePrivacyProtection
            ElseIf _context IsNot Nothing Then
                enforcePrivacyForPII = _context.INI_EnablePrivacyForSearch
            End If

            If enforcePrivacyForPII Then
                Dim piiPatterns As String() = {
                "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b",
                "\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b",
                "\b\+?\d{1,3}[-.\s]?\(?\d{1,4}\)?[-.\s]?\d{2,4}[-.\s]?\d{2,4}[-.\s]?\d{0,4}\b",
                "\b\d{3}-\d{2}-\d{4}\b",
                "\b\d{2}[\./]\d{2}[\./]\d{2,4}\b(?=.*\d{2}[\./]\d{2}[\./]\d{2,4})",
                "\b[A-Z]{2}\d{2}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{0,2}\b",
                "\b(?:4\d{3}|5[1-5]\d{2}|6011|3[47]\d{2})[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b",
                "\bAHV[\s-]?\d{3}[\.\s]?\d{4}[\.\s]?\d{4}[\.\s]?\d{2}\b"
            }
                For Each piiPattern In piiPatterns
                    If Regex.IsMatch(query, piiPattern, RegexOptions.IgnoreCase) Then
                        response.Success = False
                        response.ErrorMessage = "Search query blocked: appears to contain personal or confidential data."
                        ToolingFileLogger.LogWarn("Internal search tool: query blocked by PII filter.",
                        details:=$"CallId={toolCall.CallId}; Pattern='{piiPattern}'")
                        context.Log("  ⚠ Search query blocked — contains data that appears personal or confidential.", "warn")
                        Return response
                    End If
                Next
            End If

            ' Clamp max_results to server limit (INI_ISearch_Tries)
            Dim maxResults As Integer = INI_ISearch_Results
            If toolCall.Arguments.ContainsKey("max_results") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_results")?.ToString(), requested) AndAlso requested > 0 Then
                    maxResults = Math.Min(requested, INI_ISearch_Tries)
                End If
            End If

            ' Clamp max_depth to server limit (INI_ISearch_MaxDepth)
            Dim maxDepth As Integer = 0
            If toolCall.Arguments.ContainsKey("max_depth") Then
                Dim requested As Integer
                If Integer.TryParse(toolCall.Arguments("max_depth")?.ToString(), requested) AndAlso requested >= 0 Then
                    maxDepth = Math.Min(requested, INI_ISearch_MaxDepth)
                End If
            End If

            context.Log($"Internet search: query='{query}', max_results={maxResults}, max_depth={maxDepth}")
            ToolingFileLogger.LogStep($"Search query: '{query}'; max_results={maxResults}; max_depth={maxDepth}; engine={INI_ISearch_Name}")

            ' ── Perform the HTTP search request ──────────────────────────
            Dim searchUrl As String = INI_ISearch_URL & Uri.EscapeDataString(query)
            context.Log($"  Search URL: {searchUrl}")

            cancellationToken.ThrowIfCancellationRequested()

            Dim searchResponse As String = ""
            Using httpClient As New System.Net.Http.HttpClient()
                httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36")
                httpClient.Timeout = TimeSpan.FromSeconds(30)

                searchResponse = Await httpClient.GetStringAsync(searchUrl)
            End Using

            If String.IsNullOrWhiteSpace(searchResponse) Then
                response.Success = False
                response.ErrorMessage = "Search engine returned an empty response."
                ToolingFileLogger.LogWarn("Internal search tool: empty search response.",
                    details:=$"searchUrl={searchUrl}")
                Return response
            End If

            ' ── Extract unique URLs using response masks ─────────────────
            Dim urlPattern As String = Regex.Escape(INI_ISearch_ResponseMask1) & "(.*?)" & Regex.Escape(INI_ISearch_ResponseMask2)
            Dim matches As MatchCollection = Regex.Matches(searchResponse, urlPattern)

            Dim extractedUrls As New List(Of String)()
            For Each m As Match In matches
                Dim rawUrl As String = m.Groups(1).Value
                Dim decodedUrl As String = System.Net.WebUtility.UrlDecode(rawUrl.Replace(INI_ISearch_ResponseMask1, ""))

                If Not extractedUrls.Contains(decodedUrl) AndAlso IsSafeWebUrl(decodedUrl) Then
                    extractedUrls.Add(decodedUrl)
                End If

                If extractedUrls.Count >= INI_ISearch_Tries Then Exit For
            Next

            context.Log($"  Extracted {extractedUrls.Count} unique URL(s) from search results")
            ToolingFileLogger.LogStep($"Extracted URLs: {extractedUrls.Count}")

            If extractedUrls.Count = 0 Then
                response.Success = False
                response.ErrorMessage = "No result URLs could be extracted from the search engine response."
                ToolingFileLogger.LogWarn("Internal search tool: no URLs extracted.",
                    details:=$"searchUrl={searchUrl}; ResponseMask1='{INI_ISearch_ResponseMask1}'; ResponseMask2='{INI_ISearch_ResponseMask2}'")
                Return response
            End If

            ' ── Fetch content from each result URL ───────────────────────
            Dim results As New Text.StringBuilder()
            Dim visitedUrls As New List(Of String)()
            Dim resultIndex As Integer = 0

            ' Header: report the search query and engine
            results.AppendLine($"<SEARCH_QUERY>{query}</SEARCH_QUERY>")
            results.AppendLine($"<SEARCH_ENGINE>{If(INI_ISearch_Name, "Search")}</SEARCH_ENGINE>")
            results.AppendLine()

            For Each url In extractedUrls
                If resultIndex >= maxResults Then Exit For
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    context.Log($"  Fetching result: {url}")
                    visitedUrls.Add(url)

                    Dim content As String = ""

                    If UseWebView2 Then
                        content = Await RetrieveWebsiteContent_WebView2(url, ISearch_MaxChars)
                    Else
                        Using httpClient As New System.Net.Http.HttpClient()
                            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
                            httpClient.Timeout = TimeSpan.FromSeconds(30)
                            content = Await RetrieveWebsiteContent(url, maxDepth, httpClient)
                        End Using
                    End If

                    ' Apply character cap (ISearch_MaxChars) for WebView2 results that exceed it
                    If Not String.IsNullOrWhiteSpace(content) AndAlso ISearch_MaxChars > 0 AndAlso content.Length > ISearch_MaxChars Then
                        content = content.Substring(0, ISearch_MaxChars)
                    End If

                    ' Discard noise (pages shorter than ISearch_MinChars)
                    If Not String.IsNullOrWhiteSpace(content) AndAlso content.Length >= ISearch_MinChars Then
                        resultIndex += 1
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}_URL>{url}</SEARCHRESULT_{resultIndex}_URL>")
                        results.AppendLine($"<SEARCHRESULT_{resultIndex}>")
                        results.AppendLine(content)
                        results.AppendLine($"</SEARCHRESULT_{resultIndex}>")
                        results.AppendLine()
                        context.Log($"  Result #{resultIndex}: {content.Length} chars from {url}")
                    Else
                        Dim charCount = If(content Is Nothing, 0, content.Length)
                        context.Log($"  Skipped (too short: {charCount} chars, min {ISearch_MinChars}): {url}")
                        ToolingFileLogger.LogStep($"Search result skipped (too short: {charCount} < {ISearch_MinChars}): {url}")
                    End If

                Catch ex As OperationCanceledException
                    Throw ' Re-throw cancellation
                Catch ex As Exception
                    context.Log($"  Error fetching {url}: {ex.Message}")
                    ToolingFileLogger.LogError("Internal search tool fetch error.", details:=$"url={url}", ex:=ex)
                End Try
            Next

            ' Footer: report all visited URLs for transparency
            results.AppendLine("<URLS_VISITED>")
            For Each vUrl In visitedUrls
                results.AppendLine($"  {vUrl}")
            Next
            results.AppendLine("</URLS_VISITED>")

            context.Log($"Search complete: {resultIndex} qualifying result(s) from {visitedUrls.Count} URL(s) visited")

            response.Response = results.ToString()
            response.Success = True

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn("Internal search tool cancelled.")

        Catch ex As System.Net.Http.HttpRequestException
            response.Success = False
            response.ErrorMessage = $"Search HTTP error: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool HTTP error.", ex:=ex)

        Catch ex As TaskCanceledException
            response.Success = False
            response.ErrorMessage = $"Search request timed out: {ex.Message}"
            ToolingFileLogger.LogError("Internal search tool timeout.", ex:=ex)

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Internal search tool error.", ex:=ex)
        End Try

        Return response
    End Function


    ''' <summary>
    ''' Executes the internal knowledge store search tool by querying the merged index
    ''' via KnowledgeQueryService and returning tagged document content blocks.
    ''' </summary>
    ''' <param name="toolCall">Tool call payload containing <c>query</c> and optional <c>max_results</c>.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <returns>Tool response containing relevant document content or an error.</returns>
    Private Async Function ExecuteInternalKnowledgeTool(toolCall As ToolCall, context As ToolExecutionContext) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            Dim boundStore = GetKnowledgeStoreForToolName(toolCall.ToolName)

            If boundStore Is Nothing Then
                response.Success = False
                response.ErrorMessage = "The selected Knowledge Store source could not be resolved."
                ToolingFileLogger.LogWarn("Internal knowledge tool: bound store could not be resolved.",
                    details:=$"ToolName='{toolCall.ToolName}'")
                Return response
            End If

            Dim storeLabel As String = KnowledgeStoreCatalog.GetDisplayLabel(boundStore)
            Dim query As String = GetToolArgumentString(toolCall.Arguments, "query")
            Dim tagName As String = GetToolArgumentString(toolCall.Arguments, "tag")

            Dim maxResults As Integer = 5
            If toolCall.Arguments.ContainsKey("max_results") Then
                Dim mr As Integer
                If Integer.TryParse(toolCall.Arguments("max_results")?.ToString(), mr) Then
                    maxResults = Math.Min(Math.Max(1, mr), 10)
                End If
            End If

            context.Log($"Knowledge store source: {storeLabel}")
            ToolingFileLogger.LogStep($"Knowledge store source: '{storeLabel}'; query='{query}'; tag='{tagName}'; max_results={maxResults}")

            ' Build the query for KnowledgeQueryService.
            ' IMPORTANT: Do NOT use "store:<name>" prefix — ResolveQueryAsync splits tokens
            ' by whitespace, so multi-word store names like "VISCHER Compliance" get truncated
            ' to just the first word, causing a name mismatch and zero results.
            ' Instead, pass only the tag filter (single-word) and the free-text query,
            ' then filter the returned matches to the bound store afterward.
            Dim resolveQuery As String = ""

            If Not String.IsNullOrWhiteSpace(tagName) Then
                resolveQuery &= $"tag:{tagName} "
            End If

            If Not String.IsNullOrWhiteSpace(query) Then
                resolveQuery &= query
            End If

            resolveQuery = resolveQuery.Trim()

            If String.IsNullOrWhiteSpace(resolveQuery) Then
                ' No query and no tag — pass just the store name as a broad keyword search
                Dim storeName As String = If(boundStore.Name, "").Trim()
                If Not String.IsNullOrWhiteSpace(storeName) Then
                    resolveQuery = storeName
                Else
                    response.Success = True
                    response.Response = $"No query provided for Knowledge Store '{storeLabel}'."
                    Return response
                End If
            End If

            context.Log($"Resolving knowledge query: '{resolveQuery}'")
            ToolingFileLogger.LogStep($"KnowledgeQueryService query: '{resolveQuery}'")

            ' Use the same semantic search path that Freestyle uses.
            ' Request extra results so we have enough after filtering to the bound store.
            Dim matches = Await KnowledgeQueryService.ResolveQueryAsync(resolveQuery, _context, maxResults * 4).ConfigureAwait(False)

            ' Filter to only the bound store (by Name match, case-insensitive)
            Dim storeName2 As String = If(boundStore.Name, "").Trim()
            If Not String.IsNullOrWhiteSpace(storeName2) AndAlso matches IsNot Nothing Then
                matches = matches.
                    Where(Function(m) Not String.IsNullOrWhiteSpace(m.StoreName) AndAlso
                                      m.StoreName.Equals(storeName2, StringComparison.OrdinalIgnoreCase)).
                    ToList()
            End If

            ' Apply the requested limit
            If matches IsNot Nothing AndAlso matches.Count > maxResults Then
                matches = matches.Take(maxResults).ToList()
            End If

            If matches Is Nothing OrElse matches.Count = 0 Then
                response.Success = True
                response.Response = $"No relevant documents found in Knowledge Store '{storeLabel}'."
                Return response
            End If

            ' ── Copy source files to temp dir so they are attached to the reply ──
            Dim copiedSourceNames As New List(Of String)()
            If Not String.IsNullOrWhiteSpace(_apCurrentTempDir) AndAlso Directory.Exists(_apCurrentTempDir) Then
                Dim copiedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each m In matches
                    Dim srcPath As String = If(m.SourcePath, "").Trim().Trim("'"c, """"c)
                    If String.IsNullOrWhiteSpace(srcPath) Then
                        context.Log($"  Skip copy (no SourcePath): title='{m.Title}', wiki='{If(m.WikiPagePath, "")}'")
                        Continue For
                    End If
                    If Not File.Exists(srcPath) Then
                        context.Log($"  Skip copy (file not found): '{srcPath}'")
                        Continue For
                    End If
                    If copiedPaths.Contains(srcPath) Then Continue For
                    copiedPaths.Add(srcPath)

                    Try
                        Dim destName = Path.GetFileName(srcPath)
                        Dim destPath = Path.Combine(_apCurrentTempDir, destName)

                        ' Avoid name collisions with existing files in temp dir
                        If File.Exists(destPath) Then
                            Dim baseName = Path.GetFileNameWithoutExtension(destName)
                            Dim ext = Path.GetExtension(destName)
                            Dim counter = 1
                            Do
                                destPath = Path.Combine(_apCurrentTempDir, $"{baseName}_{counter}{ext}")
                                counter += 1
                            Loop While File.Exists(destPath)
                            destName = Path.GetFileName(destPath)
                        End If

                        File.Copy(srcPath, destPath, overwrite:=False)
                        copiedSourceNames.Add(destName)
                        _apKnowledgeSourceCopies.Add(destPath)

                        ' Register as output so CollectResultAttachments picks it up.
                        ' When no user files exist (e.g. WebExtension Agent with no uploads),
                        ' create a placeholder entry so OutputFiles registration still works.
                        If _apCurrentAttachments IsNot Nothing Then
                            If _apCurrentAttachments.Count = 0 Then
                                _apCurrentAttachments.Add(New AutoPilotAttachmentInfo() With {
                                    .OriginalFileName = "(knowledge-source-placeholder)",
                                    .TempFilePath = Nothing,
                                    .Extension = "",
                                    .SizeBytes = 0,
                                    .IsOverSizeLimit = False,
                                    .StatusMessage = "placeholder",
                                    .OutputFiles = New List(Of String)(),
                                    .IsToolOutput = True
                                })
                            End If
                            Dim firstAtt = _apCurrentAttachments(0)
                            If firstAtt.OutputFiles Is Nothing Then firstAtt.OutputFiles = New List(Of String)()
                            firstAtt.OutputFiles.Add(destPath)
                        End If

                        ' Update the match's SourcePath to point at the copy so
                        ' BuildKnowledgeContext emits the temp-dir filename the LLM can
                        ' reference as an attached file name.
                        m.SourcePath = destPath

                        context.Log($"Attached source: {destName}")
                    Catch ex As Exception
                        ToolingFileLogger.LogWarn($"Could not copy knowledge source '{srcPath}': {ex.Message}")
                    End Try
                Next
            Else
                context.Log($"File copy skipped: _apCurrentTempDir='{If(_apCurrentTempDir, "(Nothing)")}', exists={If(Not String.IsNullOrWhiteSpace(_apCurrentTempDir), Directory.Exists(_apCurrentTempDir).ToString(), "N/A")}")
            End If

            Dim knowledgeContext As String = KnowledgeQueryService.BuildKnowledgeContext(matches, 200000)

            If String.IsNullOrWhiteSpace(knowledgeContext) Then
                response.Success = True
                response.Response = $"No readable content could be built from Knowledge Store '{storeLabel}'."
                Return response
            End If

            ' Append a summary of attached source files so the LLM knows they are available
            If copiedSourceNames.Count > 0 Then
                knowledgeContext &= vbCrLf &
                    "<KNOWLEDGE_ATTACHMENTS>" & vbCrLf &
                    "The following source documents have been attached to the reply for the recipient:" & vbCrLf &
                    String.Join(vbCrLf, copiedSourceNames.Select(Function(n) $"  - {n}")) & vbCrLf &
                    "When referencing these documents, mention them by filename so the recipient can locate the attachment." & vbCrLf &
                    "</KNOWLEDGE_ATTACHMENTS>"
            End If

            response.Success = True
            response.Response = knowledgeContext

            context.Log($"Knowledge search returned content ({knowledgeContext.Length:N0} chars) from '{storeLabel}'" &
                        If(copiedSourceNames.Count > 0, $", {copiedSourceNames.Count} source file(s) attached.", "."), "success")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = $"Knowledge store search failed: {ex.Message}"
            ToolingFileLogger.LogError("Internal knowledge tool error.", ex:=ex)
        End Try

        Return response
    End Function


    ''' <summary>
    ''' Removes knowledge-store source files from the temp directory that were
    ''' not cited by the LLM in its final response. Files produced by other tools
    ''' (process_word_document, merge_pdfs, etc.) are never affected.
    ''' Must be called BEFORE CollectResultAttachments so the directory scan
    ''' does not re-pick up uncited knowledge files.
    ''' </summary>
    Private Sub RemoveUncitedKnowledgeSourceCopies(llmResponseText As String)
        If _apKnowledgeSourceCopies.Count = 0 Then Return
        If String.IsNullOrWhiteSpace(llmResponseText) Then Return

        Dim toRemove As New List(Of String)()

        For Each filePath In _apKnowledgeSourceCopies
            Dim fileName = Path.GetFileName(filePath)
            Dim baseName = Path.GetFileNameWithoutExtension(filePath)

            Dim cited = llmResponseText.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    (baseName.Length >= 4 AndAlso llmResponseText.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) >= 0)

            If Not cited Then toRemove.Add(filePath)
        Next

        ' Safety: if ALL knowledge files are uncited, keep them (LLM may have paraphrased)
        If toRemove.Count = _apKnowledgeSourceCopies.Count Then Return

        For Each uncitedPath In toRemove
            Try
                If File.Exists(uncitedPath) Then File.Delete(uncitedPath)
            Catch
            End Try
            _apKnowledgeSourceCopies.Remove(uncitedPath)

            ' Also remove from OutputFiles so CollectResultAttachments doesn't find it
            If _apCurrentAttachments IsNot Nothing Then
                For Each att In _apCurrentAttachments
                    If att.OutputFiles IsNot Nothing Then att.OutputFiles.Remove(uncitedPath)
                Next
            End If
        Next
    End Sub


    ''' <summary>
    ''' Filters result file paths to only those whose filename or base name appears in
    ''' the LLM's final response text. Prevents uncited knowledge-store source files
    ''' from being delivered. Deletes uncited files from disk so they are not picked up
    ''' by the fallback scan in CollectResultAttachments. Returns all files as a safety
    ''' fallback if none matched (the LLM may have paraphrased).
    ''' </summary>
    Friend Shared Function FilterAttachmentsByCitation(resultFiles As List(Of String), llmResponseText As String) As List(Of String)
        If resultFiles Is Nothing OrElse resultFiles.Count = 0 Then Return If(resultFiles, New List(Of String)())
        If String.IsNullOrWhiteSpace(llmResponseText) Then Return resultFiles

        Dim cited As New List(Of String)()
        Dim uncited As New List(Of String)()

        For Each filePath In resultFiles
            Dim fileName = Path.GetFileName(filePath)
            Dim baseName = Path.GetFileNameWithoutExtension(filePath)

            ' Match full filename or base name (≥4 chars to avoid false positives)
            If llmResponseText.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
           (baseName.Length >= 4 AndAlso llmResponseText.IndexOf(baseName, StringComparison.OrdinalIgnoreCase) >= 0) Then
                cited.Add(filePath)
            Else
                uncited.Add(filePath)
            End If
        Next

        ' Safety fallback: if nothing matched, LLM may have paraphrased — keep all
        If cited.Count = 0 Then Return resultFiles

        ' Remove uncited files from temp dir so CollectResultAttachments's
        ' fallback directory scan doesn't re-pick them up
        For Each uncitedPath In uncited
            Try
                If File.Exists(uncitedPath) Then File.Delete(uncitedPath)
            Catch
            End Try
        Next

        Return cited
    End Function


    Private Function GetToolParameterSchemas(toolConfig As ModelConfig) As Dictionary(Of String, JToken)
        Dim result As New Dictionary(Of String, JToken)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim propertiesObject As JObject = TryCast(toolDefinition.SelectToken("parameters.properties"), JObject)

            If propertiesObject Is Nothing Then
                Return result
            End If

            For Each prop As JProperty In propertiesObject.Properties()
                result(prop.Name) = prop.Value
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse tool parameter schemas.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolRequiredParameters(toolConfig As ModelConfig) As HashSet(Of String)
        Dim result As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If toolConfig Is Nothing OrElse String.IsNullOrWhiteSpace(toolConfig.ToolDefinition) Then
            Return result
        End If

        Try
            Dim toolDefinition As JObject = JObject.Parse(toolConfig.ToolDefinition)
            Dim requiredArray As JArray = TryCast(toolDefinition.SelectToken("parameters.required"), JArray)

            If requiredArray Is Nothing Then
                Return result
            End If

            For Each item As JToken In requiredArray
                Dim name As String = If(item, "").ToString().Trim()
                If name <> "" Then
                    result.Add(name)
                End If
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn(
                "Failed to parse required tool parameters.",
                details:=$"ToolName='{If(toolConfig.ToolName, "")}'",
                ex:=ex)
        End Try

        Return result
    End Function

    Private Function GetToolParameterType(schemaToken As JToken) As String
        If schemaToken Is Nothing Then
            Return "string"
        End If

        Dim typeToken As JToken = schemaToken("type")
        Dim typeName As String = If(typeToken, "").ToString().Trim().ToLowerInvariant()

        If typeName <> "" Then
            Return typeName
        End If

        If schemaToken("enum") IsNot Nothing Then
            Return "string"
        End If

        Return "string"
    End Function

    Private Function GetToolParameterEnumValues(schemaToken As JToken) As List(Of String)
        Dim values As New List(Of String)()

        If schemaToken Is Nothing Then
            Return values
        End If

        Dim enumArray As JArray = TryCast(schemaToken("enum"), JArray)
        If enumArray Is Nothing Then
            Return values
        End If

        For Each item As JToken In enumArray
            Dim value As String = If(item, "").ToString()
            If value <> "" Then
                values.Add(value)
            End If
        Next

        Return values
    End Function

    Private Function TryParseBooleanLiteral(value As String, ByRef result As Boolean) As Boolean
        Dim normalized As String = If(value, "").Trim()

        If Boolean.TryParse(normalized, result) Then
            Return True
        End If

        Select Case normalized.ToLowerInvariant()
            Case "1", "yes", "y"
                result = True
                Return True
            Case "0", "no", "n"
                result = False
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function FormatToolValueForPlaceholder(rawValue As String, schemaToken As JToken) As String
        Dim parameterType As String = GetToolParameterType(schemaToken)
        Dim safeValue As String = If(rawValue, "").Trim()

        Select Case parameterType
            Case "boolean"
                Dim boolValue As Boolean = False
                If TryParseBooleanLiteral(safeValue, boolValue) Then
                    Return If(boolValue, "true", "false")
                End If
                Return "false"

            Case "integer"
                Dim longValue As Long
                If Long.TryParse(safeValue, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, longValue) OrElse
                   Long.TryParse(safeValue, longValue) Then
                    Return longValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If
                Return "0"

            Case "number"
                Dim doubleValue As Double
                Dim normalized As String = safeValue.Replace(","c, "."c)

                If Double.TryParse(normalized, Globalization.NumberStyles.Float Or Globalization.NumberStyles.AllowThousands,
                                   Globalization.CultureInfo.InvariantCulture, doubleValue) OrElse
                   Double.TryParse(safeValue, doubleValue) Then
                    Return doubleValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If

                Return "0"

            Case "array"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Array Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "[]"

            Case "object"
                If safeValue <> "" Then
                    Try
                        Dim parsed As JToken = JToken.Parse(safeValue)
                        If parsed.Type = JTokenType.Object Then
                            Return parsed.ToString(Formatting.None)
                        End If
                    Catch
                    End Try
                End If
                Return "{}"

            Case Else
                Return EscapeJsonString(safeValue)
        End Select
    End Function

    Private Function ResolveToolDefaultValue(placeholderName As String,
                                             toolDefaults As IDictionary(Of String, String),
                                             schemaToken As JToken,
                                             isRequired As Boolean,
                                             ByRef shouldRemoveProperty As Boolean) As String
        shouldRemoveProperty = False

        Dim rawDefault As String = ""

        If toolDefaults IsNot Nothing AndAlso toolDefaults.ContainsKey(placeholderName) Then
            rawDefault = If(toolDefaults(placeholderName), "")
        End If

        If String.IsNullOrWhiteSpace(rawDefault) Then
            If isRequired Then
                Return "{" & placeholderName & "}"
            End If

            shouldRemoveProperty = True
            Return ""
        End If

        Return FormatToolValueForPlaceholder(rawDefault, schemaToken)
    End Function

    Private Function RemoveToolArgumentPlaceholderProperty(apiCall As String, placeholderName As String) As String
        If String.IsNullOrWhiteSpace(apiCall) OrElse String.IsNullOrWhiteSpace(placeholderName) Then
            Return apiCall
        End If

        Dim quotedPropertyName As String = Regex.Escape("""" & placeholderName & """")
        Dim rawPlaceholder As String = Regex.Escape("{" & placeholderName & "}")
        Dim quotedPlaceholder As String = Regex.Escape("""{" & placeholderName & "}""")

        Dim patterns As String() = {
            ",\s*" & quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            ",\s*" & quotedPropertyName & "\s*:\s*" & rawPlaceholder,
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & rawPlaceholder & "\s*,",
            quotedPropertyName & "\s*:\s*" & quotedPlaceholder,
            quotedPropertyName & "\s*:\s*" & rawPlaceholder
        }

        Dim result As String = apiCall

        For Each pattern As String In patterns
            result = Regex.Replace(
                result,
                pattern,
                "",
                RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        Next

        result = Regex.Replace(result, ",\s*(\}|\])", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)
        result = Regex.Replace(result, "(\{|\[)\s*,", "$1", RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        Return result
    End Function

    Private Function TryExtractToolServiceErrorMessage(rawResponse As String, ByRef errorMessage As String) As Boolean
        errorMessage = ""

        If String.IsNullOrWhiteSpace(rawResponse) Then
            Return False
        End If

        Try
            Dim root As JObject = JObject.Parse(rawResponse)

            Dim errorToken As JToken = root("error")
            If errorToken IsNot Nothing Then
                Dim message As String = If(errorToken("message"), "").ToString().Trim()
                Dim code As String = If(errorToken("code"), "").ToString().Trim()

                If message = "" Then
                    message = errorToken.ToString(Formatting.None)
                End If

                errorMessage = If(code <> "", $"{code}: {message}", message)
                Return True
            End If

            Dim isErrorToken As JToken = root.SelectToken("result.isError")
            Dim isError As Boolean = False

            If isErrorToken IsNot Nothing Then
                If isErrorToken.Type = JTokenType.Boolean Then
                    isError = isErrorToken.Value(Of Boolean)()
                Else
                    Boolean.TryParse(isErrorToken.ToString(), isError)
                End If
            End If

            If Not isError Then
                Return False
            End If

            Dim messages As New List(Of String)()
            Dim contentArray As JArray = TryCast(root.SelectToken("result.content"), JArray)

            If contentArray IsNot Nothing Then
                For Each item As JToken In contentArray
                    Dim text As String = If(item("text"), "").ToString().Trim()
                    If text <> "" Then
                        messages.Add(text)
                    End If
                Next
            End If

            If messages.Count > 0 Then
                errorMessage = String.Join(" ", messages)
            Else
                Dim resultToken As JToken = root("result")
                errorMessage = If(resultToken Is Nothing, "Tool service returned an error.", resultToken.ToString(Formatting.None))
            End If

            Return True
        Catch
            Return False
        End Try
    End Function


    ''' <summary>
    ''' Executes an external tool by applying its <see cref="ModelConfig"/> to <c>_context</c>, preparing
    ''' the tool API call payload, and invoking <c>LLM</c> in JSON response mode.
    ''' For SSE/MCP endpoints (prefixed with "sse:"), bypasses LLM() and calls the MCP server directly.
    ''' </summary>
    ''' <param name="toolCall">Tool call extracted from the LLM response.</param>
    ''' <param name="toolConfig">Tool configuration to apply for this call.</param>
    ''' <param name="context">Tool execution context for logging and diagnostics.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool response containing the tool service result or an error.</returns>
    Private Async Function ExecuteExternalTool(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)
        Dim response As New ToolResponse() With {
            .CallId = toolCall.CallId,
            .ToolName = toolCall.ToolName
        }

        Try
            cancellationToken.ThrowIfCancellationRequested()

            Dim apiCallTemplate = toolConfig.ToolAPICall
            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                apiCallTemplate = toolConfig.APICall
            End If

            If String.IsNullOrWhiteSpace(apiCallTemplate) Then
                response.Success = False
                response.ErrorMessage = "Tool has no APICall template defined"
                ToolingFileLogger.LogError("Tool has no APICall template defined.", details:=$"ToolName='{toolCall.ToolName}'")
                Return response
            End If

            Dim apiCall = apiCallTemplate
            Dim parameterSchemas = GetToolParameterSchemas(toolConfig)
            Dim requiredParameters = GetToolRequiredParameters(toolConfig)

            For Each kvp In toolCall.Arguments
                Dim placeholder = "{" & kvp.Key & "}"
                Dim schemaToken As JToken = Nothing
                parameterSchemas.TryGetValue(kvp.Key, schemaToken)

                Dim value As String
                If kvp.Value Is Nothing Then
                    value = ""
                ElseIf TypeOf kvp.Value Is JToken Then
                    Dim jt = DirectCast(kvp.Value, JToken)
                    If jt.Type = JTokenType.String Then
                        value = FormatToolValueForPlaceholder(jt.Value(Of String)(), schemaToken)
                    Else
                        value = jt.ToString(Formatting.None)
                    End If
                Else
                    value = FormatToolValueForPlaceholder(kvp.Value.ToString(), schemaToken)
                End If

                apiCall = apiCall.Replace(placeholder, value)
            Next

            Dim unreplacedPattern As New Regex("\{([a-zA-Z_][a-zA-Z0-9_]*)\}")
            Dim unreplacedMatches = unreplacedPattern.Matches(apiCall)

            If unreplacedMatches.Count > 0 Then
                ToolingFileLogger.LogWarn(
                    "Unreplaced placeholders found in tool APICall (defaults will be applied if available).",
                    details:=$"ToolName='{toolCall.ToolName}'; Count={unreplacedMatches.Count}; APICall='{apiCall}'")

                Dim toolDefaults As Dictionary(Of String, String) = Nothing
                If Not String.IsNullOrWhiteSpace(toolConfig.ToolParameterDefaults) Then
                    Try
                        toolDefaults = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(toolConfig.ToolParameterDefaults)
                    Catch ex As Exception
                        ToolingFileLogger.LogWarn(
                            "ToolParameterDefaults parse failed.",
                            details:=$"ToolName='{toolCall.ToolName}'; ToolParameterDefaults='{toolConfig.ToolParameterDefaults}'",
                            ex:=ex)
                    End Try
                End If

                For Each m As Match In unreplacedMatches
                    Dim placeholderName = m.Groups(1).Value
                    Dim schemaToken As JToken = Nothing
                    parameterSchemas.TryGetValue(placeholderName, schemaToken)

                    Dim shouldRemoveProperty As Boolean = False
                    Dim replacement As String = ResolveToolDefaultValue(
                        placeholderName,
                        toolDefaults,
                        schemaToken,
                        requiredParameters.Contains(placeholderName),
                        shouldRemoveProperty)

                    If shouldRemoveProperty Then
                        apiCall = RemoveToolArgumentPlaceholderProperty(apiCall, placeholderName)
                    Else
                        apiCall = apiCall.Replace(m.Value, replacement)
                    End If
                Next

                Dim remainingMatches = unreplacedPattern.Matches(apiCall)
                If remainingMatches.Count > 0 Then
                    Dim remainingNames = remainingMatches.
                        Cast(Of Match)().
                        Select(Function(m) m.Groups(1).Value).
                        Distinct(StringComparer.OrdinalIgnoreCase).
                        ToList()

                    response.Success = False
                    response.ErrorMessage = $"Unreplaced placeholders remain after applying defaults: {String.Join(", ", remainingNames)}"

                    ToolingFileLogger.LogError(
                        "Unreplaced placeholders remain in tool APICall after applying defaults.",
                        details:=$"ToolName='{toolCall.ToolName}'; RemainingCount={remainingMatches.Count}; Remaining='{String.Join(", ", remainingNames)}'; APICall='{apiCall}'")

                    Return response
                End If
            End If

            If toolConfig.OAuth2 Then
                toolConfig.DecodedAPI = Await SharedMethods.GetFreshAccessToken(
                    _context,
                    toolConfig.OAuth2ClientMail,
                    toolConfig.OAuth2Scopes,
                    toolConfig.APIKey,
                    toolConfig.OAuth2Endpoint,
                    toolConfig.OAuth2ATExpiry,
                    True,
                    False).ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(toolConfig.DecodedAPI) Then
                    response.Success = False
                    response.ErrorMessage = "OAuth2 authentication failed."
                    ToolingFileLogger.LogError(
                        "OAuth2 authentication failed before MCP tool execution.",
                        details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If
            End If

            ' ── SSE transport: full round-trip bypassing LLM() ───────────
            If Not String.IsNullOrWhiteSpace(toolConfig.Endpoint) AndAlso
               toolConfig.Endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then

                Dim sseBase = toolConfig.Endpoint.Substring(SharedMethods.MCP_SSE_PREFIX.Length)
                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"SSE transport: executing tool {toolCall.ToolName} via {sseBase}")
                ToolingFileLogger.LogStep($"SSE round-trip for {toolCall.ToolName} at {sseBase}")
                ToolingFileLogger.LogStep($"SSE request body: {apiCall}")

                Dim sseAttemptedRefresh As Boolean = False
                Dim sseEx As Exception = Nothing
                Dim sseDone As Boolean = False
                Dim sseCancelled As Boolean = False

                Do
                    sseEx = Nothing

                    Try
                        cancellationToken.ThrowIfCancellationRequested()

                        resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                        Dim rawResult = Await SharedMethods.ExecuteMCPSSEToolCall(
                            _context,
                            sseBase, apiCall,
                            If(toolConfig.HeaderA, ""), resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"SSE tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "SSE tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from SSE tool service"
                            ToolingFileLogger.LogError("Empty SSE response.", details:=$"ToolName='{toolCall.ToolName}'")
                        End If

                        sseDone = True

                    Catch ex As OperationCanceledException
                        sseCancelled = True
                        response.Success = False
                        response.ErrorMessage = "Operation was cancelled"
                        ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled during SSE execution.")
                    Catch ex As Exception
                        sseEx = ex
                    End Try

                    If sseDone OrElse sseCancelled Then Exit Do
                    If sseAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, sseEx) Then Exit Do

                    sseAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "SSE tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'")

                    Dim sseRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not sseRefreshOk Then Exit Do
                Loop

                If Not sseDone AndAlso Not sseCancelled AndAlso sseEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"SSE tool call failed: {sseEx.Message}"
                    ToolingFileLogger.LogError("SSE tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'", ex:=sseEx)
                End If

                Return response
            End If

            ' ── MCP Streamable HTTP transport: full round-trip bypassing LLM() ─
            If IsMCPStreamableToolCall(toolConfig.Endpoint, apiCall) Then
                Dim mcpUrl As String = If(toolConfig.Endpoint, "")
                If mcpUrl.StartsWith(SharedMethods.MCP_STREAMABLE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                    mcpUrl = mcpUrl.Substring(SharedMethods.MCP_STREAMABLE_PREFIX.Length)
                End If

                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"MCP Streamable HTTP: executing tool {toolCall.ToolName} via {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP round-trip for {toolCall.ToolName} at {mcpUrl}")
                ToolingFileLogger.LogStep($"MCP Streamable HTTP request body: {apiCall}")

                Dim streamAttemptedRefresh As Boolean = False
                Dim streamEx As Exception = Nothing
                Dim streamDone As Boolean = False
                Dim streamCancelled As Boolean = False

                Do
                    streamEx = Nothing

                    Try
                        cancellationToken.ThrowIfCancellationRequested()

                        resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                        Dim rawResult = Await SharedMethods.ExecuteMCPStreamableToolCall(
                            mcpUrl,
                            apiCall,
                            If(toolConfig.HeaderA, ""),
                            resolvedHeaderB,
                            CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                        ToolingFileLogger.LogRawResponseStub($"MCP Streamable HTTP tool result ({toolCall.ToolName})", rawResult)

                        If Not String.IsNullOrWhiteSpace(rawResult) Then
                            Dim toolErrorMessage As String = ""
                            response.Response = rawResult

                            If TryExtractToolServiceErrorMessage(rawResult, toolErrorMessage) Then
                                response.Success = False
                                response.ErrorMessage = toolErrorMessage
                                ToolingFileLogger.LogWarn(
                                    "MCP Streamable HTTP tool service returned a logical error.",
                                    details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                            Else
                                response.Success = True
                            End If
                        Else
                            response.Success = False
                            response.ErrorMessage = "Empty response from MCP Streamable HTTP tool service"
                            ToolingFileLogger.LogError(
                                "Empty MCP Streamable HTTP response.",
                                details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")
                        End If

                        streamDone = True

                    Catch ex As OperationCanceledException
                        streamCancelled = True
                        response.Success = False
                        response.ErrorMessage = "Operation was cancelled"
                        ToolingFileLogger.LogWarn(
                            $"Tool {toolCall.ToolName} cancelled during MCP Streamable HTTP execution.")
                    Catch ex As Exception
                        streamEx = ex
                    End Try

                    If streamDone OrElse streamCancelled Then Exit Do
                    If streamAttemptedRefresh Then Exit Do
                    If Not ShouldRetryMCPAfterUnauthorized(toolConfig, streamEx) Then Exit Do

                    streamAttemptedRefresh = True
                    ToolingFileLogger.LogWarn(
                        "MCP Streamable HTTP tool call returned Unauthorized. Forcing MCP OAuth refresh and retrying once.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'")

                    Dim streamRefreshOk As Boolean = Await ForceRefreshToolOAuthToken(toolConfig, toolCall.ToolName).ConfigureAwait(False)
                    If Not streamRefreshOk Then Exit Do
                Loop

                If Not streamDone AndAlso Not streamCancelled AndAlso streamEx IsNot Nothing Then
                    response.Success = False
                    response.ErrorMessage = $"MCP Streamable HTTP tool call failed: {streamEx.Message}"
                    ToolingFileLogger.LogError(
                        "MCP Streamable HTTP tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; Endpoint='{mcpUrl}'",
                        ex:=streamEx)
                End If

                Return response
            End If

            ' ── Standard transport: route through LLM() ──────────────────
            Dim backupConfig = GetCurrentConfig(_context)

            Try
                cancellationToken.ThrowIfCancellationRequested()

                Dim errorFlag As Boolean = False
                ApplyModelConfig(_context, toolConfig, errorFlag)
                If errorFlag Then
                    response.Success = False
                    response.ErrorMessage = "Failed to apply tool configuration"
                    ToolingFileLogger.LogError("Failed to apply tool configuration.", details:=$"ToolName='{toolCall.ToolName}'")
                    Return response
                End If

                _context.INI_APICall_2 = apiCall

                Dim originalResponse = _context.INI_Response_2
                _context.INI_Response_2 = "JSON"

                context.Log($"Calling external service for tool: {toolCall.ToolName}")
                ToolingFileLogger.LogPreToolLlmCallSnapshot(_context)

                Dim result = Await LLM("", "", "", "", 0, True, True, "", "", cancellationToken, EnsureUI:=False)

                ToolingFileLogger.LogRawResponseStub($"Tool LLM() result ({toolCall.ToolName})", result)

                _context.INI_Response_2 = originalResponse

                If Not String.IsNullOrWhiteSpace(result) Then
                    Dim toolErrorMessage As String = ""

                    response.Response = result

                    If TryExtractToolServiceErrorMessage(result, toolErrorMessage) Then
                        response.Success = False
                        response.ErrorMessage = toolErrorMessage
                        ToolingFileLogger.LogWarn(
                            "Tool service returned a logical error.",
                            details:=$"ToolName='{toolCall.ToolName}'; Error='{toolErrorMessage}'")
                    Else
                        response.Success = True
                    End If
                Else
                    response.Success = False
                    response.ErrorMessage = "Empty response from tool service"
                    ToolingFileLogger.LogError("Empty response from tool service.", details:=$"ToolName='{toolCall.ToolName}'; APICall='{apiCall}'")
                End If

            Finally
                RestoreDefaults(_context, backupConfig)
            End Try

        Catch ex As OperationCanceledException
            response.Success = False
            response.ErrorMessage = "Operation was cancelled"
            ToolingFileLogger.LogWarn($"Tool {toolCall.ToolName} cancelled during external execution.")

        Catch ex As Exception
            response.Success = False
            response.ErrorMessage = ex.Message
            ToolingFileLogger.LogError("Tool execution error.", details:=$"ToolName='{toolCall.ToolName}'", ex:=ex)
        End Try

        Return response
    End Function

    ''' <summary>
    ''' Loads tooling service configurations from an INI file and returns tool-capable <see cref="ModelConfig"/> entries.
    ''' </summary>
    ''' <param name="iniPath">INI path containing tool model sections.</param>
    ''' <param name="toolsOnly">When True, filters to entries that have tool-specific prompt/definition fields.</param>
    ''' <returns>List of available tool configurations.</returns>
    Public Function LoadToolingServices(iniPath As String, Optional toolsOnly As Boolean = True) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then
            Return tools
        End If

        Try
            Dim allModels = LoadAlternativeModels(iniPath, _context, StartWithUpcase(ToolFriendlyName), includeToolOnly:=True, toolsOnly:=toolsOnly)

            For Each mc In allModels
                If mc.Deprecated Then Continue For

                If toolsOnly Then
                    If String.IsNullOrWhiteSpace(mc.ToolInstructionsPrompt) AndAlso
                       String.IsNullOrWhiteSpace(mc.ToolDefinition) Then
                        Continue For
                    End If
                End If

                mc.Tool = True
                tools.Add(mc)
            Next

        Catch ex As Exception
            Debug.WriteLine($"LoadToolingServices error: {ex.Message}")
            ToolingFileLogger.LogError("LoadToolingServices error.", ex:=ex)
        End Try

        Return tools
    End Function

    ''' <summary>
    ''' Shows the tool selection dialog and persists the selected tool names into <c>My.Settings.SelectedToolNames</c>.
    ''' </summary>
    ''' <param name="availableTools">List of available tool configurations.</param>
    ''' <param name="preselectAll">Unused parameter in this method body (caller passes a value).</param>
    ''' <returns>Selected tools when the dialog result is OK; otherwise Nothing.</returns>
    Public Function ShowToolSelectionDialog(availableTools As List(Of ModelConfig), Optional preselectAll As Boolean = True, Optional FriendlyName As String = "Tools") As List(Of ModelConfig)
        If availableTools Is Nothing OrElse availableTools.Count = 0 Then
            Return New List(Of ModelConfig)()
        End If

        ' Note: Do NOT add ToolingSuffix here. ToolingSuffix is for models that CAN USE tools/sources,
        ' not for the tools/sources themselves. The tools in this list ARE the sources.
        ' InternalToolSuffix is already applied to the internal web tool in GetInternalWebTool().

        Dim selector As New MultiModelSelectorForm(
        availableTools,
        "",
        $"{AN} - Select {FriendlyName}",
        resetChecked:=False,
        preselectMany:=If(SelectedToolNames, New List(Of String)()),
        $"Select the {Globals.ThisAddIn.ToolFriendlyName.ToLower} you want to make available to the model:")

        If selector.ShowDialog() = DialogResult.OK Then
            Dim selected = selector.SelectedModels
            SelectedToolNames = selected.Select(Function(t) t.ToolName).ToList()
            Try
                My.Settings.SelectedToolNames = String.Join("|", SelectedToolNames)
                My.Settings.Save()
            Catch ex As Exception
                ToolingFileLogger.LogWarn("Failed to persist SelectedToolNames.", ex:=ex)
            End Try
            Return selected
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Returns all available tools by loading external tools from <c>INI_SpecialServicePath</c>,
    ''' adding the internal web tool, conditionally adding the internal search tool
    ''' (only when <c>INI_ISearch</c> is enabled and <c>INI_ISearch_URL</c> is configured),
    ''' and conditionally adding the internal knowledge store search tool
    ''' (only when a knowledge store path is configured and at least one store is indexed).
    ''' </summary>
    ''' <returns>List of available tools.</returns>
    Public Function GetAvailableTools(Optional includeInteractiveM365Tools As Boolean = False) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        If Not String.IsNullOrWhiteSpace(INI_SpecialServicePath) Then
            Dim externalTools = LoadToolingServices(INI_SpecialServicePath, True)
            tools.AddRange(externalTools)
        End If

        tools.Add(GetInternalWebTool())

        If INI_ISearch AndAlso Not String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
            tools.Add(GetInternalSearchTool(enforcePrivacy:=INI_EnablePrivacyForSearch))
        End If

        tools.AddRange(GetInternalKnowledgeTools())

        ' M365 tools are interactive-only. Show them for Local Chat Agent source
        ' selection and execution, but never for AutoPilot.
        If (includeInteractiveM365Tools OrElse _chatAgentActive) AndAlso Not _apActive Then
            tools.AddRange(SharedLibrary.SharedLibrary.M365ToolService.GetTools(_context, InternalToolSuffix))
        End If

        Return tools
    End Function

    ''' <summary>
    ''' Loads persisted tool selection from <c>My.Settings.SelectedToolNames</c> into <c>SelectedToolNames</c>.
    ''' </summary>
    Public Sub LoadPersistedToolSelection()
        Try
            Dim saved = My.Settings.SelectedToolNames
            If Not String.IsNullOrWhiteSpace(saved) Then
                SelectedToolNames = saved.Split("|"c).ToList()
            End If
        Catch ex As Exception
            SelectedToolNames = New List(Of String)()
            ToolingFileLogger.LogWarn("Failed to load persisted tool selection.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Selects tools for the current session either by reusing persisted selections or by showing the tool selection dialog.
    ''' </summary>
    ''' <param name="forceDialog">If True, always shows the selection dialog.</param>
    ''' <returns>Selected tool configurations, or Nothing when the dialog is canceled or no tools are available.</returns>
    Public Function SelectToolsForSession(Optional forceDialog As Boolean = False,
                                      Optional FriendlyName As String = ToolFriendlyName,
                                      Optional includeInteractiveM365Tools As Boolean = False) As List(Of ModelConfig)
        Dim availableTools = GetAvailableTools(includeInteractiveM365Tools)

        If availableTools.Count = 0 Then
            ShowCustomMessageBox($"No {FriendlyName.ToLower} are available. Configure 'Tooling' in the Special Services configuration file.")
            Return Nothing
        End If

        LoadPersistedToolSelection()

        If Not forceDialog AndAlso SelectedToolNames IsNot Nothing AndAlso SelectedToolNames.Count > 0 Then
            Dim selectedNameSet As New HashSet(Of String)(SelectedToolNames, StringComparer.OrdinalIgnoreCase)

            Dim selected = availableTools.
        Where(Function(t) Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso selectedNameSet.Contains(t.ToolName)).
        ToList()

            If selected.Count > 0 Then
                Return selected
            End If
        End If

        Dim hasPersistedSelection As Boolean = (SelectedToolNames IsNot Nothing AndAlso SelectedToolNames.Count > 0)

        Return ShowToolSelectionDialog(availableTools, preselectAll:=Not hasPersistedSelection, FriendlyName)
    End Function

    Private Class ToolSourceLink
        Public Property Url As String
        Public Property Title As String
        Public Property Source As String
    End Class

    Private Function AppendM365SourcesFooter(finalAnswer As String,
                                             toolResponses As List(Of ToolResponse)) As String
        Dim answer As String = If(finalAnswer, "").Trim()
        Dim links As List(Of ToolSourceLink) = ExtractM365SourceLinks(toolResponses, answer)

        If links.Count = 0 Then
            Return answer
        End If

        Dim sb As New StringBuilder()

        If answer.Length > 0 Then
            sb.AppendLine(answer)
            sb.AppendLine()
        End If

        sb.AppendLine("### Sources")

        For Each link In links
            Dim label As String = BuildSourceLinkLabel(link)
            sb.AppendLine($"- [{EscapeMarkdownLinkText(label)}]({link.Url})")
        Next

        Return sb.ToString().Trim()
    End Function

    Private Function ExtractM365SourceLinks(toolResponses As List(Of ToolResponse),
                                            existingAnswer As String) As List(Of ToolSourceLink)
        Dim results As New List(Of ToolSourceLink)()
        Dim seenUrls As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim answerText As String = If(existingAnswer, "")

        If toolResponses Is Nothing OrElse toolResponses.Count = 0 Then
            Return results
        End If

        For Each response As ToolResponse In toolResponses
            If response Is Nothing OrElse Not response.Success OrElse String.IsNullOrWhiteSpace(response.Response) Then
                Continue For
            End If

            If String.Equals(response.ToolName, "m365_search", StringComparison.OrdinalIgnoreCase) Then
                ExtractM365SearchLinks(response.Response, answerText, seenUrls, results)
            ElseIf IsM365RetrievalToolName(response.ToolName) Then
                ExtractM365WrappedContentLink(response.ToolName, response.Response, answerText, seenUrls, results)
            End If

            If results.Count >= 12 Then
                Exit For
            End If
        Next

        Return results
    End Function

    Private Sub ExtractM365SearchLinks(responseText As String,
                                       existingAnswer As String,
                                       seenUrls As HashSet(Of String),
                                       results As List(Of ToolSourceLink))
        Dim root As JObject = Nothing

        Try
            root = JObject.Parse(responseText)
        Catch
            Exit Sub
        End Try

        Dim hits As JArray = TryCast(root("hits"), JArray)
        If hits Is Nothing OrElse hits.Count = 0 Then
            Exit Sub
        End If

        For Each hitToken As JToken In hits
            Dim hit As JObject = TryCast(hitToken, JObject)
            If hit Is Nothing Then Continue For

            Dim url As String = If(hit("web_url")?.ToString(), "").Trim()
            Dim title As String = If(hit("title")?.ToString(), "").Trim()
            Dim source As String = If(hit("source")?.ToString(), "").Trim()

            TryAddSourceLink(url, title, source, existingAnswer, seenUrls, results)

            If results.Count >= 12 Then
                Exit For
            End If
        Next
    End Sub

    Private Sub ExtractM365WrappedContentLink(toolName As String,
                                              responseText As String,
                                              existingAnswer As String,
                                              seenUrls As HashSet(Of String),
                                              results As List(Of ToolSourceLink))
        Dim urlMatch As Match = Regex.Match(
            responseText,
            "<WEB_URL>\s*(.*?)\s*</WEB_URL>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        If Not urlMatch.Success Then
            Exit Sub
        End If

        Dim titleMatch As Match = Regex.Match(
            responseText,
            "^<(?<kind>[A-Z_]+)\s+id=""[^""]*""\s+title=""(?<title>[^""]*)""[^>]*>",
            RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        Dim title As String = ""
        If titleMatch.Success Then
            title = titleMatch.Groups("title").Value.Trim()
        End If

        Dim url As String = urlMatch.Groups(1).Value.Trim()
        Dim source As String = GetM365SourceFromToolName(toolName)

        TryAddSourceLink(url, title, source, existingAnswer, seenUrls, results)
    End Sub

    Private Sub TryAddSourceLink(url As String,
                                 title As String,
                                 source As String,
                                 existingAnswer As String,
                                 seenUrls As HashSet(Of String),
                                 results As List(Of ToolSourceLink))
        Dim cleanUrl As String = If(url, "").Trim()
        If String.IsNullOrWhiteSpace(cleanUrl) Then
            Exit Sub
        End If

        If Not String.IsNullOrWhiteSpace(existingAnswer) AndAlso
           existingAnswer.IndexOf(cleanUrl, StringComparison.OrdinalIgnoreCase) >= 0 Then
            Exit Sub
        End If

        If Not seenUrls.Add(cleanUrl) Then
            Exit Sub
        End If

        results.Add(New ToolSourceLink With {
            .Url = cleanUrl,
            .Title = If(title, "").Trim(),
            .Source = If(source, "").Trim()
        })
    End Sub

    Private Function IsM365RetrievalToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then
            Return False
        End If

        Select Case toolName.Trim().ToLowerInvariant()
            Case "m365_get_mail",
                 "m365_get_mail_thread",
                 "m365_get_file",
                 "m365_get_event",
                 "m365_get_chat_thread",
                 "m365_get_onenote_page"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function GetM365SourceFromToolName(toolName As String) As String
        If String.IsNullOrWhiteSpace(toolName) Then
            Return ""
        End If

        Select Case toolName.Trim().ToLowerInvariant()
            Case "m365_get_mail", "m365_get_mail_thread"
                Return "mail"
            Case "m365_get_file"
                Return "file"
            Case "m365_get_event"
                Return "calendar"
            Case "m365_get_chat_thread"
                Return "teams"
            Case "m365_get_onenote_page"
                Return "onenote"
            Case Else
                Return "m365"
        End Select
    End Function

    Private Function BuildSourceLinkLabel(link As ToolSourceLink) As String
        If link Is Nothing Then Return "Open source"

        Dim title As String = If(link.Title, "").Trim()
        Dim source As String = If(link.Source, "").Trim().ToLowerInvariant()

        If String.IsNullOrWhiteSpace(title) Then
            title = "Open item"
        End If

        Select Case source
            Case "mail"
                Return title & " (e-mail)"
            Case "onedrive"
                Return title & " (OneDrive)"
            Case "sharepoint"
                Return title & " (SharePoint)"
            Case "file"
                Return title & " (file)"
            Case "teams"
                Return title & " (Teams)"
            Case "calendar"
                Return title & " (calendar)"
            Case "onenote"
                Return title & " (OneNote)"
            Case Else
                Return title
        End Select
    End Function

    Private Function EscapeMarkdownLinkText(value As String) As String
        Dim s As String = If(value, "")
        s = s.Replace("\", "\\")
        s = s.Replace("[", "\[")
        s = s.Replace("]", "\]")
        Return s
    End Function

    Private Function IsMCPStreamableToolCall(endpoint As String, apiCall As String) As Boolean
        If String.IsNullOrWhiteSpace(endpoint) Then
            Return False
        End If

        If endpoint.StartsWith(SharedMethods.MCP_STREAMABLE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        If endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If String.IsNullOrWhiteSpace(apiCall) Then
            Return False
        End If

        Try
            Dim requestObj As JObject = JObject.Parse(apiCall)

            Return String.Equals(
                If(requestObj("jsonrpc")?.ToString(), ""),
                "2.0",
                StringComparison.OrdinalIgnoreCase) AndAlso
                String.Equals(
                    If(requestObj("method")?.ToString(), ""),
                    "tools/call",
                    StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Function ShouldRetryMCPAfterUnauthorized(toolConfig As ModelConfig, ex As Exception) As Boolean
        If toolConfig Is Nothing OrElse Not toolConfig.OAuth2 Then Return False
        If ex Is Nothing Then Return False

        Dim message As String = If(ex.Message, "")
        Return message.IndexOf("401", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               message.IndexOf("Invalid or expired access token", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Async Function ForceRefreshToolOAuthToken(toolConfig As ModelConfig, toolName As String) As Task(Of Boolean)
        Try
            toolConfig.DecodedAPI = Await SharedMethods.GetFreshAccessToken(
                _context,
                toolConfig.OAuth2ClientMail,
                toolConfig.OAuth2Scopes,
                toolConfig.APIKey,
                toolConfig.OAuth2Endpoint,
                toolConfig.OAuth2ATExpiry,
                True,
                False,
                forceRefresh:=True).ConfigureAwait(False)

            If String.IsNullOrWhiteSpace(toolConfig.DecodedAPI) Then
                ToolingFileLogger.LogError(
                    "Forced MCP OAuth refresh returned an empty token.",
                    details:=$"ToolName='{toolName}'")
                Return False
            End If

            Return True

        Catch refreshEx As Exception
            ToolingFileLogger.LogError(
                "Forced MCP OAuth refresh failed.",
                details:=$"ToolName='{toolName}'",
                ex:=refreshEx)
            Return False
        End Try
    End Function


#End Region

End Class