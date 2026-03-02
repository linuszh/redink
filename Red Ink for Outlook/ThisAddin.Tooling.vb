' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddin.Tooling.vb
' Purpose:
'   Implements a model-agnostic tool/function-calling loop for LLM interactions.
'   Detects tool calls in model output, executes internal/external tools, and
'   injects tool responses into subsequent model iterations until a final answer
'   is produced or limits/cancellation are reached.
'
' Responsibilities:
'   - Tool loop orchestration via `ExecuteToolingLoop`:
'       * Build enhanced system prompt with tool usage guidance.
'       * Generate model-specific tool definition payloads for `INI_APICall_ToolInstructions_2`.
'       * Invoke `LLM(...)` iteratively, detect tool calls, and inject responses via
'         `INI_APICall_ToolResponses_2`.
'       * Enforce iteration limits, timeouts, cancellation, and forced final synthesis.
'   - Tool call parsing:
'       * Detect tool calls via regex (`ContainsToolCalls`) using tooling model patterns.
'       * Extract tool calls from JSON responses (`ExtractToolCalls`) via a model-provided map.
'   - Tool execution:
'       * Internal tool: `retrieve_web_content` with SSRF safeguards (WebView2 or HTTP fallback).
'       * External tools: apply selected `ModelConfig`, build API call payloads, force JSON
'         response mode, invoke the model, and restore prior config.
'       * AutoPilot routing for internal tools when AutoPilot is active.
'   - Tool selection and persistence:
'       * Load tool configurations from `INI_SpecialServicePath`.
'       * Add a built-in web retrieval tool.
'       * Persist selected tool names and restore them on demand.
'   - Diagnostics:
'       * Optional per-run file logging via `ToolingFileLogger` (controlled by `INI_APIDebug`).
'       * Optional UI logging via `LogWindow` when enabled.
'
' Notes:
'   - This file is a partial definition of `ThisAddIn` and depends on shared configuration
'     values (INI variables), `_context`, and the `LLM(...)` entry point provided elsewhere.
'   - Web retrieval uses WebView2 when enabled to capture JavaScript-rendered content.
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
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides tooling support helpers for model-agnostic tool/function calling in LLM interactions.
''' </summary>
Partial Public Class ThisAddIn

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
        "retrieve_web_content: Retrieves text content from one or more URLs. Use this to fetch information from websites when needed."

    ''' <summary>
    ''' Canonical tool definition JSON used for the internal web tool.
    ''' </summary>
    Private Const InternalWebToolDefinition As String =
        "{""name"":""retrieve_web_content"",""description"":""Retrieves the text content from one or more web URLs"",""parameters"":{""type"":""object"",""properties"":{""urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Array of URLs to retrieve content from""}},""required"":[""urls""]}}"

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
                                                                ' Simple extraction - get full body text
                                                                Dim extractScript As String = "
                                                (function() {
                                                    // Remove script/style/noscript to reduce noise
                                                    var toRemove = document.querySelectorAll('script, style, noscript, nav, footer, header');
                                                    toRemove.forEach(function(el) { try { el.remove(); } catch(e) {} });

                                                    // Get body text
                                                    var text = document.body ? (document.body.innerText || '') : '';

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

    ''' <summary>User-friendly name for the tooling feature.</summary>
    Public Const ToolFriendlyName As String = "Sources"

    ''' <summary>Auto-close delay for tooling log window.</summary>
    Public Shared Property ToolingLog_AutoCloseDefaultSeconds As Integer = 30

    ''' <summary>Selected tool names for persistence.</summary>
    Public Shared Property SelectedToolNames As List(Of String) = New List(Of String)()


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

        ''' <summary>
        ''' Initializes a new tool execution context with default collections and limits.
        ''' </summary>
        Public Sub New()
            SelectedTools = New List(Of ModelConfig)()
            AllToolResponses = New List(Of ToolResponse)()
            LogEntries = New List(Of String)()
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
                            Else
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

                        Dim toolResponse = Await ExecuteToolCall(tc, toolConfig, context, cancellationToken)
                        toolResponse.OriginalCallJson = tc.RawJson
                        context.AllToolResponses.Add(toolResponse)

                        If Not toolResponse.Success Then
                            context.LogError(
                                $"Tool error ({tc.ToolName}): {toolResponse.ErrorMessage}",
                                details:=$"CallId={tc.CallId}; RawCall={tc.RawJson}")

                            Select Case toolConfig.ToolErrorHandling?.ToLowerInvariant()
                                Case "abort"
                                    context.LogError("Aborting due to tool error (ToolErrorHandling=abort)")
                                    ShowCustomMessageBox($"Tool execution failed: {toolResponse.ErrorMessage}")
                                    ToolingFileLogger.EndSession(False, $"Tool error: {toolResponse.ErrorMessage}")
                                    Return ""
                                Case "retry"
                                    context.LogWarn("Will retry on next iteration (ToolErrorHandling=retry)")
                                Case Else
                                    context.LogWarn("Skipping tool error (ToolErrorHandling=skip)")
                            End Select
                        Else
                            context.Log($"Tool completed successfully ({toolResponse.Response?.Length} chars)", "success")
                        End If
                    Next

                    Dim toolResponses = BuildToolResponsesForModel(context.AllToolResponses, context.ToolingModel)
                    INI_APICall_ToolResponses_2 = toolResponses
                    context.Log("Tool responses prepared for next iteration")

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

            If context.IsCancelled OrElse cancellationToken.IsCancellationRequested Then
                context.LogWarn("Session cancelled")
                ToolingFileLogger.EndSession(False, "Cancelled")
                Return If(cancellationToken.IsCancellationRequested, "Operation was canceled by the user.", "")
            End If

            If iteration >= context.MaxIterations Then
                context.LogWarn($"Maximum iterations ({context.MaxIterations}) reached")
                ShowCustomMessageBox($"Maximum tool iterations ({context.MaxIterations}) reached. The response may be incomplete.")
                ToolingFileLogger.LogWarn("Maximum iterations reached.", details:=$"MaxIterations={context.MaxIterations}")
            End If

            context.Log("=== Session Summary ===")
            context.Log($"Total iterations: {iteration}")
            context.Log($"Total tool calls: {context.AllToolResponses.Count}")
            Dim successCount As Integer = context.AllToolResponses.Where(Function(r) r.Success).Count()
            Dim failedCount As Integer = context.AllToolResponses.Where(Function(r) Not r.Success).Count()
            context.Log($"Successful: {successCount}", If(failedCount = 0, "success", "step"))
            context.Log($"Failed: {failedCount}", If(failedCount = 0, "step", "warn"))

            ToolingFileLogger.EndSession(True, $"Iterations: {iteration}, Tool calls: {context.AllToolResponses.Count}, Success: {successCount}, Failed: {failedCount}")
            Return currentResponse

        Catch ex As OperationCanceledException
            context.LogWarn("Operation cancelled")
            ToolingFileLogger.EndSession(False, "Cancelled")
            Return "Operation was canceled by the user."

        Catch ex As Exception
            context.LogError($"Error in tooling loop: {ex.Message}", ex:=ex)
            ShowCustomMessageBox($"Error during tool execution: {ex.Message}")
            ToolingFileLogger.EndSession(False, $"Exception: {ex.Message}", ex:=ex)
            Return ""
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

            ' Clear the context reference so ApDashboardLog stops routing
            _activeToolingContext = Nothing
        End Try
    End Function
#End Region

#Region "Tooling Helper Functions"


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
            Dim parameters As String = If(jDef("parameters")?.ToString(), "{}")

            Dim result As String = template
            result = result.Replace("{name}", name)
            result = result.Replace("{description}", description)
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

        Dim sortedTools = selectedTools.OrderBy(Function(t) t.ToolPriority).ToList()

        For Each tool In sortedTools
            If Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt) Then
                sb.AppendLine()
                sb.AppendLine($"- {tool.ToolInstructionsPrompt}")
            End If
        Next

        Return sb.ToString()
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
    ''' Executes a single tool call using either the internal tool implementation or an external tool configuration.
    ''' </summary>
    ''' <param name="toolCall">Tool call extracted from the LLM response.</param>
    ''' <param name="toolConfig">Tool configuration selected for this call.</param>
    ''' <param name="context">Tool execution context for logging and state collection.</param>
    ''' <param name="cancellationToken">Optional cancellation token to abort execution.</param>
    ''' <returns>Tool execution response.</returns>
    ''' <summary>
    ''' Executes a single tool call using either the internal tool implementation or an external tool configuration.
    ''' </summary>
    Public Async Function ExecuteToolCall(toolCall As ToolCall, toolConfig As ModelConfig, context As ToolExecutionContext, Optional cancellationToken As System.Threading.CancellationToken = Nothing) As Task(Of ToolResponse)

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
                    If valueStr.Length > maxLength Then
                        valueStr = valueStr.Substring(0, maxLength - 3) & "..."
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
            ' Check cancellation before starting
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

            For Each kvp In toolCall.Arguments
                Dim placeholder = "{" & kvp.Key & "}"
                Dim value As String
                If kvp.Value Is Nothing Then
                    value = ""
                ElseIf TypeOf kvp.Value Is Boolean Then
                    ' JSON requires lowercase true/false
                    value = If(CBool(kvp.Value), "true", "false")
                ElseIf TypeOf kvp.Value Is JToken Then
                    Dim jt = DirectCast(kvp.Value, JToken)
                    If jt.Type = JTokenType.String Then
                        ' Escape the string value for safe JSON embedding in the template
                        value = JsonConvert.ToString(jt.Value(Of String)())
                        ' JsonConvert.ToString wraps in quotes → strip the outer quotes
                        value = value.Substring(1, value.Length - 2)
                    Else
                        ' Preserve JSON token as-is (handles nested objects/arrays)
                        value = jt.ToString(Formatting.None)
                    End If
                ElseIf TypeOf kvp.Value Is Double OrElse TypeOf kvp.Value Is Single OrElse TypeOf kvp.Value Is Decimal Then
                    ' Ensure invariant culture (dot decimal separator)
                    value = System.Convert.ToDouble(kvp.Value).ToString(Globalization.CultureInfo.InvariantCulture)
                ElseIf TypeOf kvp.Value Is Long OrElse TypeOf kvp.Value Is Integer OrElse TypeOf kvp.Value Is Short Then
                    value = System.Convert.ToInt64(kvp.Value).ToString(Globalization.CultureInfo.InvariantCulture)
                Else
                    ' Escape the string for safe JSON embedding (handles ", \, newlines, etc.)
                    Dim raw = kvp.Value.ToString()
                    Dim escaped = JsonConvert.ToString(raw)
                    ' JsonConvert.ToString wraps in quotes → strip the outer quotes
                    value = escaped.Substring(1, escaped.Length - 2)
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
                    Dim defaultValue As String = ""
                    If toolDefaults IsNot Nothing AndAlso toolDefaults.ContainsKey(placeholderName) Then
                        defaultValue = toolDefaults(placeholderName)
                    End If
                    apiCall = apiCall.Replace(m.Value, defaultValue)
                Next

                ' Re-check: if anything remains unreplaced after applying defaults, that's a real error.
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

            ' ── SSE transport: full round-trip bypassing LLM() ───────────
            If Not String.IsNullOrWhiteSpace(toolConfig.Endpoint) AndAlso
               toolConfig.Endpoint.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then

                Dim sseBase = toolConfig.Endpoint.Substring(SharedMethods.MCP_SSE_PREFIX.Length)
                Dim resolvedHeaderB = If(toolConfig.HeaderB, "").Replace("{apikey}", If(toolConfig.DecodedAPI, ""))

                context.Log($"SSE transport: executing tool {toolCall.ToolName} via {sseBase}")
                ToolingFileLogger.LogStep($"SSE round-trip for {toolCall.ToolName} at {sseBase}")
                ToolingFileLogger.LogStep($"SSE request body: {apiCall}")

                Try
                    Dim rawResult = Await SharedMethods.ExecuteMCPSSEToolCall(
                        sseBase, apiCall,
                        If(toolConfig.HeaderA, ""), resolvedHeaderB,
                        CInt(Math.Min(If(toolConfig.Timeout > 0, toolConfig.Timeout, 60000L), Integer.MaxValue)))

                    ToolingFileLogger.LogRawResponseStub($"SSE tool result ({toolCall.ToolName})", rawResult)

                    If Not String.IsNullOrWhiteSpace(rawResult) Then
                        response.Response = rawResult
                        response.Success = True
                    Else
                        response.Success = False
                        response.ErrorMessage = "Empty response from SSE tool service"
                        ToolingFileLogger.LogError("Empty SSE response.", details:=$"ToolName='{toolCall.ToolName}'")
                    End If

                Catch ex As Exception
                    response.Success = False
                    response.ErrorMessage = $"SSE tool call failed: {ex.Message}"
                    ToolingFileLogger.LogError("SSE tool call failed.",
                        details:=$"ToolName='{toolCall.ToolName}'; SseBase='{sseBase}'", ex:=ex)
                End Try

                Return response
            End If

            ' ── Standard transport: route through LLM() ──────────────────
            Dim backupConfig = GetCurrentConfig(_context)

            Try
                ' Check cancellation before applying config
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

                ' Pass cancellation token to LLM call
                Dim result = Await LLM("", "", "", "", 0, True, True, "", "", cancellationToken, EnsureUI:=False)

                ToolingFileLogger.LogRawResponseStub($"Tool LLM() result ({toolCall.ToolName})", result)

                _context.INI_Response_2 = originalResponse

                If Not String.IsNullOrWhiteSpace(result) Then
                    response.Response = result
                    response.Success = True
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
        "Select the sources you want to make available to the model:")

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
    ''' Returns all available tools by loading external tools from <c>INI_SpecialServicePath</c> and adding the internal web tool.
    ''' </summary>
    ''' <returns>List of available tools.</returns>
    Public Function GetAvailableTools() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        If Not String.IsNullOrWhiteSpace(INI_SpecialServicePath) Then
            Dim externalTools = LoadToolingServices(INI_SpecialServicePath, True)
            tools.AddRange(externalTools)
        End If

        tools.Add(GetInternalWebTool())
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
    Public Function SelectToolsForSession(Optional forceDialog As Boolean = False, Optional FriendlyName As String = ToolFriendlyName) As List(Of ModelConfig)
        Dim availableTools = GetAvailableTools()

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

        ' Only preselect all on first use (no persisted selection yet).
        Return ShowToolSelectionDialog(availableTools, preselectAll:=Not hasPersistedSelection, FriendlyName)
    End Function

#End Region

End Class