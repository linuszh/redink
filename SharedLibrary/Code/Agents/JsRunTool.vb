' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: JsRunTool.vb
' Purpose: ModelConfig + dispatcher entry for the js_run tool. The actual
'          execution happens in WebView2JsSandbox.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports SharedLibrary.SharedLibrary

Namespace Agents

    Public NotInheritable Class JsRunTool

        Private Sub New()
        End Sub

        Public Const ToolName As String = "js_run"

        Public Shared Function IsJsTool(name As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(name) AndAlso
                   String.Equals(name, ToolName, StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function Build() As ModelConfig
            Dim def =
"{""name"":""" & ToolName & """," &
"""description"":""Run sandboxed JavaScript inside a hidden WebView2. The 'code' parameter is executed as the BODY of an async function. Therefore, do NOT wrap it in 'async function ... { }' and do NOT invent wrapper parameters such as browser_mode. Always produce the final value with an explicit top-level 'return'. console.log/console.warn/console.error output is captured. Network access is DISABLED by default; set allow_network=true to permit fetch or controlled page navigation. Browser mode: set navigate_url to load a page into the hidden browser before the code runs against the live DOM. Optional wait_for_selector and wait_after_load_ms may be used. Security: only absolute http/https URLs are allowed; localhost, loopback, and private-network destinations are blocked. Default timeout 15s.""," &
"""parameters"":{""type"":""object""," &
"""properties"":{" &
"""code"":{""type"":""string"",""description"":""JavaScript source. IMPORTANT: this is already the BODY of an async function. Write statements directly and end with a top-level return of the final value.""}," &
"""timeout_ms"":{""type"":""integer"",""description"":""Wall-clock limit (500..120000; default 15000).""}," &
"""allow_network"":{""type"":""boolean"",""description"":""Permit network requests or browser navigation (default false).""}," &
"""navigate_url"":{""type"":""string"",""description"":""Optional absolute http/https URL to open in the hidden browser before the code runs. Requires allow_network=true.""}," &
"""wait_after_load_ms"":{""type"":""integer"",""description"":""Optional extra delay after page load before executing code (0..30000; default 1500 when navigate_url is set).""}," &
"""wait_for_selector"":{""type"":""string"",""description"":""Optional CSS selector to wait for before running the code. Shadow-DOM roots are searched recursively.""}}," &
"""required"":[""code""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt =
                    ToolName & ": Run sandboxed JavaScript and receive {ok, result, logs} or {ok:false, error}. " &
                    "IMPORTANT: 'code' is already the BODY of an async function. Do not declare 'async function ...'. " &
                    "Always return the final value explicitly at top level, for example: " &
                    "'const links = [...document.querySelectorAll(""a[href]"")].map(a => a.href); return links;'. " &
                    "For page DOM access, use allow_network=true and navigate_url='https://...'. Do not invent browser_mode.",
                .ModelDescription = "JS sandbox (WebView2)",
                .Tool = True,
                .ToolPriority = 860,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Async Function ExecuteAsync(arguments As IDictionary(Of String, Object),
                                                  Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)
            Try
                Return Await WebView2JsSandbox.RunAsync(
                    code:=GetStr(arguments, "code"),
                    timeoutMs:=GetInt(arguments, "timeout_ms", 15000),
                    allowNetwork:=GetBool(arguments, "allow_network", False),
                    navigateUrl:=GetStr(arguments, "navigate_url"),
                    waitAfterLoadMs:=GetInt(arguments, "wait_after_load_ms", 1500),
                    waitForSelector:=GetStr(arguments, "wait_for_selector"),
                    cancellationToken:=cancellationToken).ConfigureAwait(False)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "js_run_failed", Key .message = ex.Message})
            End Try
        End Function

        Private Shared Function GetStr(args As IDictionary(Of String, Object), name As String) As String
            If args Is Nothing Then Return ""
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return ""
            Return System.Convert.ToString(v)
        End Function

        Private Shared Function GetInt(args As IDictionary(Of String, Object), name As String, dflt As Integer) As Integer
            If args Is Nothing Then Return dflt
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return dflt
            Try : Return System.Convert.ToInt32(v) : Catch
                Dim n As Integer
                If Integer.TryParse(System.Convert.ToString(v), n) Then Return n
                Return dflt
            End Try
        End Function

        Private Shared Function GetBool(args As IDictionary(Of String, Object), name As String, dflt As Boolean) As Boolean
            If args Is Nothing Then Return dflt
            Dim v As Object = Nothing
            If Not args.TryGetValue(name, v) OrElse v Is Nothing Then Return dflt
            Try : Return System.Convert.ToBoolean(v) : Catch
                Select Case System.Convert.ToString(v).Trim().ToLowerInvariant()
                    Case "true", "1", "yes" : Return True
                    Case "false", "0", "no" : Return False
                    Case Else : Return dflt
                End Select
            End Try
        End Function

    End Class

End Namespace