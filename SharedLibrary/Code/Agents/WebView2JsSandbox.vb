' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WebView2JsSandbox.vb
' Purpose: Runs untrusted JavaScript inside a hidden, pooled WebView2 instance.
'          Pool size = 1 (reused across calls; per-call lifecycle is just a
'          semaphore + script execution; the underlying control is long-lived).
'
' Threading:
'   - WebView2 lives on the host UI (STA) thread.
'   - Initialize(...) MUST be called by the host once, on the UI thread,
'     supplying a SynchronizationContext bound to that thread.
'
' Security:
'   - Network is DISABLED by default (all WebResourceRequested are short-circuited
'     with 403). Set allow_network=true on the tool to permit fetch().
'   - DevTools/context-menus/host-objects disabled.
'   - User code is wrapped: console.log is captured into a JSON 'logs' field,
'     exceptions become 'ok:false / error', and timeouts return 'error:timeout'.
'   - Return value is JSON-stringified by the wrapper, then unwrapped here.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public NotInheritable Class WebView2JsSandbox

        Private Sub New()
        End Sub

        Private Shared ReadOnly _gate As New SemaphoreSlim(1, 1)
        Private Shared _uiSync As SynchronizationContext = Nothing
        Private Shared _userDataDir As String = Nothing
        Private Shared _hostForm As Form = Nothing
        Private Shared _wv As WebView2 = Nothing
        Private Shared _initialized As Boolean = False

        ''' <summary>Host call (UI thread). syncContext must be a WindowsFormsSynchronizationContext.</summary>
        Public Shared Sub Initialize(syncContext As SynchronizationContext, Optional userDataDir As String = Nothing)
            _uiSync = syncContext
            _userDataDir = If(String.IsNullOrWhiteSpace(userDataDir),
                              Path.Combine(Path.GetTempPath(), "RedInk_JsSandbox"),
                              userDataDir)
        End Sub

        Public Shared ReadOnly Property IsConfigured As Boolean
            Get
                Return _uiSync IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' Runs the supplied JS source. Returns the inner JSON envelope produced by the wrapper:
        '''   { "ok": true|false, "result": ..., "logs": [...], "error"?: "..." }
        ''' On timeout returns { "error": "timeout", "timeout_ms": N }.
        ''' </summary>
        Public Shared Async Function RunAsync(code As String,
                                              Optional timeoutMs As Integer = 15000,
                                              Optional allowNetwork As Boolean = False,
                                              Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)
            If String.IsNullOrWhiteSpace(code) Then
                Return JsonConvert.SerializeObject(New With {Key .error = "missing_code"})
            End If
            If _uiSync Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .error = "sandbox_uninitialized",
                    Key .message = "WebView2JsSandbox.Initialize must be called by the host at startup."})
            End If
            If timeoutMs < 500 Then timeoutMs = 500
            If timeoutMs > 120000 Then timeoutMs = 120000

            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)
            Try
                Await EnsureInitialized(allowNetwork).ConfigureAwait(False)

                Dim wrapped = BuildWrappedScript(code)

                ' Important: ExecuteScriptAsync serializes Promise objects as "{}".
                ' Use DevTools Runtime.evaluate with awaitPromise:=true so async wrapper results are awaited.
                Dim execTask = RunOnUiAsync(Of String)(Function() ExecuteScriptAwaitPromiseAsync(wrapped))
                Dim delay = Task.Delay(timeoutMs, cancellationToken)
                Dim winner = Await Task.WhenAny(execTask, delay).ConfigureAwait(False)

                If winner Is delay Then
                    Dim navTask As Task = Nothing
                    Try
                        navTask = RunOnUiAsync(Sub() _wv.CoreWebView2.NavigateToString("<!doctype html><html></html>"))
                    Catch
                        navTask = Nothing
                    End Try

                    If navTask IsNot Nothing Then
                        Try
                            Await navTask.ConfigureAwait(False)
                        Catch
                        End Try
                    End If

                    Return JsonConvert.SerializeObject(New With {Key .error = "timeout", Key .timeout_ms = timeoutMs})
                End If

                Dim inner As String = Await execTask.ConfigureAwait(False)

                If String.IsNullOrWhiteSpace(inner) Then
                    inner = JsonConvert.SerializeObject(New With {Key .ok = True, Key .result = Nothing})
                End If

                Return inner
            Catch oce As OperationCanceledException
                Return JsonConvert.SerializeObject(New With {Key .error = "cancelled"})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {Key .error = "js_run_failed", Key .message = ex.Message})
            Finally
                _gate.Release()
            End Try
        End Function



        ' --------------------------------------------------------------- wrapper

        Private Shared Function BuildWrappedScript(userCode As String) As String
            Dim literal = JsonConvert.SerializeObject(userCode)
            Return _
"(async () => {" & vbLf &
"  const __logs = [];" & vbLf &
"  console.log = (...a) => { try { __logs.push(a.map(x => x===undefined?'undefined':(x===null?'null':(typeof x==='object'?JSON.stringify(x):String(x)))).join(' ')); } catch(_) {} };" & vbLf &
"  console.error = console.log; console.warn = console.log;" & vbLf &
"  try {" & vbLf &
"    const __fn = new Function('return (async () => { ' + " & literal & " + ' \n })();');" & vbLf &
"    const __r = await __fn();" & vbLf &
"    return JSON.stringify({ ok: true, result: __r === undefined ? null : __r, logs: __logs });" & vbLf &
"  } catch (e) {" & vbLf &
"    return JSON.stringify({ ok: false, error: (e && e.message) ? String(e.message) : String(e), logs: __logs });" & vbLf &
"  }" & vbLf &
"})()"
        End Function

        ' --------------------------------------------------------------- lifecycle

        Private Shared Async Function ExecuteScriptAwaitPromiseAsync(script As String) As Task(Of String)
            Try
                Dim payload As String = New JObject(
            New JProperty("expression", script),
            New JProperty("awaitPromise", True),
            New JProperty("returnByValue", True)
        ).ToString(Formatting.None)

                Dim raw As String = Await _wv.CoreWebView2.CallDevToolsProtocolMethodAsync(
            "Runtime.evaluate",
            payload).ConfigureAwait(True)

                Dim root As JObject = JObject.Parse(raw)

                Dim exceptionToken As JToken = root("exceptionDetails")
                If exceptionToken IsNot Nothing Then
                    Dim message As String = If(
                exceptionToken.SelectToken("exception.description")?.ToString(),
                exceptionToken.SelectToken("text")?.ToString())

                    If String.IsNullOrWhiteSpace(message) Then
                        message = exceptionToken.ToString(Formatting.None)
                    End If

                    Return JsonConvert.SerializeObject(New With {
                Key .ok = False,
                Key .error = message
            })
                End If

                Dim valueToken As JToken = root.SelectToken("result.value")
                If valueToken IsNot Nothing Then
                    If valueToken.Type = JTokenType.String Then
                        Return valueToken.ToString()
                    End If

                    Return valueToken.ToString(Formatting.None)
                End If

                Dim unserializableToken As JToken = root.SelectToken("result.unserializableValue")
                If unserializableToken IsNot Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                Key .ok = True,
                Key .result = unserializableToken.ToString()
            })
                End If

                Return JsonConvert.SerializeObject(New With {
            Key .ok = True,
            Key .result = CType(Nothing, Object)
        })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
            Key .ok = False,
            Key .error = "js_run_eval_failed",
            Key .message = ex.Message
        })
            End Try
        End Function

        Private Shared Async Function EnsureInitialized(allowNetwork As Boolean) As Task
            If _initialized Then
                Await RunOnUiAsync(Sub() ConfigureNetworkPolicy(allowNetwork)).ConfigureAwait(False)
                Return
            End If
            Await RunOnUiAsync(Async Function() As Task
                                   If _hostForm Is Nothing Then
                                       _hostForm = New Form() With {
                                           .Text = "RedInk JS Sandbox",
                                           .FormBorderStyle = FormBorderStyle.None,
                                           .ShowInTaskbar = False,
                                           .StartPosition = FormStartPosition.Manual,
                                           .Location = New Point(-32000, -32000),
                                           .Size = New Size(2, 2),
                                           .Opacity = 0.0
                                       }
                                       AddHandler _hostForm.FormClosing, Sub(s, e) e.Cancel = True
                                       _hostForm.Show()
                                       _hostForm.Hide()
                                   End If
                                   If _wv Is Nothing Then
                                       _wv = New WebView2() With {.Dock = DockStyle.Fill}
                                       _hostForm.Controls.Add(_wv)
                                       Dim env As CoreWebView2Environment = Nothing
                                       Dim envOk As Boolean = True
                                       Try
                                           env = Await CoreWebView2Environment.CreateAsync(Nothing, _userDataDir).ConfigureAwait(True)
                                       Catch
                                           envOk = False
                                       End Try
                                       If Not envOk Then
                                           env = Await CoreWebView2Environment.CreateAsync().ConfigureAwait(True)
                                       End If
                                       Await _wv.EnsureCoreWebView2Async(env).ConfigureAwait(True)
                                       Await _wv.EnsureCoreWebView2Async(env).ConfigureAwait(True)
                                       Dim s = _wv.CoreWebView2.Settings
                                       s.AreDevToolsEnabled = False
                                       s.AreDefaultContextMenusEnabled = False
                                       s.IsStatusBarEnabled = False
                                       s.IsZoomControlEnabled = False
                                       s.IsBuiltInErrorPageEnabled = False
                                       ConfigureNetworkPolicy(allowNetwork)
                                       _wv.CoreWebView2.NavigateToString("<!doctype html><html><head><meta charset=""utf-8""><title>sandbox</title></head><body></body></html>")
                                   End If
                                   _initialized = True
                               End Function).ConfigureAwait(False)
        End Function

        Private Shared Sub ConfigureNetworkPolicy(allow As Boolean)
            Try
                RemoveHandler _wv.CoreWebView2.WebResourceRequested, AddressOf BlockAll
            Catch
            End Try
            Try
                _wv.CoreWebView2.RemoveWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All)
            Catch
            End Try
            If Not allow Then
                Try
                    _wv.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All)
                    AddHandler _wv.CoreWebView2.WebResourceRequested, AddressOf BlockAll
                Catch
                End Try
            End If
        End Sub

        Private Shared Sub BlockAll(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
            Try
                Dim env = _wv.CoreWebView2.Environment
                e.Response = env.CreateWebResourceResponse(Nothing, 403, "Forbidden", "Content-Type: text/plain")
            Catch
            End Try
        End Sub

        ' --------------------------------------------------------------- UI marshalling

        Private Shared Function RunOnUiAsync(action As Action) As Task
            Dim tcs As New TaskCompletionSource(Of Boolean)()
            _uiSync.Post(Sub()
                             Try : action() : tcs.SetResult(True)
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

        Private Shared Function RunOnUiAsync(Of T)(func As Func(Of Task(Of T))) As Task(Of T)
            Dim tcs As New TaskCompletionSource(Of T)()
            _uiSync.Post(Async Sub()
                             Try : tcs.SetResult(Await func().ConfigureAwait(True))
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

        Private Shared Function RunOnUiAsync(func As Func(Of Task)) As Task
            Dim tcs As New TaskCompletionSource(Of Boolean)()
            _uiSync.Post(Async Sub()
                             Try : Await func().ConfigureAwait(True) : tcs.SetResult(True)
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

    End Class

End Namespace