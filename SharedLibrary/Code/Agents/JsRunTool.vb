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
"""description"":""Run sandboxed JavaScript inside a hidden WebView2. The 'code' is executed as the body of an async function; use 'return' to produce a value (it will be JSON-serialized). console.log/console.warn/console.error output is captured. Network access is DISABLED by default (set allow_network=true to permit fetch). Default timeout 15s. Use for math, parsing, regex, JSON manipulation, and other deterministic algorithms instead of asking the model to compute them in-band.""," &
"""parameters"":{""type"":""object""," &
"""properties"":{" &
"""code"":{""type"":""string"",""description"":""JavaScript source.""}," &
"""timeout_ms"":{""type"":""integer"",""description"":""Wall-clock limit (500..120000; default 15000).""}," &
"""allow_network"":{""type"":""boolean"",""description"":""Permit network requests (default false).""}}," &
"""required"":[""code""]}}"

            Return New ModelConfig() With {
                .ToolName = ToolName,
                .ToolDefinition = def,
                .ToolInstructionsPrompt = ToolName & ": Run sandboxed JavaScript and receive {ok, result, logs} (or {ok:false, error}).",
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