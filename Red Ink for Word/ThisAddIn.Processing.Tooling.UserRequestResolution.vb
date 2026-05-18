' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.UserRequestResolution.vb
' Purpose: User language detection and request text extraction for tooling sessions.
'
' Responsibilities:
'  - Detect user's preferred language via LLM classification.
'  - Extract latest user turn from dialog/prompt structures.
'  - Parse BCP-47 language tags and localization preferences.
'  - Support fallback language handling.
'
' External Dependencies:
'  - LLM() for language detection classification.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn


    Private Async Function ResolveToolingUserLanguageAsync(userText As String,
                                                           otherPrompt As String,
                                                           fullPromptOverride As String,
                                                           useSecondAPI As Boolean,
                                                           hideSplash As Boolean) As Task(Of String)
        Dim sourceText As String =
            ResolveLatestUserRequestRaw(userText, otherPrompt, fullPromptOverride)

        sourceText = If(sourceText, "").Trim()
        If sourceText = "" Then Return ""

        If sourceText.Length > 4000 Then
            sourceText = sourceText.Substring(0, 4000)
        End If

        Dim detectionSystemPrompt As String =
            "Determine the language in which the assistant must answer the user's latest request. " &
            "Return ONLY valid JSON in the form {""language"":""...""}. " &
            "Use a concrete runtime language value suitable for later localization, preferably a BCP-47 tag when clear. " &
            "Do not add explanations."

        Dim detectionUserPrompt As String =
            "<USER_ENTRY>" & sourceText & "</USER_ENTRY>"

        Try
            Dim raw As String = Await LLM(
                detectionSystemPrompt,
                detectionUserPrompt,
                "", "", 0,
                useSecondAPI,
                hideSplash,
                "",
                "",
                True)

            If String.IsNullOrWhiteSpace(raw) Then Return ""

            Try
                Dim obj As JObject = JObject.Parse(raw)
                Return If(obj.Value(Of String)("language"), "").Trim()
            Catch
                Return raw.Trim().Trim(""""c)
            End Try
        Catch
            Return ""
        End Try
    End Function




End Class
