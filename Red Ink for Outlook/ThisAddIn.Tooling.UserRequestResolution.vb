' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Tooling.UserRequestResolution.vb
' Purpose: Extracts and resolves the authoritative user request from multi-source prompts.
'          Analyzes user intent to determine if a deliverable artifact is explicitly required.
'
' Architecture:
'  - Latest User Request Extraction:
'      - ExtractLatestUserTurnFromDialog(): Parses structured dialog format [USER]/[ASSISTANT]/[TOOL].
'          - Matches backward from end to find latest [USER] block.
'          - Handles multi-turn dialog with optional </DIALOG> closing tag.
'          - Falls through to raw text if no dialog markers present.
'  - Request Resolution Priority:
'      - ResolveLatestUserRequestRaw(): Selects authoritative source in order:
'          1. fullPromptOverride (explicit full prompt provided).
'          2. userText (primary user input).
'          3. otherPrompt (fallback system/context prompt).
'          4. "" (empty if no source available).
'  - Task Summary Resolution:
'      - ResolveOptionalHostTaskSummary(): Builds optional task context summary.
'          - Combines userText, fullPromptOverride, otherPrompt, insertDocs, slideInsert, bubblesText.
'          - Calls BuildToolSelectionHintText() to extract task hints.
'          - Deduplicates against latestUserRequestRaw to avoid redundancy.
'  - Deliverable Intent Detection:
'      - RequestExplicitlyRequiresCreatedDeliverable(): Analyzes for artifact creation patterns.
'          - Regex: Tool names like "create_word_document", "create_pdf", etc.
'          - Regex: Strong artifact verb + noun pairs (verb within 80 chars of noun).
'          - Regex: Document artifact patterns ("create document", "erstellen dokument").
'          - Supports English and German terminology for internationalization.
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

Partial Public Class ThisAddIn



    Private Function ExtractLatestUserTurnFromDialog(promptText As String) As String
        Dim raw As String = If(promptText, "").Trim()
        If raw = "" Then
            Return ""
        End If

        If raw.IndexOf("<DIALOG>", StringComparison.OrdinalIgnoreCase) < 0 OrElse
           raw.IndexOf("[USER]", StringComparison.OrdinalIgnoreCase) < 0 Then
            Return raw
        End If

        Dim matches As MatchCollection =
            Regex.Matches(
                raw,
                "\[USER\]\s*(?<body>.*?)(?=\s*\[(?:USER|ASSISTANT|SYSTEM|TOOL)\]|\s*</DIALOG>|\z)",
                RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.CultureInvariant)

        If matches.Count = 0 Then
            Return raw
        End If

        Dim latest As String = matches(matches.Count - 1).Groups("body").Value.Trim()
        latest = Regex.Replace(latest, "\s*</DIALOG>\s*$", "", RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant).Trim()

        If latest = "" Then
            Return raw
        End If

        Return latest
    End Function

    Private Function ResolveLatestUserRequestRaw(userText As String,
                                                 otherPrompt As String,
                                                 fullPromptOverride As String) As String
        If Not String.IsNullOrWhiteSpace(fullPromptOverride) Then
            Return ExtractLatestUserTurnFromDialog(fullPromptOverride)
        End If

        If Not String.IsNullOrWhiteSpace(userText) Then
            Return ExtractLatestUserTurnFromDialog(userText)
        End If

        If Not String.IsNullOrWhiteSpace(otherPrompt) Then
            Return ExtractLatestUserTurnFromDialog(otherPrompt)
        End If

        Return ""
    End Function

    Private Function RequestExplicitlyRequiresCreatedDeliverable(latestUserRequestRaw As String) As Boolean
        Dim raw As String = If(latestUserRequestRaw, "").Trim()
        If raw = "" Then
            Return False
        End If

        Dim normalized As String = raw.ToLowerInvariant()

        If Regex.IsMatch(
            normalized,
            "\b(create_word_document|create_pdf|create_excel_spreadsheet|create_powerpoint|create_code_file|word_to_pdf|word_save_as|workspace_write|text_write|word_write)\b",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Dim strongArtifactNounPattern As String =
            "(?:\b(?:pdf|docx|xlsx|pptx|spreadsheet|workbook|presentation|powerpoint|file|datei|attachment|anhang|workspace|downloadable|ausgabedatei)\b|word[- ]?(?:document|dokument)|output\s+file)"

        Dim strongArtifactVerbPattern As String =
            "\b(?:create|generate|produce|make|write|save|export|attach|store|persist|download|erstellen|generieren|erzeugen|machen|schreiben|speichern|exportieren|anhängen|ablegen|herunterladen)\b"

        If Regex.IsMatch(
            normalized,
            strongArtifactVerbPattern & "[\s\S]{0,80}" & strongArtifactNounPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            strongArtifactNounPattern & "[\s\S]{0,80}" & strongArtifactVerbPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Dim documentArtifactVerbPattern As String =
            "\b(?:create|generate|save|export|attach|store|persist|erstellen|generieren|speichern|exportieren|anhängen|ablegen)\b"

        Dim documentArtifactNounPattern As String =
            "\b(?:document|dokument)\b"

        If Regex.IsMatch(
            normalized,
            documentArtifactVerbPattern & "[\s\S]{0,80}" & documentArtifactNounPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            documentArtifactNounPattern & "[\s\S]{0,80}" & documentArtifactVerbPattern,
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        If Regex.IsMatch(
            normalized,
            "\b(?:as|als)\s+(?:a\s+|an\s+|eine\s+|einen\s+)?(?:pdf|docx|xlsx|pptx|word[- ]?(?:document|dokument)|spreadsheet|presentation|powerpoint|file|datei)\b",
            RegexOptions.IgnoreCase Or RegexOptions.CultureInvariant) Then
            Return True
        End If

        Return False
    End Function


End Class
