' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: LanguageContract.vb
' Purpose: Centralizes (1) hard system-prompt rule that final user-facing
'          prose must be in the detected user language, and (2) post-localization
'          decision for blocked finals (per Q3: only blocked finals are eligible).
'
' Rules:
'  - System prompt fragment mandates all FINAL prose be in user's language.
'  - Tool arguments and JSON envelopes remain in English.
'  - Only BLOCKED finals are post-localized (if short enough); COMPLETE finals trust.
'  - Heuristic detection of target language in prose (accents, common words).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions

Namespace Agents

    Public Module LanguageContract

        ''' <summary>
        ''' Builds the language-contract block that MUST be appended to every iteration's
        ''' system prompt. This ensures the model answers in the user's language regardless
        ''' of guard-prompt language or tool-output language.
        ''' </summary>
        Public Function BuildSystemPromptFragment(userLanguage As String) As String
            Dim lang As String = If(userLanguage, "").Trim()
            If lang = "" Then Return ""

            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("[LANGUAGE CONTRACT — MANDATORY]")
            sb.AppendLine("- The user's language for THIS run is: " & lang)
            sb.AppendLine("- All FINAL user-facing prose (the message the user reads) MUST be written in " & lang & ".")
            sb.AppendLine("- This applies to both successful and blocked final answers, and to any explanation of failures or limitations.")
            sb.AppendLine("- Tool call arguments, JSON envelopes, structured payloads, and the <TASK_STATUS> footer remain in English regardless of user language.")
            sb.AppendLine("- Do NOT switch the language of your final prose just because tool output, host guard prompts, or system text are in English.")
            sb.AppendLine("[/LANGUAGE CONTRACT]")
            Return sb.ToString().TrimEnd()
        End Function

        ''' <summary>
        ''' Per Q3, only BLOCKED finals are eligible for host-side post-translation.
        ''' Complete finals trust the language contract above and are not retranslated.
        ''' Additionally the prose must be short enough to translate cheaply.
        ''' </summary>
        Public Function ShouldPostLocalizeBlockedFinal(prose As String,
                                                       userLanguage As String,
                                                       maxLocalizableChars As Integer) As Boolean
            If String.IsNullOrWhiteSpace(prose) Then Return False
            If String.IsNullOrWhiteSpace(userLanguage) Then Return False
            If LooksLikeEnglish(userLanguage) Then Return False
            If prose.Length > maxLocalizableChars Then Return False
            ' If the prose already looks like it is in the target language we skip,
            ' otherwise we always translate (cheap, predictable).
            Return Not ProseLooksLikeTargetLanguage(prose, userLanguage)
        End Function

        Private Function LooksLikeEnglish(language As String) As Boolean
            Dim l As String = If(language, "").Trim().ToLowerInvariant()
            If l = "" Then Return False
            If l = "en" Then Return True
            If l.StartsWith("en-", StringComparison.OrdinalIgnoreCase) Then Return True
            If l = "english" Then Return True
            Return False
        End Function

        ''' <summary>
        ''' Cheap heuristic: counts language-specific letter clusters. Not perfect, but
        ''' good enough to skip retranslation for obvious matches.
        ''' </summary>
        Private Function ProseLooksLikeTargetLanguage(prose As String, language As String) As Boolean
            Dim l As String = If(language, "").Trim().ToLowerInvariant()
            If l = "" Then Return False

            If l.StartsWith("de", StringComparison.OrdinalIgnoreCase) OrElse l = "german" Then
                Return Regex.IsMatch(prose, "[äöüÄÖÜß]") OrElse
                       Regex.IsMatch(prose, "\b(?:nicht|ich|und|aber|kann|werden|wurde|leider)\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("fr", StringComparison.OrdinalIgnoreCase) OrElse l = "french" Then
                Return Regex.IsMatch(prose, "[àâçéèêëîïôûùüÿñæœ]") OrElse
                       Regex.IsMatch(prose, "\b(?:je|ne|pas|nous|mais|peut|impossible|d(?:é|e)sol(?:é|e))\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("it", StringComparison.OrdinalIgnoreCase) OrElse l = "italian" Then
                Return Regex.IsMatch(prose, "\b(?:non|sono|ma|posso|impossibile|spiacente)\b", RegexOptions.IgnoreCase)
            End If

            If l.StartsWith("es", StringComparison.OrdinalIgnoreCase) OrElse l = "spanish" Then
                Return Regex.IsMatch(prose, "[ñáéíóúü¿¡]") OrElse
                       Regex.IsMatch(prose, "\b(?:no|pero|puedo|imposible|lo\s+siento)\b", RegexOptions.IgnoreCase)
            End If

            Return False
        End Function

    End Module

End Namespace