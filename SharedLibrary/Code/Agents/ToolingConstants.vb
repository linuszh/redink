' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingConstants.vb
' Purpose: Central constants for the tooling pipeline shared by Outlook, Word
'          and Excel. These replace the per-host scattered defaults.
'
' Constants:
'  - DefaultMaxToolIterations = 50 (unified across all hosts).
'  - LlmTimeoutBufferSeconds = 60 (added per iteration).
'  - MaxContinuationRetries = 5 (repair attempts for recovery).
'  - SubAgentLargeToolResponseThresholdChars = 30000 (compaction trigger).
'  - SubAgentLargeToolResponseExcerptChars = 8000 (excerpt size when compacted).
'  - MaxLocalizableBlockedFinalChars = 1500 (host post-translation willing threshold).
'  - MaxFallbackToolsListedInGuard = 6 (fallback tools listed in guard prompts).
' =============================================================================


Option Explicit On
Option Strict On

Namespace Agents

    ''' <summary>
    ''' Central tooling-pipeline constants. Per Q7 the user wants a single, unified default
    ''' across Outlook and Word. Hosts may still override via INI_ToolingMaximumIterations
    ''' at runtime, but the default seeded into INI must come from here.
    ''' </summary>
    Public Module ToolingConstants

        ''' <summary>Unified default for INI_ToolingMaximumIterations across all hosts.</summary>
        Public Const DefaultMaxToolIterations As Integer = 50

        ''' <summary>Additional seconds added to the configured LLM timeout per iteration.</summary>
        Public Const LlmTimeoutBufferSeconds As Integer = 60

        ''' <summary>Maximum repair attempts for premature-text / invalid-turn / empty-response recovery.</summary>
        Public Const MaxContinuationRetries As Integer = 5

        ''' <summary>Char threshold over which a sub-agent tool response is compacted for model replay.</summary>
        Public Const SubAgentLargeToolResponseThresholdChars As Integer = 30000

        ''' <summary>Excerpt size kept when sub-agent tool responses are compacted.</summary>
        Public Const SubAgentLargeToolResponseExcerptChars As Integer = 8000

        ''' <summary>Maximum length of a blocked-final string that the host is willing to translate.</summary>
        Public Const MaxLocalizableBlockedFinalChars As Integer = 1500

        ''' <summary>Maximum number of fallback tools to list in a deliverable-fallback guard prompt.</summary>
        Public Const MaxFallbackToolsListedInGuard As Integer = 6

    End Module

End Namespace