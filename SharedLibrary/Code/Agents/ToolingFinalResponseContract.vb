' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingFinalResponseContract.vb
' Purpose: Defines the contract for the final response of tooling operations.
' =============================================================================


Option Explicit On
Option Strict On

Namespace Agents

    Public Enum ToolingFinalResponseContract
        UserFacingTaskStatus = 0
        RawCallerText = 1
    End Enum

    Public Module ToolingFinalResponseContractHelpers

        Public Function RequiresTaskStatusFooter(contract As ToolingFinalResponseContract) As Boolean
            Return contract = ToolingFinalResponseContract.UserFacingTaskStatus
        End Function

        Public Function FormatToolingFinalResponseContract(contract As ToolingFinalResponseContract) As String
            Select Case contract
                Case ToolingFinalResponseContract.RawCallerText
                    Return "raw_caller_text"
                Case Else
                    Return "user_facing_task_status"
            End Select
        End Function

    End Module

End Namespace