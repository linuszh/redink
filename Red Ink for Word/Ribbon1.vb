' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Win32
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary

Public Class Ribbon1

    Private Enum OfficeTheme
        Unknown
        Light
        Dark
    End Enum

    Private Sub ApplyThemeAwareMenuIcon()
        Try
            Dim theme = DetectOfficeTheme()
            Select Case theme
                Case OfficeTheme.Light
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Medium)
                Case Else
                    Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            End Select
            Menu1.ShowImage = True
        Catch
            Menu1.Image = SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)
            Menu1.ShowImage = True
        End Try
    End Sub


    Private Function DetectOfficeTheme() As OfficeTheme
        Const registryPath As String = "Software\Microsoft\Office\16.0\Common"
        Const valueName As String = "UI Theme"

        Try
            Using key = Registry.CurrentUser.OpenSubKey(registryPath)
                If key Is Nothing Then Return OfficeTheme.Unknown

                Dim raw = key.GetValue(valueName)
                If raw Is Nothing Then Return OfficeTheme.Unknown

                Dim value As Integer
                If Integer.TryParse(raw.ToString(), value) Then
                    Select Case value
                        Case 0 ' Colorful
                            Return OfficeTheme.Light
                        Case 1, 2 ' Dark Gray, Black
                            Return OfficeTheme.Dark
                        Case 3 ' White
                            Return OfficeTheme.Light
                        Case 4 ' Use system setting -> resolve via Windows app theme
                            Return If(IsWindowsAppsLightTheme(), OfficeTheme.Light, OfficeTheme.Dark)
                    End Select
                End If
            End Using
        Catch
            ' fall through
        End Try

        Return OfficeTheme.Unknown
    End Function

    Private Function IsWindowsAppsLightTheme() As Boolean
        Const personalizePath As String = "Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        Const appsUseLightTheme As String = "AppsUseLightTheme"
        Try
            Using key = Registry.CurrentUser.OpenSubKey(personalizePath)
                If key Is Nothing Then Return True ' default to light if unknown
                Dim raw = key.GetValue(appsUseLightTheme)
                If raw Is Nothing Then Return True
                Dim v As Integer
                If Integer.TryParse(raw.ToString(), v) Then
                    Return v <> 0 ' 1=Light, 0=Dark
                End If
            End Using
        Catch
            ' default to light on error
        End Try
        Return True
    End Function

    Public Sub RI_Correct_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Word invoked")
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Correct2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Correct2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Correct_Word invoked")
        Globals.ThisAddIn.Correct()
    End Sub

    Public Sub RI_Summarize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Summarize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Summarize_Word invoked")
        Globals.ThisAddIn.Summarize()
    End Sub

    Public Sub RI_Shorten_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Shorten.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Shorten_Word invoked")
        Globals.ThisAddIn.Shorten()
    End Sub

    Public Sub RI_PrimLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Primlang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Word invoked")
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_PrimLang2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_PrimLang2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrimLang_Word invoked")
        Globals.ThisAddIn.InLanguage1()
    End Sub

    Public Sub RI_SecLang_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SecLang.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SecLang_Word invoked")
        Globals.ThisAddIn.InLanguage2()
    End Sub
    Public Sub RI_Improve_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Improve.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Improve_Word invoked")
        Globals.ThisAddIn.Improve()
    End Sub

    Public Sub RI_FreestyleNM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleNM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleNM_Word invoked")
        Globals.ThisAddIn.FreeStyleNM()
    End Sub

    Public Sub RI_Anonymize_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Anonymize.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Anonymize_Word invoked")
        Globals.ThisAddIn.Anonymize()
    End Sub

    Public Sub RI_Chat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Chat_Word invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_Chat2_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Chat2.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Chat_Word invoked")
        Globals.ThisAddIn.ShowChatForm()
    End Sub

    Public Sub RI_TimeSpan_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_TimeSpan.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "TimeSpan_Word invoked")
        Globals.ThisAddIn.CalculateUserMarkupTimeSpan()
    End Sub

    Public Sub RI_AcceptFormat_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_AcceptFormat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AcceptFormat_Word invoked")
        Globals.ThisAddIn.AcceptFormatting()
    End Sub

    Private Sub RI_Translate_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Translate.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Translate_Word invoked")
        Globals.ThisAddIn.InOther()
    End Sub

    Private Sub Settings_Click(sender As Object, e As RibbonControlEventArgs) 'Handles Settings.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Settings_Word invoked")
        Globals.ThisAddIn.ShowSettings()
    End Sub

    Private Sub RI_FreestyleAM_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_FreestyleAM.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleAM_Word invoked")
        Globals.ThisAddIn.FreeStyleAM()
    End Sub

    Private Sub RI_SwitchParty_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_SwitchParty.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SwitchParty_Word invoked")
        Globals.ThisAddIn.SwitchParty()
    End Sub

    Private Sub RI_Regex_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Regex.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Regex_Word invoked")
        Globals.ThisAddIn.RegexSearchReplace()
    End Sub

    Private Sub RI_Import_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Import_Word invoked")
        Globals.ThisAddIn.ImportTextFile()
    End Sub

    Private Sub RI_Halves_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Halves.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Halves_Word invoked")
        Globals.ThisAddIn.CompareSelectionHalves()
    End Sub

    Private Sub RI_Search_Click(sender As Object, e As RibbonControlEventArgs) 'Handles RI_Import.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Search_Word invoked")
        Globals.ThisAddIn.ContextSearch()
    End Sub

    Private Sub Easteregg_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Easteregg_Word invoked")
        Globals.ThisAddIn.EasterEgg()
    End Sub

    Private Sub RI_Transcriptor_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Transcriptor_Word invoked")
        Globals.ThisAddIn.Transcriptor()
    End Sub

    Private Sub RI_Explain_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Explain_Word invoked")
        Globals.ThisAddIn.Explain()
    End Sub

    Private Sub RI_SuggestTitles_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SuggestTitles_Word invoked")
        Globals.ThisAddIn.SuggestTitles()
    End Sub

    Private Sub RI_CreatePodcast_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CreatePodcast_Word invoked")
        Globals.ThisAddIn.CreatePodcast()
    End Sub

    Private Sub RI_CreateAudio_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CreateAudio_Word invoked")
        Globals.ThisAddIn.CreateAudio()
    End Sub

    Private Sub RI_NoFillers_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "NoFillers_Word invoked")
        Globals.ThisAddIn.NoFillers()
    End Sub

    Private Sub RI_Friendly_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Friendly_Word invoked")
        Globals.ThisAddIn.Friendly()
    End Sub

    Private Sub RI_Convincing_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Convincing_Word invoked")
        Globals.ThisAddIn.Convincing()
    End Sub

    Private Sub RI_SpecialModel_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SpecialModel_Word invoked")
        Globals.ThisAddIn.SpecialModel()
    End Sub

    Private Sub RI_Anonymization_Click(sender As Object, e As RibbonControlEventArgs)
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Anonymization_Word invoked")
        Globals.ThisAddIn.AnonymizeSelection()
    End Sub

    Private Sub RI_InsertClipboard_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_InsertClipboard.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "InsertClipboard_Word invoked")
        Globals.ThisAddIn.InsertClipboard()
    End Sub

    Private Sub RI_BallooMergePart_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePart.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BallooMergePart_Word invoked")
        Globals.ThisAddIn.BalloonMerge(False, True)
    End Sub

    Private Sub RI_BalloonMergeFull_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFull.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BalloonMergeFull_Word invoked")
        Globals.ThisAddIn.BalloonMerge(True, True)
    End Sub

    Private Sub RI_BalloonMergePartPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePartPrompt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BalloonMergePartPrompt_Word invoked")
        Globals.ThisAddIn.BalloonMerge(False, False)
    End Sub

    Private Sub RI_BalloonMergeFullPrompt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFullPrompt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BalloonMergeFullPrompt_Word invoked")
        Globals.ThisAddIn.BalloonMerge(True, False)
    End Sub

    Private Sub RI_FreestyleRepeat_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FreestyleRepeat.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FreestyleRepeat_Word invoked")
        Globals.ThisAddIn.FreeStyleRepeat()
    End Sub

    Private Sub RI_ApplyMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ApplyMyStyle_Word invoked")
        Globals.ThisAddIn.ApplyMyStyle()
    End Sub

    Private Sub RI_DefineMyStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DefineMyStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "DefineMyStyle_Word invoked")
        Globals.ThisAddIn.DefineMyStyle()
    End Sub

    Private Sub RI_DocCheck_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DocCheck.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "DocCheck_Word invoked")
        Globals.ThisAddIn.RunDocCheck()
    End Sub

    Private Sub RI_FindClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FindClause.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FindClause_Word invoked")
        Globals.ThisAddIn.FindClause()
    End Sub

    Private Sub RI_AddClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_AddClause.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "AddClause_Word invoked")
        Globals.ThisAddIn.AddClause()
    End Sub

    Private Sub RI_WebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_WebAgent.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "WebAgent_Word invoked")
        Globals.ThisAddIn.WebAgent()
    End Sub

    Private Sub RI_EditWebAgent_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_EditWebAgent.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "EditWebAgent_Word invoked")
        Globals.ThisAddIn.CreateModifyWebAgentScript()
    End Sub

    Private Sub RI_Markdown_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Markdown.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Markdown_Word invoked")
        Globals.ThisAddIn.ConvertMarkdownToWord()
    End Sub

    Private Sub RI_FindHidden_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FindHidden.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FindHidden_Word invoked")
        Globals.ThisAddIn.FindHiddenPrompts()
    End Sub

    Private Sub RI_ContentControls_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ContentControls.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ContentControls_Word invoked")
        Globals.ThisAddIn.RemoveContentControlsRespectSelection()
    End Sub

    Private Sub RI_HelpMe_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_HelpMe.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "HelpMe_Word invoked")
        Globals.ThisAddIn.HelpMeInky()
    End Sub

    Private Sub RI_PrepareRedactions_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_PrepareRedactions.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "PrepareRedactions_Word invoked")
        Globals.ThisAddIn.PrepareRedactedPDF()
    End Sub

    Private Sub RI_FinalizeRedactions_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FinalizeRedactions.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FinalizeRedactions_Word invoked")
        Globals.ThisAddIn.FlattenRedactedPDF()
    End Sub

    Private Sub RI_CheckDocumentsII_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_CheckDocumentsII.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "CheckDocumentsII_Word invoked")
        Globals.ThisAddIn.CheckDocumentII()
    End Sub

    Private Sub RI_EditRedact_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_EditRedact.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "EditRedact_Word invoked")
        Globals.ThisAddIn.EditRedactionInstructions()
    End Sub

    Private Sub RI_Filibuster_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Filibuster.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Filibuster_Word invoked")
        Globals.ThisAddIn.Filibuster()
    End Sub

    Private Sub RI_ArgueAgainst_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ArgueAgainst.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ArgueAgainst_Word invoked")
        Globals.ThisAddIn.ArgueAgainst()
    End Sub

    Private Sub RI_LiveCompare_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_LiveCompare.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LiveCompare_Word invoked")
        Globals.ThisAddIn.CompareActiveDocWithOtherOpenDoc()
    End Sub

    Private Sub RI_RevisionSummary_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_RevisionsSummary.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "RevisionSummary_Word invoked")
        Globals.ThisAddIn.SummarizeDocumentChanges()
    End Sub

    Private Sub RI_DiscussInky_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_DiscussInky.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "DiscussInky_Word invoked")
        Globals.ThisAddIn.DiscussInky()
    End Sub

    Private Sub RI_LearnDocStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_LearnDocStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "LearnDocStyle_Word invoked")
        Globals.ThisAddIn.ExtractParagraphStylesToJson()
    End Sub

    Private Sub RI_ApplyDocStyle_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ApplyDocStyle.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ApplyDocStyle_Word invoked")
        Globals.ThisAddIn.ApplyStyleTemplate()
    End Sub

    Private Sub RI_ConvertDocToTxt_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_ConvertDocToTxt.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "ConvertDocToTxt_Word invoked")
        Globals.ThisAddIn.ExportFileContentToText()
    End Sub

    Private Sub RI_FlattenPDF_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_FlattenPDF.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "FlattenPDF_Word invoked")
        Globals.ThisAddIn.FlattenPdfToImages()
    End Sub

    Private Sub RI_Charting_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Charting.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Charting_Word invoked")
        Globals.ThisAddIn.OpenExistingDrawioFileForEditing()
    End Sub

    Private Sub RI_Snapshot_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Snapshot.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Snapshot_Word invoked")
        Globals.ThisAddIn.SelectSnapshotDocument()
    End Sub

    Private Sub RI_Remove_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Remove.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Remove_Word invoked")
        Globals.ThisAddIn.RemoveRIPrefixFromComments()
    End Sub

    Private Sub RI_WebApp_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_WebApp.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "WebApp_Word invoked")
        Globals.ThisAddIn.ConvertDrawioToHtml()
    End Sub

    Private Sub RI_SplitPDF_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_SplitPDF.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "SplitPDF_Word invoked")
        Globals.ThisAddIn.SplitPdfByExhibits()
    End Sub

    Private Sub RI_StoreOriginalClause_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_StoreOriginalClause.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "StoreOriginalClause_Word invoked")
        Globals.ThisAddIn.StoreOriginalClause()
    End Sub

    Private Sub RI_JustifyMarkup_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_JustifyMarkup.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "JustifyMarkup_Word invoked")
        Globals.ThisAddIn.JustifyMarkup()
    End Sub

    Private Sub RI_BalloonMergePartJustify_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergePartJustify.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BalloonMergePartJustify_Word invoked")
        Globals.ThisAddIn.BalloonMergeWithJustification(False, True)
    End Sub

    Private Sub RI_BalloonMergeFullJustify_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_BalloonMergeFullJustify.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "BalloonMergeFullJustify_Word invoked")
        Globals.ThisAddIn.BalloonMergeWithJustification(True, True)
    End Sub

    Private Sub RI_Stamper_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Stamper.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Stamper_Word invoked")
        Globals.ThisAddIn.StampExhibitPDF()
    End Sub

    Private Sub RI_Image_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Image.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "GenerateImage_Word invoked")
        Globals.ThisAddIn.GenerateImage()
    End Sub

    Private Sub RI_Tabular_Click(sender As Object, e As RibbonControlEventArgs) Handles RI_Tabular.Click
        SharedLogger.Log(ThisAddIn._context, ThisAddIn._context.RDV, "Tabular_Word invoked")
        Globals.ThisAddIn.TabularOverview()
    End Sub
End Class