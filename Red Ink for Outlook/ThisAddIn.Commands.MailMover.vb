' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Commands.MailMover.vb
' Purpose: AI-assisted mail sorting. Scans Inbox, Sent Items, or selected mails,
'          asks an LLM to assign each to a target subfolder based on user-defined
'          rules, presents results in a review dialog, and moves confirmed mails.
'
' Architecture:
'  - Rule files: Loaded from INI_MailMoverPath (central) and INI_MailMoverPathLocal
'    (local). Files may contain segments delimited by [SegmentTitle]. Comment
'    lines starting with ";" are ignored.
'  - Folder enumeration: Recursively collects all MAPI folders from the mailbox.
'  - Mail extraction: Metadata (sender, recipients, date, subject, read/replied/
'    forwarded status, attachment names) plus optionally truncated body text.
'  - Batched LLM calls: Mails are sent in configurable batches; the LLM returns
'    a JSON array of assignments with folder path and confidence.
'  - Review dialog: Custom DPI-aware WinForms dialog styled consistently with
'    ShowCustomYesNoBox. Each row shows mail summary with a ComboBox pre-set to
'    the AI suggestion. Unsure items have no pre-selection. Rows are color-coded
'    by confidence.
'  - Undo support: After moving, original folder paths are persisted so the
'    operation can be reversed.
'  - Threading: LLM calls are async; all Outlook COM access is marshalled to
'    the UI thread via SwitchToUi / mainThreadControl.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "MailMover Constants"

    ''' <summary>Maximum number of mails to scan from a folder (guard).</summary>
    Private Const MailMover_MaxMails As Integer = 500

    ''' <summary>Number of mails sent to the LLM in a single batch.</summary>
    Private Const MailMover_BatchSize As Integer = 5

    ''' <summary>Hard cap for full body mode (characters).</summary>
    Private Const MailMover_BodyMaxFull As Integer = 10000

    ''' <summary>Threshold above which a dry-run confirmation is shown.</summary>
    Private Const MailMover_DryRunThreshold As Integer = 20

    ''' <summary>Maximum subject length shown in the review dialog.</summary>
    Private Const MailMover_SubjectDisplayLen As Integer = 60

    ''' <summary>MAPI property for last verb executed (Reply=102, ReplyAll=103, Forward=104).</summary>
    Private Const PR_LAST_VERB_EXECUTED As String = "http://schemas.microsoft.com/mapi/proptag/0x10810003"

    ''' <summary>Sample content for a newly created local rule file.</summary>
    Private Const MailMover_SampleRules As String =
        "; Mail Mover Rules" & vbCrLf &
        "; Lines starting with ; are comments and will be ignored." & vbCrLf &
        "; You can define segments using [SegmentTitle] headers." & vbCrLf &
        "; Within each segment, write natural-language rules that describe" & vbCrLf &
        "; how mails should be sorted into folders." & vbCrLf &
        ";" & vbCrLf &
        "; Example:" & vbCrLf &
        "; [Client Sorting]" & vbCrLf &
        "; Mails from @clientA.com or mentioning ""Project Alpha"" -> Inbox\Clients\Alpha" & vbCrLf &
        "; Mails from john.doe@example.com -> Inbox\Internal\John" & vbCrLf &
        "; Newsletter or marketing mails -> Inbox\Newsletters" & vbCrLf &
        ";" & vbCrLf &
        "; [Personal]" & vbCrLf &
        "; Mails from family members or personal contacts -> Inbox\Personal" & vbCrLf &
        ""

#End Region

#Region "MailMover Data Classes"

    ''' <summary>Represents a single mail being assessed for folder assignment.</summary>
    Private Class MailMoverEntry
        Public Property EntryID As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Recipients As String
        Public Property ReceivedTime As DateTime
        Public Property IsRead As Boolean
        Public Property StatusFlag As String  ' "Replied", "Forwarded", "RepliedAll", ""
        Public Property AttachmentNames As String
        Public Property BodyExcerpt As String
        Public Property SourceFolderPath As String
        Public Property SuggestedFolder As String   ' LLM suggestion (Nothing = unsure)
        Public Property Confidence As String        ' "high", "medium", "low", ""
        Public Property SelectedFolder As String    ' User's final choice (Nothing = skip)
        Public Property IsSent As Boolean
    End Class

    ''' <summary>Represents a rule segment parsed from a rule file.</summary>
    Private Class MailMoverRuleSegment
        Public Property Title As String
        Public Property IsLocal As Boolean
        Public Property Rules As String
    End Class

    ''' <summary>Stores undo information for a single moved mail.</summary>
    Private Class MailMoverUndoEntry
        Public Property EntryID As String
        Public Property OriginalFolderPath As String
        Public Property TargetFolderPath As String
    End Class

    ''' <summary>Persisted undo state from the last MailMover operation.</summary>
    Private _mailMoverUndoList As List(Of MailMoverUndoEntry) = Nothing

#End Region

#Region "MailMover Entry Point"

    ''' <summary>
    ''' Main entry point for the MailMover feature.
    ''' </summary>
    Public Async Sub MailMover()

        Try
            ' 1. Load and select rules
            Dim selectedRules As String = SelectMailMoverRules()
            If selectedRules Is Nothing Then Return

            ' 2. Determine mail source
            Dim mails As List(Of MailMoverEntry) = CollectMailsForProcessing()
            If mails Is Nothing OrElse mails.Count = 0 Then Return

            ' 3. Ask body mode: latest reply only or full body (capped)
            Dim useLatestOnly As Boolean = True
            Dim bodyChoice As Integer = ShowCustomYesNoBox(
                "How much of the email body should be sent to the AI?" & vbCrLf & vbCrLf &
                $"{ChrW(&H2022)} Latest reply: Only the most recent part of each mail (no previous chain). Faster and cheaper." & vbCrLf &
                $"{ChrW(&H2022)} Full body: The entire mail body (capped at {MailMover_BodyMaxFull:N0} chars). Slower, but more context for long threads.",
                "Latest reply (recommended)", "Full body",
                $"{AN} - Mail Mover")
            If bodyChoice = 0 Then Return
            useLatestOnly = (bodyChoice = 1)

            ' 4. Dry-run confirmation for large sets
            If mails.Count > MailMover_DryRunThreshold Then
                Dim confirmChoice As Integer = ShowCustomYesNoBox(
                    $"Will scan {mails.Count} mails using the selected rule set. This requires " &
                    $"{Math.Ceiling(mails.Count / CDbl(MailMover_BatchSize))} AI calls. Continue?",
                    "Yes, continue", "No, abort",
                    $"{AN} - Mail Mover")
                If confirmChoice <> 1 Then Return
            End If

            ' 5. Collect folder list
            Dim folders As List(Of String) = CollectAllFolders()
            If folders Is Nothing OrElse folders.Count = 0 Then
                ShowCustomMessageBox("No mailbox folders found.", $"{AN} - Mail Mover")
                Return
            End If

            ' 6. Process mails via LLM in batches
            Dim success As Boolean = Await ProcessMailBatchesAsync(mails, selectedRules, folders, useLatestOnly)
            If Not success Then Return

            ' 7. Show review dialog
            Dim approved As Boolean = ShowMailMoverReviewDialog(mails, folders)
            If Not approved Then Return

            ' 8. Move mails
            MoveApprovedMails(mails)

        Catch ex As System.Exception
            ShowCustomMessageBox($"MailMover error: {ex.Message}", $"{AN} - Mail Mover")
        End Try
    End Sub

    ''' <summary>
    ''' Undoes the last MailMover operation by moving mails back to their original folders.
    ''' </summary>
    Public Sub MailMoverUndo()
        Try
            If _mailMoverUndoList Is Nothing OrElse _mailMoverUndoList.Count = 0 Then
                ShowCustomMessageBox("No MailMover operation to undo.", $"{AN} - Mail Mover")
                Return
            End If

            Dim confirmChoice As Integer = ShowCustomYesNoBox(
                $"Undo the last MailMover operation? This will move {_mailMoverUndoList.Count} mail(s) back to their original folders.",
                "Yes, undo", "No, keep",
                $"{AN} - Mail Mover Undo")
            If confirmChoice <> 1 Then Return

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim movedCount As Integer = 0
            Dim failedCount As Integer = 0

            For Each entry In _mailMoverUndoList
                Try
                    Dim entryId As String = entry.EntryID
                    Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
                    If mi Is Nothing Then
                        failedCount += 1
                        Continue For
                    End If

                    Dim targetFolder As MAPIFolder = FindFolderByPath(ns, entry.OriginalFolderPath)
                    If targetFolder Is Nothing Then
                        failedCount += 1
                        Continue For
                    End If

                    mi.Move(targetFolder)
                    movedCount += 1
                Catch
                    failedCount += 1
                End Try
            Next

            _mailMoverUndoList = Nothing
            ShowCustomMessageBox(
                $"Undo complete: {movedCount} mail(s) moved back." &
                If(failedCount > 0, $"{vbCrLf}{failedCount} mail(s) could not be restored.", ""),
                $"{AN} - Mail Mover Undo")

        Catch ex As System.Exception
            ShowCustomMessageBox($"Undo error: {ex.Message}", $"{AN} - Mail Mover")
        End Try
    End Sub

#End Region

#Region "MailMover Rule Loading"

    ''' <summary>
    ''' Loads rule files, parses segments, presents selection to user, returns chosen rules text.
    ''' Returns Nothing if cancelled or no rules available.
    ''' </summary>
    Private Function SelectMailMoverRules() As String
        Dim centralPath As String = Environment.ExpandEnvironmentVariables(If(INI_MailMoverPath, ""))
        Dim localPath As String = Environment.ExpandEnvironmentVariables(If(INI_MailMoverPathLocal, ""))

        Dim segments As New List(Of MailMoverRuleSegment)()

        ' Parse central file
        If Not String.IsNullOrWhiteSpace(centralPath) AndAlso File.Exists(centralPath) Then
            segments.AddRange(ParseRuleFile(centralPath, isLocal:=False))
        End If

        ' Parse local file
        If Not String.IsNullOrWhiteSpace(localPath) AndAlso File.Exists(localPath) Then
            segments.AddRange(ParseRuleFile(localPath, isLocal:=True))
        End If

        ' Build selection items
        Dim items As New List(Of SelectionItem)()
        Dim segMap As New Dictionary(Of Integer, MailMoverRuleSegment)()
        Dim idx As Integer = 1

        For Each seg In segments
            Dim displayName As String = seg.Title & If(seg.IsLocal, " (local)", "")
            items.Add(New SelectionItem(displayName, idx))
            segMap(idx) = seg
            idx += 1
        Next

        ' Add "Edit local rules..." option if local path is defined
        Const EditLocalValue As Integer = -1
        If Not String.IsNullOrWhiteSpace(localPath) Then
            items.Add(New SelectionItem("Edit local rules...", EditLocalValue))
        End If

        ' Add "Undo last Mail Mover" option if undo is available
        Const UndoValue As Integer = -2
        If _mailMoverUndoList IsNot Nothing AndAlso _mailMoverUndoList.Count > 0 Then
            items.Add(New SelectionItem("Undo last Mail Mover operation", UndoValue))
        End If

        If items.Count = 0 Then
            ShowCustomMessageBox("No Mail Mover rules configured. Please define 'MailMoverPath' or 'MailMoverPathLocal' in the configuration.", $"{AN} - Mail Mover")
            Return Nothing
        End If

        Dim result As Integer = SelectValue(items, If(segments.Count > 0, 1, EditLocalValue),
                                             "Select the rule set to use for mail sorting:",
                                             $"{AN} - Mail Mover - Rule Selection")

        If result = 0 Then Return Nothing

        If result = EditLocalValue Then
            ' Create or open local rules file
            If Not File.Exists(localPath) Then
                Try
                    Dim dir As String = Path.GetDirectoryName(localPath)
                    If Not String.IsNullOrEmpty(dir) Then Directory.CreateDirectory(dir)
                    File.WriteAllText(localPath, MailMover_SampleRules, Encoding.UTF8)
                Catch ex As System.Exception
                    ShowCustomMessageBox($"Could not create rule file: {ex.Message}", $"{AN} - Mail Mover")
                    Return Nothing
                End Try
            End If
            ShowTextFileEditor(localPath, $"{AN} - Mail Mover - Edit Local Rules", False, _context)
            ' After editing, re-run selection
            Return SelectMailMoverRules()
        End If

        If result = UndoValue Then
            MailMoverUndo()
            Return Nothing
        End If

        If segMap.ContainsKey(result) Then
            Return segMap(result).Rules
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Parses a rule file into segments. If no [segment] headers exist, returns a single
    ''' "Standard rules" segment. Comment lines (starting with ;) are stripped.
    ''' </summary>
    Private Function ParseRuleFile(filePath As String, isLocal As Boolean) As List(Of MailMoverRuleSegment)
        Dim segments As New List(Of MailMoverRuleSegment)()
        Try
            Dim lines As String() = File.ReadAllLines(filePath, Encoding.UTF8)
            Dim currentTitle As String = Nothing
            Dim currentRules As New StringBuilder()

            For Each rawLine In lines
                Dim line As String = rawLine.TrimStart()

                ' Skip comments
                If line.StartsWith(";") Then Continue For

                ' Detect segment header
                Dim headerMatch As Match = Regex.Match(line, "^\[(.+)\]\s*$")
                If headerMatch.Success Then
                    ' Save previous segment
                    If currentTitle IsNot Nothing Then
                        Dim ruleText As String = currentRules.ToString().Trim()
                        If ruleText.Length > 0 Then
                            segments.Add(New MailMoverRuleSegment With {
                                .Title = currentTitle,
                                .IsLocal = isLocal,
                                .Rules = ruleText
                            })
                        End If
                    End If
                    currentTitle = headerMatch.Groups(1).Value.Trim()
                    currentRules.Clear()
                Else
                    currentRules.AppendLine(rawLine)
                End If
            Next

            ' Save last segment
            If currentTitle IsNot Nothing Then
                Dim ruleText As String = currentRules.ToString().Trim()
                If ruleText.Length > 0 Then
                    segments.Add(New MailMoverRuleSegment With {
                        .Title = currentTitle,
                        .IsLocal = isLocal,
                        .Rules = ruleText
                    })
                End If
            ElseIf currentRules.Length > 0 Then
                ' No segments found - treat entire file as one segment
                Dim ruleText As String = currentRules.ToString().Trim()
                If ruleText.Length > 0 Then
                    segments.Add(New MailMoverRuleSegment With {
                        .Title = "Standard rules",
                        .IsLocal = isLocal,
                        .Rules = ruleText
                    })
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"ParseRuleFile error: {ex.Message}")
        End Try
        Return segments
    End Function

#End Region

#Region "MailMover Mail Collection"

    ''' <summary>
    ''' Determines mail source (selected mails or current folder) and collects MailMoverEntry items.
    ''' Returns Nothing if cancelled.
    ''' </summary>
    Private Function CollectMailsForProcessing() As List(Of MailMoverEntry)
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim explorer As Outlook.Explorer = ComRetry(Of Outlook.Explorer)(Function() outlookApp.ActiveExplorer())
        If explorer Is Nothing Then
            ShowCustomMessageBox("No active Outlook window found.", $"{AN} - Mail Mover")
            Return Nothing
        End If

        Dim sel As Outlook.Selection = ComRetry(Of Outlook.Selection)(Function() explorer.Selection)
        Dim selCount As Integer = 0
        If sel IsNot Nothing Then selCount = ComRetry(Of Integer)(Function() sel.Count)

        ' If mails are selected, use them
        If selCount > 1 Then
            Dim useSelection As Integer = ShowCustomYesNoBox(
                $"You have {selCount} mail(s) selected. Use the selection or scan the current folder?",
                $"Use selection ({selCount} mails)", "Scan current folder",
                $"{AN} - Mail Mover")
            If useSelection = 0 Then Return Nothing

            If useSelection = 1 Then
                Return ExtractMailsFromSelection(sel)
            End If
        End If

        ' Use the folder the user is currently viewing
        Dim currentFolder As MAPIFolder = Nothing
        Try
            currentFolder = ComRetry(Of MAPIFolder)(Function() explorer.CurrentFolder)
        Catch
        End Try

        If currentFolder Is Nothing Then
            ShowCustomMessageBox("Could not determine the current folder.", $"{AN} - Mail Mover")
            Return Nothing
        End If

        Dim folderPath As String = ""
        Try : folderPath = currentFolder.FolderPath : Catch : End Try

        ' Determine if this is a Sent Items folder (heuristic: check default folder identity)
        Dim isSent As Boolean = False
        Try
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim sentFolder As MAPIFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail))
            If sentFolder IsNot Nothing Then
                isSent = String.Equals(currentFolder.EntryID, sentFolder.EntryID, StringComparison.OrdinalIgnoreCase)
            End If
        Catch
        End Try

        Dim confirmChoice As Integer = ShowCustomYesNoBox(
            $"Scan the current folder for sorting?" & vbCrLf & vbCrLf &
            $"Folder: {folderPath}",
            "Yes, scan this folder", "No, abort",
            $"{AN} - Mail Mover")
        If confirmChoice <> 1 Then Return Nothing

        Return ExtractMailsFromFolder(currentFolder, isSent)
    End Function

    Private Function ExtractMailsFromSelection(sel As Outlook.Selection) As List(Of MailMoverEntry)
        Dim entries As New List(Of MailMoverEntry)()
        Dim count As Integer = ComRetry(Of Integer)(Function() sel.Count)

        For i As Integer = 1 To Math.Min(count, MailMover_MaxMails)
            Try
                Dim idx As Integer = i
                Dim item As Object = ComRetry(Function() sel.Item(idx))
                Dim mi As MailItem = TryCast(item, MailItem)
                If mi IsNot Nothing Then
                    entries.Add(BuildMailEntry(mi, isSent:=False))
                End If
            Catch
            End Try
        Next
        Return entries
    End Function

    Private Function ExtractMailsFromFolder(folder As MAPIFolder, isSent As Boolean) As List(Of MailMoverEntry)
        Dim entries As New List(Of MailMoverEntry)()
        Dim folderItems As Outlook.Items = ComRetry(Of Outlook.Items)(Function() folder.Items)
        Dim totalCount As Integer = ComRetry(Of Integer)(Function() folderItems.Count)
        Dim scanCount As Integer = Math.Min(totalCount, MailMover_MaxMails)

        folderItems.Sort("[ReceivedTime]", True)

        For i As Integer = 1 To scanCount
            Try
                Dim idx As Integer = i
                Dim item As Object = ComRetry(Function() folderItems.Item(idx))
                Dim mi As MailItem = TryCast(item, MailItem)
                If mi IsNot Nothing Then
                    entries.Add(BuildMailEntry(mi, isSent))
                End If
            Catch
            End Try
        Next
        Return entries
    End Function

    Private Function BuildMailEntry(mi As MailItem, isSent As Boolean) As MailMoverEntry
        Dim entry As New MailMoverEntry()

        entry.EntryID = ComRetry(Of String)(Function() mi.EntryID)
        entry.Subject = ComRetry(Of String)(Function() If(mi.Subject, "(no subject)"))
        entry.SenderName = ComRetry(Of String)(Function() If(mi.SenderName, ""))
        entry.SenderEmail = ""
        Try : entry.SenderEmail = ComRetry(Of String)(Function() If(mi.SenderEmailAddress, "")) : Catch : End Try
        entry.Recipients = ""
        Try : entry.Recipients = ComRetry(Of String)(Function() CStr(mi.To)) : Catch : End Try
        entry.ReceivedTime = ComRetry(Of Date)(Function() mi.ReceivedTime)
        entry.IsRead = ComRetry(Of Boolean)(Function() Not mi.UnRead)
        entry.IsSent = isSent

        entry.StatusFlag = ""
        Try
            Dim lastVerb As Integer = CInt(mi.PropertyAccessor.GetProperty(PR_LAST_VERB_EXECUTED))
            Select Case lastVerb
                Case 102 : entry.StatusFlag = "Replied"
                Case 103 : entry.StatusFlag = "RepliedAll"
                Case 104 : entry.StatusFlag = "Forwarded"
            End Select
        Catch
        End Try

        entry.AttachmentNames = ""
        Try
            Dim attachCount As Integer = ComRetry(Of Integer)(Function() mi.Attachments.Count)
            If attachCount > 0 Then
                Dim names As New List(Of String)()
                For a As Integer = 1 To attachCount
                    Try
                        Dim aidx As Integer = a
                        names.Add(ComRetry(Of String)(Function() mi.Attachments.Item(aidx).FileName))
                    Catch
                    End Try
                Next
                entry.AttachmentNames = String.Join(", ", names)
            End If
        Catch
        End Try

        entry.BodyExcerpt = GetMailBody(mi)

        entry.SourceFolderPath = ""
        Try
            Dim parent As MAPIFolder = TryCast(mi.Parent, MAPIFolder)
            If parent IsNot Nothing Then
                entry.SourceFolderPath = parent.FolderPath
            End If
        Catch
        End Try

        Return entry
    End Function

#End Region

#Region "MailMover Folder Enumeration"

    Private Function CollectAllFolders() As List(Of String)
        Dim folders As New List(Of String)()
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim defaultStore As Outlook.Store = ns.DefaultStore
            Dim rootFolder As MAPIFolder = defaultStore.GetRootFolder()
            EnumerateFolders(rootFolder, folders)
        Catch ex As System.Exception
            Debug.WriteLine($"CollectAllFolders error: {ex.Message}")
        End Try
        Return folders
    End Function

    Private Sub EnumerateFolders(parentFolder As MAPIFolder, folders As List(Of String))
        Try
            folders.Add(parentFolder.FolderPath)
            Dim subFolders As Outlook.Folders = parentFolder.Folders
            For Each subFolder As MAPIFolder In subFolders
                EnumerateFolders(subFolder, folders)
            Next
        Catch
        End Try
    End Sub

    Private Function FindFolderByPath(ns As Outlook.NameSpace, folderPath As String) As MAPIFolder
        If String.IsNullOrWhiteSpace(folderPath) Then Return Nothing
        Try
            Dim parts As String() = folderPath.Split(New Char() {"\"c}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length < 1 Then Return Nothing

            Dim defaultStore As Outlook.Store = ns.DefaultStore
            Dim current As MAPIFolder = defaultStore.GetRootFolder()

            For i As Integer = 1 To parts.Length - 1
                Dim partName As String = parts(i)
                Dim found As Boolean = False
                Dim subFolders As Outlook.Folders = current.Folders
                For Each sub_ As MAPIFolder In subFolders
                    If String.Equals(sub_.Name, partName, StringComparison.OrdinalIgnoreCase) Then
                        current = sub_
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then Return Nothing
            Next

            Return current
        Catch
            Return Nothing
        End Try
    End Function

#End Region

#Region "MailMover LLM Processing"

    Private Async Function ProcessMailBatchesAsync(
        mails As List(Of MailMoverEntry),
        rules As String,
        folders As List(Of String),
        useLatestOnly As Boolean) As Task(Of Boolean)

        Dim folderListText As String = String.Join(vbCrLf, folders)
        Dim totalBatches As Integer = CInt(Math.Ceiling(mails.Count / CDbl(MailMover_BatchSize)))

        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = totalBatches
        ProgressBarModule.GlobalProgressLabel = "Analyzing mails..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread($"{AN} - Mail Mover", "Preparing...")

        Try
            Dim batchIndex As Integer = 0

            While batchIndex < mails.Count
                If ProgressBarModule.CancelOperation Then
                    Return False
                End If

                Dim batchStart As Integer = batchIndex
                Dim batchEnd As Integer = Math.Min(batchIndex + MailMover_BatchSize - 1, mails.Count - 1)
                Dim currentBatch As Integer = CInt(Math.Floor(batchIndex / CDbl(MailMover_BatchSize))) + 1

                ProgressBarModule.GlobalProgressValue = currentBatch - 1
                ProgressBarModule.GlobalProgressLabel = $"Analyzing batch {currentBatch}/{totalBatches} ({batchEnd - batchStart + 1} mails)..."

                Dim userPrompt As New StringBuilder()
                userPrompt.AppendLine("<RULES>")
                userPrompt.AppendLine(rules)
                userPrompt.AppendLine("</RULES>")
                userPrompt.AppendLine()
                userPrompt.AppendLine("<AVAILABLE_FOLDERS>")
                userPrompt.AppendLine(folderListText)
                userPrompt.AppendLine("</AVAILABLE_FOLDERS>")
                userPrompt.AppendLine()
                userPrompt.AppendLine("<EMAILS>")

                For i As Integer = batchStart To batchEnd
                    Dim mail As MailMoverEntry = mails(i)
                    Dim mailNum As Integer = i - batchStart + 1

                    userPrompt.AppendLine($"<EMAIL id=""{mailNum}"">")
                    userPrompt.AppendLine($"Subject: {mail.Subject}")
                    userPrompt.AppendLine($"From: {mail.SenderName} <{mail.SenderEmail}>")
                    userPrompt.AppendLine($"To: {mail.Recipients}")
                    userPrompt.AppendLine($"Date: {mail.ReceivedTime:yyyy-MM-dd HH:mm}")
                    userPrompt.AppendLine($"Read: {mail.IsRead}")
                    If Not String.IsNullOrEmpty(mail.StatusFlag) Then
                        userPrompt.AppendLine($"Status: {mail.StatusFlag}")
                    End If
                    userPrompt.AppendLine($"Direction: {If(mail.IsSent, "Sent", "Received")}")
                    If Not String.IsNullOrEmpty(mail.AttachmentNames) Then
                        userPrompt.AppendLine($"Attachments: {mail.AttachmentNames}")
                    End If
                    userPrompt.AppendLine($"Current folder: {mail.SourceFolderPath}")
                    userPrompt.AppendLine()

                    Dim body As String = mail.BodyExcerpt
                    If Not String.IsNullOrEmpty(body) Then
                        If useLatestOnly Then
                            ' Extract latest reply only (no truncation)
                            body = GetLatestMailBody(body)
                        Else
                            ' Full body, capped at hard maximum
                            If body.Length > MailMover_BodyMaxFull Then
                                body = body.Substring(0, MailMover_BodyMaxFull) & " [...]"
                            End If
                        End If
                    End If

                    If Not String.IsNullOrEmpty(body) Then
                        userPrompt.AppendLine("Body:")
                        userPrompt.AppendLine(body)
                    End If
                    userPrompt.AppendLine("</EMAIL>")
                    userPrompt.AppendLine()
                Next

                userPrompt.AppendLine("</EMAILS>")

                Dim response As String = Await LLM(
                    SP_MailMover, userPrompt.ToString(),
                    "", "", 0, False, True)

                If String.IsNullOrWhiteSpace(response) Then
                    batchIndex = batchEnd + 1
                    Continue While
                End If

                ParseMailMoverResponse(response, mails, batchStart, batchEnd)

                batchIndex = batchEnd + 1
            End While

            Return True

        Finally
            ProgressBarModule.CancelOperation = True
        End Try
    End Function

    Private Sub ParseMailMoverResponse(response As String, mails As List(Of MailMoverEntry), batchStart As Integer, batchEnd As Integer)
        Try
            Dim jsonText As String = response.Trim()
            If jsonText.StartsWith("```") Then
                Dim startIdx As Integer = jsonText.IndexOf("[")
                Dim endIdx As Integer = jsonText.LastIndexOf("]")
                If startIdx >= 0 AndAlso endIdx > startIdx Then
                    jsonText = jsonText.Substring(startIdx, endIdx - startIdx + 1)
                End If
            End If

            Debug.WriteLine($"[MailMover] Parsing JSON for batch {batchStart}-{batchEnd}")
            Debug.WriteLine($"[MailMover] Raw JSON (first 500 chars): {jsonText.Substring(0, Math.Min(500, jsonText.Length))}")

            Dim arr As JArray = JArray.Parse(jsonText)

            For Each item As JObject In arr
                Dim id As Integer = 0
                If item("id") IsNot Nothing Then id = CInt(item("id"))
                Dim folder As String = Nothing
                If item("folder") IsNot Nothing AndAlso item("folder").Type <> JTokenType.Null Then
                    folder = CStr(item("folder")).Trim()
                End If
                Dim confidence As String = "low"
                If item("confidence") IsNot Nothing Then confidence = CStr(item("confidence")).ToLowerInvariant()

                Debug.WriteLine($"[MailMover] id={id}, confidence={confidence}, folder=[{If(folder, "NULL")}]")

                Dim absoluteIndex As Integer = batchStart + id - 1
                If absoluteIndex < batchStart OrElse absoluteIndex > batchEnd OrElse absoluteIndex >= mails.Count Then
                    Debug.WriteLine($"[MailMover]   -> SKIPPED: absoluteIndex={absoluteIndex} out of range [{batchStart}..{batchEnd}]")
                    Continue For
                End If

                If confidence = "low" OrElse String.IsNullOrWhiteSpace(folder) Then
                    mails(absoluteIndex).SuggestedFolder = Nothing
                    mails(absoluteIndex).Confidence = confidence
                    Debug.WriteLine($"[MailMover]   -> Set to Nothing (low/null)")
                Else
                    mails(absoluteIndex).SuggestedFolder = folder
                    mails(absoluteIndex).Confidence = confidence
                    Debug.WriteLine($"[MailMover]   -> SuggestedFolder=[{folder}]")
                End If
            Next

        Catch ex As System.Exception
            Debug.WriteLine($"ParseMailMoverResponse error: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "MailMover Review Dialog"

    Private Function ShowMailMoverReviewDialog(mails As List(Of MailMoverEntry), folders As List(Of String)) As Boolean
        Dim approved As Boolean = False

        Using frm As New Form()
            frm.Text = $"{AN} - MailMover - Review & Confirm"
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.Size = New Drawing.Size(1400, 700)
            frm.MinimumSize = New Drawing.Size(1100, 400)
            frm.FormBorderStyle = FormBorderStyle.Sizable
            frm.MaximizeBox = True
            frm.MinimizeBox = False
            frm.ShowInTaskbar = True

            ' Standard font consistent with ShowCustomYesNoBox / ShowCustomMessageBox
            Dim standardFont As New Drawing.Font("Segoe UI", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
            frm.Font = standardFont

            ' Set icon using GetLogoBitmap (same as DragDropForm / other UI)
            Try
                Dim bmp As New Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                frm.Icon = Drawing.Icon.FromHandle(bmp.GetHicon())
                bmp.Dispose()
            Catch
            End Try

            ' Try to bring Outlook forward
            Try
                Dim hwnd As IntPtr = FindWindow("rctrl_renwnd32", Nothing)
                If hwnd <> IntPtr.Zero Then
                    NativeMethods.SetForegroundWindow(hwnd)
                End If
            Catch
            End Try

            ' Main layout
            Dim mainLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(12, 12, 12, 2),
                .Margin = New Padding(0),
                .ColumnCount = 1,
                .RowCount = 3
            }
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            frm.Controls.Add(mainLayout)

            ' Header label
            Dim assignedCount As Integer = mails.Where(Function(m) m.SuggestedFolder IsNot Nothing).Count()
            Dim unsureCount As Integer = mails.Count - assignedCount
            Dim headerLabel As New Label() With {
                .Text = $"Review mail assignments ({assignedCount} assigned, {unsureCount} unassigned). " &
                        "Adjust target folders as needed, then confirm. Click column headers to sort.",
                .AutoSize = True,
                .Dock = DockStyle.Fill,
                .TextAlign = Drawing.ContentAlignment.MiddleLeft,
                .Font = standardFont,
                .Margin = New Padding(0, 2, 0, 10),
                .Padding = New Padding(0, 4, 0, 4)
            }
            mainLayout.Controls.Add(headerLabel, 0, 0)

            ' DataGridView
            Dim dgv As New DataGridView() With {
                .Dock = DockStyle.Fill,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .AllowUserToResizeRows = False,
                .AllowUserToResizeColumns = True,
                .RowHeadersVisible = False,
                .ColumnHeadersVisible = True,
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                .EnableHeadersVisualStyles = True,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .MultiSelect = True,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                .BackgroundColor = Drawing.SystemColors.Window,
                .BorderStyle = BorderStyle.FixedSingle,
                .ReadOnly = False,
                .Font = standardFont,
                .Margin = New Padding(0, 0, 0, 8)
            }
            dgv.ColumnHeadersDefaultCellStyle.Padding = New Padding(0, 2, 0, 2)
            mainLayout.Controls.Add(dgv, 0, 1)

            ' Columns - with SortMode set for sortable headers
            Dim colCheck As New DataGridViewCheckBoxColumn() With {
                .Name = "colMove", .HeaderText = "Move", .Width = 60, .MinimumWidth = 60,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }
            Dim colDate As New DataGridViewTextBoxColumn() With {
                .Name = "colDate", .HeaderText = "Date", .Width = 130, .MinimumWidth = 120, .ReadOnly = True,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }
            colDate.DefaultCellStyle.Format = "yy-MM-dd HH:mm"

            Dim colFrom As New DataGridViewTextBoxColumn() With {
                .Name = "colFrom", .HeaderText = "From/To", .Width = 220, .MinimumWidth = 160, .ReadOnly = True,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }
            Dim colSubject As New DataGridViewTextBoxColumn() With {
                .Name = "colSubject", .HeaderText = "Subject", .Width = 320, .MinimumWidth = 200, .ReadOnly = True,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }
            Dim colConf As New DataGridViewTextBoxColumn() With {
                .Name = "colConf", .HeaderText = "Conf.", .Width = 60, .MinimumWidth = 50, .ReadOnly = True,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }
            colConf.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            Dim colFolder As New DataGridViewComboBoxColumn() With {
                .Name = "colFolder", .HeaderText = "Target Folder",
                .Width = 420, .MinimumWidth = 220,
                .FlatStyle = FlatStyle.Flat,
                .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton,
                .SortMode = DataGridViewColumnSortMode.Automatic,
                .Resizable = DataGridViewTriState.True
            }

            Dim folderItems As New List(Of String)()
            folderItems.Add("(skip - do not move)")
            folderItems.AddRange(folders)
            colFolder.Items.AddRange(folderItems.ToArray())

            dgv.Columns.AddRange(New DataGridViewColumn() {colCheck, colDate, colFrom, colSubject, colConf, colFolder})

            ' Auto-size Target Folder column to fill remaining space on resize
            Dim resizeFolder As New System.Action(
                Sub()
                    Dim usedWidth As Integer = 0
                    For Each col As DataGridViewColumn In dgv.Columns
                        If col.Name <> "colFolder" Then usedWidth += col.Width
                    Next
                    Dim remaining As Integer = dgv.ClientSize.Width - usedWidth - SystemInformation.VerticalScrollBarWidth - 2
                    If remaining > 220 Then
                        dgv.Columns("colFolder").Width = remaining
                    End If
                End Sub)
            AddHandler dgv.Resize, Sub(s, e) resizeFolder.Invoke()
            AddHandler dgv.ColumnWidthChanged, Sub(s, e) resizeFolder.Invoke()
            AddHandler frm.Shown, Sub(s, e) resizeFolder.Invoke()

            ' Sorting helper (handles DateTime and Boolean correctly)
            AddHandler dgv.SortCompare,
                Sub(s, e)
                    Dim val1 As Object = e.CellValue1
                    Dim val2 As Object = e.CellValue2

                    Select Case e.Column.Name
                        Case "colDate"
                            Dim d1 As DateTime = If(TypeOf val1 Is DateTime, DirectCast(val1, DateTime), DateTime.MinValue)
                            Dim d2 As DateTime = If(TypeOf val2 Is DateTime, DirectCast(val2, DateTime), DateTime.MinValue)
                            e.SortResult = DateTime.Compare(d1, d2)
                        Case "colMove"
                            Dim b1 As Boolean = False
                            Dim b2 As Boolean = False
                            If TypeOf val1 Is Boolean Then b1 = CBool(val1) Else Boolean.TryParse(CStr(val1), b1)
                            If TypeOf val2 Is Boolean Then b2 = CBool(val2) Else Boolean.TryParse(CStr(val2), b2)
                            e.SortResult = b1.CompareTo(b2)
                        Case Else
                            e.SortResult = String.Compare(CStr(If(val1, "")), CStr(If(val2, "")), StringComparison.CurrentCultureIgnoreCase)
                    End Select

                    e.Handled = True
                End Sub

            ' Populate rows
            For Each mail In mails
                Dim rowIdx As Integer = dgv.Rows.Add()
                Dim row As DataGridViewRow = dgv.Rows(rowIdx)

                Dim hasSuggestion As Boolean = Not String.IsNullOrEmpty(mail.SuggestedFolder)
                row.Cells("colMove").Value = hasSuggestion
                row.Cells("colDate").Value = mail.ReceivedTime

                Dim fromTo As String = If(mail.IsSent,
                    ChrW(&H2192) & " " & If(mail.Recipients, "").Split({";"c, ","c})(0).Trim(),
                    If(mail.SenderName, mail.SenderEmail))
                If fromTo.Length > 25 Then fromTo = fromTo.Substring(0, 22) & "..."
                row.Cells("colFrom").Value = fromTo

                Dim subj As String = If(mail.Subject, "")
                If subj.Length > MailMover_SubjectDisplayLen Then
                    subj = subj.Substring(0, MailMover_SubjectDisplayLen - 3) & "..."
                End If
                row.Cells("colSubject").Value = subj

                Dim confDisplay As String = ""
                Select Case If(mail.Confidence, "")
                    Case "high" : confDisplay = ChrW(&H25CF) & ChrW(&H25CF) & ChrW(&H25CF)
                    Case "medium" : confDisplay = ChrW(&H25CF) & ChrW(&H25CF)
                    Case "low" : confDisplay = ChrW(&H25CF)
                End Select
                row.Cells("colConf").Value = confDisplay

                Dim matchedFolder As String = Nothing
                If hasSuggestion Then
                    Dim normalizedSuggestion As String = mail.SuggestedFolder.Trim()
                    If normalizedSuggestion.StartsWith("\") AndAlso Not normalizedSuggestion.StartsWith("\\") Then
                        normalizedSuggestion = "\" & normalizedSuggestion
                    End If

                    matchedFolder = folderItems.FirstOrDefault(
                        Function(f) String.Equals(f.Trim(), normalizedSuggestion, StringComparison.OrdinalIgnoreCase))
                End If

                If matchedFolder IsNot Nothing Then
                    row.Cells("colFolder").Value = matchedFolder
                Else
                    row.Cells("colFolder").Value = "(skip - do not move)"
                    If hasSuggestion Then row.Cells("colMove").Value = False
                End If

                Select Case If(mail.Confidence, "")
                    Case "high"
                        row.DefaultCellStyle.BackColor = Drawing.Color.FromArgb(232, 245, 233)
                    Case "medium"
                        row.DefaultCellStyle.BackColor = Drawing.Color.FromArgb(255, 249, 196)
                    Case Else
                        row.DefaultCellStyle.BackColor = Drawing.Color.FromArgb(255, 235, 238)
                End Select

                row.Tag = mail
            Next

            ' Bottom button layout
            Dim btnLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .AutoSize = True,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            btnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            btnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            btnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            mainLayout.Controls.Add(btnLayout, 0, 2)

            Dim leftFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0)
            }

            Dim rightFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0)
            }

            Dim btnSpacing As Integer = 8
            Dim buttonPadding As New Padding(6, 4, 6, 4)

            Dim btnSelectAll As New Button() With {
                .Text = "Select All",
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlatStyle = FlatStyle.Flat,
                .Font = standardFont,
                .Padding = buttonPadding,
                .Margin = New Padding(0, 0, btnSpacing, 0)
            }
            btnSelectAll.FlatAppearance.BorderColor = Drawing.Color.FromArgb(180, 180, 180)

            Dim btnDeselectAll As New Button() With {
                .Text = "Deselect All",
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlatStyle = FlatStyle.Flat,
                .Font = standardFont,
                .Padding = buttonPadding,
                .Margin = New Padding(0, 0, btnSpacing, 0)
            }
            btnDeselectAll.FlatAppearance.BorderColor = Drawing.Color.FromArgb(180, 180, 180)

            Dim lblCount As New Label() With {
                .Name = "lblCount",
                .AutoSize = True,
                .Font = standardFont,
                .Margin = New Padding(0, 4, 0, 0)
            }

            leftFlow.Controls.Add(btnSelectAll)
            leftFlow.Controls.Add(btnDeselectAll)
            leftFlow.Controls.Add(lblCount)

            Dim btnCancel As New Button() With {
                .Text = "Cancel",
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlatStyle = FlatStyle.Flat,
                .Font = standardFont,
                .DialogResult = System.Windows.Forms.DialogResult.Cancel,
                .Padding = buttonPadding,
                .Margin = New Padding(0, 0, 0, 0)
            }
            btnCancel.FlatAppearance.BorderColor = Drawing.Color.FromArgb(180, 180, 180)

            Dim btnMove As New Button() With {
                .Text = "Move Mails",
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .FlatStyle = FlatStyle.Flat,
                .Font = standardFont,
                .Padding = buttonPadding,
                .Margin = New Padding(0, 0, btnSpacing, 0)
            }
            btnMove.FlatAppearance.BorderColor = Drawing.Color.FromArgb(180, 180, 180)

            rightFlow.Controls.Add(btnCancel)
            rightFlow.Controls.Add(btnMove)

            btnLayout.Controls.Add(leftFlow, 0, 0)
            btnLayout.Controls.Add(New Panel() With {.Dock = DockStyle.Fill}, 1, 0)
            btnLayout.Controls.Add(rightFlow, 2, 0)

            ' Update count helper
            Dim updateCount As New System.Action(
                Sub()
                    Dim moveCount As Integer = 0
                    For Each r As DataGridViewRow In dgv.Rows
                        If CBool(If(r.Cells("colMove").Value, False)) AndAlso
                           CStr(If(r.Cells("colFolder").Value, "")) <> "(skip - do not move)" Then
                            moveCount += 1
                        End If
                    Next
                    lblCount.Text = $"{moveCount} of {mails.Count} will be moved"
                End Sub)

            AddHandler dgv.CellValueChanged, Sub(s, e) updateCount.Invoke()
            AddHandler dgv.CurrentCellDirtyStateChanged, Sub(s, e)
                                                             If dgv.IsCurrentCellDirty Then dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                                                         End Sub

            AddHandler btnSelectAll.Click, Sub(s, e)
                                               For Each row As DataGridViewRow In dgv.Rows
                                                   row.Cells("colMove").Value = True
                                               Next
                                           End Sub

            AddHandler btnDeselectAll.Click, Sub(s, e)
                                                 For Each row As DataGridViewRow In dgv.Rows
                                                     row.Cells("colMove").Value = False
                                                 Next
                                             End Sub

            AddHandler btnMove.Click, Sub(s, e)
                                          For Each row As DataGridViewRow In dgv.Rows
                                              Dim mailEntry As MailMoverEntry = DirectCast(row.Tag, MailMoverEntry)
                                              Dim shouldMove As Boolean = CBool(If(row.Cells("colMove").Value, False))
                                              Dim selFolder As String = CStr(If(row.Cells("colFolder").Value, ""))

                                              If shouldMove AndAlso selFolder <> "(skip - do not move)" Then
                                                  mailEntry.SelectedFolder = selFolder
                                              Else
                                                  mailEntry.SelectedFolder = Nothing
                                              End If
                                          Next

                                          Dim moveCount As Integer = mails.Where(Function(m) m.SelectedFolder IsNot Nothing).Count()
                                          If moveCount = 0 Then
                                              ShowCustomMessageBox("No mails selected for moving.", $"{AN} - Mail Mover")
                                              Return
                                          End If

                                          approved = True
                                          frm.Close()
                                      End Sub

            frm.AcceptButton = btnMove
            frm.CancelButton = btnCancel

            AddHandler frm.Shown,
                Sub()
                    Dim minW As Integer = 1100
                    Dim minH As Integer = Math.Max(360, headerLabel.Height + dgv.ColumnHeadersHeight + rightFlow.PreferredSize.Height + mainLayout.Padding.Vertical + 40)
                    frm.MinimumSize = New Drawing.Size(minW, minH)
                End Sub

            updateCount.Invoke()
            frm.ShowDialog()
        End Using

        Return approved
    End Function

#End Region

#Region "MailMover Move Execution"

    Private Sub MoveApprovedMails(mails As List(Of MailMoverEntry))
        Dim toMove As List(Of MailMoverEntry) = mails.Where(Function(m) m.SelectedFolder IsNot Nothing).ToList()
        If toMove.Count = 0 Then Return

        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")

        Dim undoList As New List(Of MailMoverUndoEntry)()
        Dim movedCount As Integer = 0
        Dim failedCount As Integer = 0
        Dim failedNames As New List(Of String)()

        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressMax = toMove.Count
        ProgressBarModule.GlobalProgressLabel = "Moving mails..."
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.ShowProgressBarInSeparateThread($"{AN} - Mail Mover", "Moving mails...")

        Try
            For i As Integer = 0 To toMove.Count - 1
                If ProgressBarModule.CancelOperation Then Exit For

                ProgressBarModule.GlobalProgressValue = i
                ProgressBarModule.GlobalProgressLabel = $"Moving {i + 1}/{toMove.Count}..."

                Dim entry As MailMoverEntry = toMove(i)
                Try
                    Dim entryId As String = entry.EntryID
                    Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
                    If mi Is Nothing Then
                        failedCount += 1
                        failedNames.Add(entry.Subject)
                        Continue For
                    End If

                    Dim targetFolder As MAPIFolder = FindFolderByPath(ns, entry.SelectedFolder)
                    If targetFolder Is Nothing Then
                        failedCount += 1
                        failedNames.Add($"{entry.Subject} (folder not found: {entry.SelectedFolder})")
                        Continue For
                    End If

                    undoList.Add(New MailMoverUndoEntry With {
                        .EntryID = entry.EntryID,
                        .OriginalFolderPath = entry.SourceFolderPath,
                        .TargetFolderPath = entry.SelectedFolder
                    })

                    Dim movedItem As Object = mi.Move(targetFolder)

                    If movedItem IsNot Nothing Then
                        Dim movedMail As MailItem = TryCast(movedItem, MailItem)
                        If movedMail IsNot Nothing Then
                            undoList(undoList.Count - 1).EntryID = movedMail.EntryID
                        End If
                    End If

                    movedCount += 1
                Catch ex As System.Exception
                    failedCount += 1
                    failedNames.Add($"{entry.Subject}: {ex.Message}")
                End Try
            Next
        Finally
            ProgressBarModule.CancelOperation = True
        End Try

        _mailMoverUndoList = If(undoList.Count > 0, undoList, Nothing)

        Dim summary As New StringBuilder()
        summary.AppendLine($"Successfully moved: {movedCount} mail(s)")
        If failedCount > 0 Then
            summary.AppendLine()
            summary.AppendLine($"Failed: {failedCount} mail(s)")
            For Each f In failedNames.Take(10)
                summary.AppendLine($"  {ChrW(&H2022)} {f}")
            Next
            If failedNames.Count > 10 Then
                summary.AppendLine($"  ... and {failedNames.Count - 10} more")
            End If
        End If
        If movedCount > 0 Then
            summary.AppendLine()
            summary.AppendLine("You can undo this operation by again calling up Mail Mover and selecting 'Undo'.")
        End If

        ShowCustomMessageBox(summary.ToString().TrimEnd(), $"{AN} - Mail Mover")
    End Sub

#End Region

End Class