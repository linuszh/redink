' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.vb
' Purpose: Inky AutoPilot — AI mail watcher that monitors incoming e-mails,
'          filters them by domain/sender/subject rules, processes them via the
'          LLM tooling loop, and replies to the sender with AI-generated responses.
'
' Architecture:
'  - Event-based mail watching via Outlook's Application.NewMailEx event.
'  - Dual-mode operation: AutoPilot (auto-send for whitelisted senders) and
'    CoPilot (approval required, with holding response for non-whitelisted).
'  - Filter engine: domain wildcards, sender wildcards, subject trigger words,
'    folder filtering, negative filters, auto-reply/OOF detection, loop prevention.
'  - Anti-loop safeguards: custom MAPI property, per-sender cooldown, max replies
'    per session, thread depth limit, auto-reply header detection.
'  - Stateless: each incoming mail is processed independently using the e-mail
'    thread body for context.
'  - Integrates with existing ExecuteToolingLoop for tool calls (document processing,
'    PDF tools, web retrieval, and all existing specialservices tools).
'  - Dashboard via extended LogWindow for live monitoring.
'  - Replies are placed in Sent Items\Inky Replies subfolder.
'
' Security Model for Attachments:
'  - Each incoming e-mail gets a unique temp directory (GUID-based) under %TEMP%.
'  - Attachments are saved into this isolated directory before processing.
'  - Output files (processed docs, merged PDFs, compare docs) are also written
'    into the same per-mail temp directory.
'  - CollectResultAttachments only picks up files registered in OutputFiles AND
'    validates they reside within the per-mail temp directory (path prefix check).
'  - After the reply is sent (or if processing fails/is rejected), the entire
'    temp directory is deleted recursively in the Finally block.
'  - This ensures no sensitive attachment data persists on disk after processing.
'
' Model Override Command (#model:):
'  - The sender can include "#model: <ModelName>" in the first few lines of the
'    latest e-mail body (not in quoted replies) to request a specific model.
'  - Only the latest/top-most portion of the body is scanned (before reply separators).
'  - The command line is stripped from the body before passing to the LLM.
'  - The model is applied only for this single e-mail, then restored.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Concurrent
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Outlook
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AP_LoopHeaderProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoReply"
    Private Const AP_LoopHeaderValue As String = "true"
    Private Const AP_CategoryName As String = "Inky AutoPilot"
    Private Const AP_SentSubfolder As String = "Inky Replies"
    Private Const AP_MaxThreadDepth As Integer = 20
    Private Const AP_DefaultCooldownSeconds As Integer = 60
    Private Const AP_DefaultMaxRepliesPerSession As Integer = 200
    Private Const AP_DefaultMaxAttachmentBytes As Long = 10 * 1024 * 1024
    Private Const AP_TempPrefix As String = AN2 & "_autopilot_"

    ''' <summary>Command prefix scanned in the first few lines of the latest e-mail body.</summary>
    Private Const AP_ModelCommandPrefix As String = "#model:"

    ''' <summary>Max lines from top of latest mail body to scan for #model: command.</summary>
    Private Const AP_ModelCommandScanLines As Integer = 5

    Private Const SP_AutoPilot_HoldingResponse As String =
        "Thank you for your message. I have received it and will respond once your request has been approved. " &
        "— " & AN6

    Private Const AP_MaxToolIterations As Integer = 30

    ' ═══════════════════════════════════════════════════════════════════════════
    '  STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    Private _apActive As Boolean = False
    Private _apConfig As AutoPilotConfig = Nothing
    Private _apCts As CancellationTokenSource = Nothing
    Private _apDashboard As LogWindow = Nothing
    Private ReadOnly _apSenderCooldowns As New ConcurrentDictionary(Of String, DateTime)(StringComparer.OrdinalIgnoreCase)
    Private _apSessionReplyCount As Integer = 0
    Private ReadOnly _apMailQueue As New ConcurrentQueue(Of String)()
    Private ReadOnly _apProcessingSemaphore As New SemaphoreSlim(1, 1)
    Private _apSelectedTools As List(Of ModelConfig) = Nothing
    Private _apUseSecondApi As Boolean = False

    Private Shared ReadOnly AP_AutoReplySubjectPatterns As String() = {
        "automatic reply:", "automatische antwort:", "réponse automatique:",
        "risposta automatica:", "out of office:", "abwesenheitsnotiz:",
        "absence du bureau:", "fuori sede:", "auto-reply:", "auto reply:",
        "undeliverable:", "nicht zustellbar:", "delivery status notification",
        "mail delivery failed", "returned mail:"
    }

    ' ── Per-mail transient state ──
    Private _apCurrentTempDir As String = Nothing
    Private _apCurrentAttachments As List(Of AutoPilotAttachmentInfo) = Nothing
    Private _apCurrentMailInfo As AutoPilotMailInfo = Nothing

    ''' <summary>Tracks tool calls during processing of a single e-mail for the "Sources used:" footer.</summary>
    Private _apCurrentToolCallLog As List(Of AutoPilotToolCallEntry) = Nothing

    ''' <summary>Captured at AutoPilot startup — the model config to re-apply before every LLM call.</summary>
    Private _apBaseModelConfig As ModelConfig = Nothing

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PUBLIC ENTRY POINTS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Stops AutoPilot processing and clears runtime state.</summary>
    Public Sub StopAutoPilot()
        If Not _apActive Then Return
        _apActive = False
        Try : RemoveHandler Application.NewMailEx, AddressOf AutoPilot_NewMailEx : Catch : End Try
        Try
            If _apDashboard IsNot Nothing Then
                RemoveHandler _apDashboard.CancelRequested, AddressOf AutoPilot_DashboardCancelRequested
            End If
        Catch : End Try
        Try : _apCts?.Cancel() : Catch : End Try
        Try : _apCts?.Dispose() : Catch : End Try
        _apCts = Nothing
        Dim dummy As String = Nothing
        While _apMailQueue.TryDequeue(dummy) : End While
        _apSenderCooldowns.Clear()
        _apSessionReplyCount = 0
        _apSelectedTools = Nothing
        _apConfig = Nothing
        _apBaseModelConfig = Nothing
        _apSenderCooldowns.Clear()
        _apProcessedConversations.Clear()
        ApDashboardLog("AutoPilot stopped.", "info")
        ApDashboardMarkComplete()
        ShowCustomMessageBox($"{AN6} AutoPilot has been stopped.", AN)
    End Sub

    ''' <summary>Shows the AutoPilot dashboard if it is available.</summary>
    Public Sub ShowAutoPilotDashboard()
        If _apDashboard IsNot Nothing Then
            Try
                _apDashboard.Show()
                _apDashboard.BringToFront()
            Catch
            End Try
        End If
    End Sub

    ''' <summary>Starts AutoPilot using the configuration dialog and saved settings.</summary>
    Public Sub StartAutoPilot()
        If Not IsAutoPilotPermitted() Then
            ShowCustomMessageBox(
                $"{AN6} AutoPilot is not available for your user account ({Environment.UserName}). Contact your administrator to request access.",
                AN)
            Return
        End If
        If _apActive Then
            Dim choice = ShowCustomYesNoBox(
                AN6 & " AutoPilot is already running." & vbCrLf &
                "Would you like to stop it or view the dashboard?",
                "Stop AutoPilot", "Show Dashboard",
                header:=AN6 & " AutoPilot")
            If choice = 1 Then StopAutoPilot()
            If choice = 2 Then ShowAutoPilotDashboard()
            Return
        End If
        Dim config = ShowAutoPilotConfigDialog()
        If config Is Nothing Then Return
        SaveAutoPilotConfigToSettings(config)
        StartAutoPilotWithConfig(config)
    End Sub

    ''' <summary>Starts AutoPilot with an explicit configuration instance.</summary>
    Private Sub StartAutoPilotWithConfig(config As AutoPilotConfig)
        _apConfig = config
        _apActive = True
        _apCts = New CancellationTokenSource()
        _apSessionReplyCount = 0
        _apSenderCooldowns.Clear()
        _apUseSecondApi = config.UseSecondApi

        ' Capture the current model config at startup so we can re-apply it
        ' before every LLM call, ensuring tooling templates are always pristine.
        _apBaseModelConfig = GetCurrentConfig(_context)

        ' Determine if the selected model supports tooling
        Dim modelCanCallTools As Boolean = ModelSupportsTooling(_apBaseModelConfig)

        ' Only load tools if the model can actually call them
        If modelCanCallTools Then
            _apSelectedTools = New List(Of ModelConfig)()
            _apSelectedTools.AddRange(GetAutoPilotInternalTools())
            If config.SelectedExternalTools IsNot Nothing Then _apSelectedTools.AddRange(config.SelectedExternalTools)
        Else
            _apSelectedTools = Nothing
        End If

        _apDashboard = New LogWindow()
        _apDashboard.Text = $"{AN6} AutoPilot Dashboard"
        AddHandler _apDashboard.CancelRequested, AddressOf AutoPilot_DashboardCancelRequested
        _apDashboard.Show()

        Dim modelLabel As String
        If Not String.IsNullOrWhiteSpace(config.SelectedModelKey) Then
            modelLabel = config.SelectedModelKey
        ElseIf config.UseSecondApi Then
            modelLabel = INI_Model_2
        Else
            modelLabel = INI_Model
        End If

        ApDashboardLog($"AutoPilot started.", "info")
        ApDashboardLog($"Model: {modelLabel}", "info")

        If Not modelCanCallTools Then
            ApDashboardLog("⚠ Model does not support tool calling — running without tools", "warn")
        Else
            ApDashboardLog($"Filters: {config.FilterRules.Count} rule(s) active", "info")
            ApDashboardLog($"{ToolFriendlyName}: {_apSelectedTools.Count} available (incl. built-in)", "info")

            ' Log each tool so operator can verify the court database tool is loaded
            For Each tool In _apSelectedTools
                Dim toolLabel = If(Not String.IsNullOrEmpty(tool.ModelDescription), tool.ModelDescription,
                               If(Not String.IsNullOrEmpty(tool.ToolName), tool.ToolName, "(unnamed)"))
                Dim hasDefinition = Not String.IsNullOrWhiteSpace(tool.ToolDefinition)
                Dim hasInstructions = Not String.IsNullOrWhiteSpace(tool.ToolInstructionsPrompt)
                ApDashboardLog($"  • {toolLabel} [def:{hasDefinition}, instr:{hasInstructions}]", "step")
            Next

            ' Log tooling config to verify detection/extraction patterns are set
            'ApDashboardLog($"ToolCallDetectionPattern: {If(Not String.IsNullOrWhiteSpace(_context.INI_ToolCallDetectionPattern_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolCallExtractionMap: {If(Not String.IsNullOrWhiteSpace(_context.INI_ToolCallExtractionMap_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolInstructions template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolInstructions_Template_2), "set (" & _context.INI_APICall_ToolInstructions_Template_2.Length.ToString() & " chars)", "EMPTY")}", "step")
            'ApDashboardLog($"ToolResponses template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolResponses_Template_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"ToolCallPart template: {If(Not String.IsNullOrWhiteSpace(_context.INI_APICall_ToolCallPart_Template_2), "set", "EMPTY")}", "step")
            'ApDashboardLog($"APICall_2 contains toolinstructions placeholder: {If(_context.INI_APICall_2 IsNot Nothing AndAlso _context.INI_APICall_2.Contains("toolinstructions"), "YES", "NO")}", "step")
        End If

        ApDashboardLog($"Filters: {config.FilterRules.Count} rule(s) active", "info")
        ApDashboardLog($"Mode: {If(config.RequireApprovalForNonWhitelisted, "CoPilot (approval for non-whitelisted)", "AutoPilot (auto-send all)")}", "info")
        ApDashboardLog("Watching for new mail...", "info")

        CatchUpMissedMails()

        AddHandler Application.NewMailEx, AddressOf AutoPilot_NewMailEx
        Task.Run(Function() AutoPilotProcessingPump(_apCts.Token))
    End Sub

    ''' <summary>Handles a dashboard stop request.</summary>
    Private Sub AutoPilot_DashboardCancelRequested(sender As Object, e As EventArgs)
        StopAutoPilot()
    End Sub

    ''' <summary>
    ''' Scans the Inbox (or monitored folder) for mails that arrived after the last
    ''' processed timestamp, applies the full filter/trigger logic, and shows a
    ''' preview dialog where the operator can check/uncheck individual mails.
    ''' Only mails that have NOT already been processed (no AP_CategoryName tag) are shown.
    ''' </summary>
    Private Sub CatchUpMissedMails()
        Try
            Dim lastProcessed As DateTime = My.Settings.AP_LastProcessedTime
            Dim isFirstRun As Boolean = False

            If lastProcessed = DateTime.MinValue OrElse lastProcessed = New DateTime(1, 1, 1) Then
                ' No previous session timestamp — default to 24 hours ago in local time.
                ' We use local time directly because MailItem.ReceivedTime is always local,
                ' and we want exactly 24 hours of local-time coverage.
                lastProcessed = DateTime.Now.AddHours(-24)
                isFirstRun = True
                ApDashboardLog("No previous session timestamp — scanning last 24 hours for relevant mails.", "info")
            End If

            ' For persisted values: ensure we treat the stored value as UTC
            ' (it was saved as DateTime.UtcNow) and convert to local for comparison.
            ' For the first-run fallback: already local, skip conversion.
            Dim lastProcessedLocal As DateTime
            If isFirstRun Then
                lastProcessedLocal = lastProcessed
            Else
                If lastProcessed.Kind = DateTimeKind.Unspecified Then
                    lastProcessed = DateTime.SpecifyKind(lastProcessed, DateTimeKind.Utc)
                End If
                lastProcessedLocal = lastProcessed.ToLocalTime()
            End If

            lastProcessedLocal = DateTime.Now.AddHours(-24)

            If isFirstRun Then
                ApDashboardLog($"Scanning for mails received after: {lastProcessedLocal:yyyy-MM-dd HH:mm:ss} (local)", "info")
            Else
                ApDashboardLog($"Last processed: {lastProcessedLocal:yyyy-MM-dd HH:mm:ss} (local) / {lastProcessed:yyyy-MM-dd HH:mm:ss} (UTC)", "info")
            End If
            ApDashboardLog("Scanning inbox for newer mails...", "info")

            Dim ns = Application.GetNamespace("MAPI")
            Dim inbox As MAPIFolder = Nothing

            ' Resolve the correct inbox when a specific mailbox is monitored
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Try
                    For i As Integer = 1 To ns.Accounts.Count
                        Dim acct = ns.Accounts(i)
                        If acct.SmtpAddress.Equals(_apConfig.MonitoredMailbox, StringComparison.OrdinalIgnoreCase) Then
                            Dim deliveryStore = acct.DeliveryStore
                            If deliveryStore IsNot Nothing Then
                                inbox = deliveryStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
                            End If
                            Exit For
                        End If
                    Next
                Catch
                End Try
            End If

            ' Fallback to default inbox
            If inbox Is Nothing Then
                inbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
            End If

            ' Sort descending so we process newest first and can stop early
            ' once we hit mails older than lastProcessedLocal.
            Dim allItems = inbox.Items
            allItems.Sort("[ReceivedTime]", Descending:=True)

            Dim candidates As New List(Of CatchUpCandidate)()
            Dim skippedAlreadyProcessed As Integer = 0
            Dim skippedOther As Integer = 0
            Dim totalScanned As Integer = 0

            For Each item As Object In allItems
                If Not TypeOf item Is MailItem Then Continue For
                Dim mi = DirectCast(item, MailItem)
                Try

                    ' MailItem.ReceivedTime is always local time.
                    ' Stop scanning once we've gone past the cutoff.
                    If mi.ReceivedTime <= lastProcessedLocal Then
                        Exit For
                    End If

                    totalScanned += 1

                    ' 1. Skip our own auto-replies (header on the mail itself)
                    Try
                        Dim prop = mi.PropertyAccessor.GetProperty(AP_LoopHeaderProperty)
                        If prop IsNot Nothing AndAlso prop.ToString() = AP_LoopHeaderValue Then
                            Continue For
                        End If
                    Catch
                    End Try

                    ' 2. Skip mails tagged with AP category (fast check, kept as fallback)
                    Try
                        Dim cats = mi.Categories
                        If cats IsNot Nothing AndAlso cats.IndexOf(AP_CategoryName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            skippedAlreadyProcessed += 1
                            Continue For
                        End If
                    Catch
                    End Try

                    ' 3. Skip mails already replied to by AutoPilot — walk the conversation
                    '    to find any reply with the X-RedInk-AutoReply header.
                    '    This is the authoritative check and does not depend on the category tag.
                    Try
                        If IsPartOfAutoPilotConversation(mi) Then
                            skippedAlreadyProcessed += 1
                            Continue For
                        End If
                    Catch
                    End Try

                    ' Full pre-filter: same logic as ProcessIncomingMailAsync
                    Dim mailInfo = ExtractMailInfo(mi)
                    If mailInfo Is Nothing Then Continue For
                    If mailInfo.HasAutoReplyHeader Then Continue For
                    If IsAutoReplyOrOof(mailInfo) Then
                        skippedOther += 1
                        Continue For
                    End If
                    If Not MatchesFilterRules(mailInfo) Then
                        skippedOther += 1
                        Continue For
                    End If
                    If MatchesNegativeFilters(mailInfo) Then
                        skippedOther += 1
                        Continue For
                    End If

                    ' Subject trigger word check
                    If Not String.IsNullOrWhiteSpace(_apConfig.SubjectTriggerWord) Then
                        If mailInfo.Subject.IndexOf(_apConfig.SubjectTriggerWord, StringComparison.OrdinalIgnoreCase) < 0 Then
                            skippedOther += 1
                            Continue For
                        End If
                    End If

                    candidates.Add(New CatchUpCandidate() With {
                        .EntryID = mi.EntryID,
                        .SenderName = mailInfo.SenderName,
                        .SenderEmail = mailInfo.SenderEmail,
                        .Subject = mailInfo.Subject,
                        .ReceivedTime = mailInfo.ReceivedTime
                    })
                Finally
                    Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                End Try
            Next

            ApDashboardLog($"Scan complete: {totalScanned} mail(s) scanned, {skippedAlreadyProcessed} already processed, {skippedOther} filtered out.", "step")

            If candidates.Count = 0 Then
                ApDashboardLog("No missed mails found matching filters.", "step")
                Return
            End If

            ' Re-sort candidates oldest-first for processing order
            candidates.Sort(Function(a, b) a.ReceivedTime.CompareTo(b.ReceivedTime))

            ApDashboardLog($"Found {candidates.Count} unprocessed mail(s) matching filters.", "info")

            ' Show preview dialog using MultiModelSelectorForm pattern
            Dim selectedEntryIds = ShowCatchUpPreviewDialog(candidates, lastProcessed)

            If selectedEntryIds IsNot Nothing AndAlso selectedEntryIds.Count > 0 Then
                For Each entryId In selectedEntryIds
                    _apMailQueue.Enqueue(entryId)
                Next
                ApDashboardLog($"Queued {selectedEntryIds.Count} of {candidates.Count} missed mail(s) for processing.", "info")
            Else
                ApDashboardLog("Operator skipped catch-up processing.", "step")
            End If

        Catch ex As System.Exception
            ApDashboardLog($"Catch-up scan error: {ex.Message}", "warn")
            Debug.WriteLine($"[AutoPilot] CatchUpMissedMails error: {ex}")
        End Try
    End Sub

    ''' <summary>Data class for catch-up preview items.</summary>
    Private Class CatchUpCandidate
        Public Property EntryID As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Subject As String
        Public Property ReceivedTime As DateTime

        ''' <summary>Returns a display label for the checked list box.</summary>
        Public Function ToDisplayLabel() As String
            Dim timeStr = ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            Dim senderStr = If(Not String.IsNullOrWhiteSpace(SenderName), SenderName, SenderEmail)
            Dim subjectStr = If(Subject.Length > 60, Subject.Substring(0, 57) & "...", Subject)
            Return $"[{timeStr}]  {senderStr}  —  {subjectStr}"
        End Function
    End Class

    ''' <summary>
    ''' Shows a preview dialog listing all catch-up candidate mails with checkboxes.
    ''' All items are pre-checked. The operator can uncheck mails to exclude them.
    ''' Returns the list of EntryIDs for checked items, or Nothing if cancelled.
    ''' </summary>
    Private Function ShowCatchUpPreviewDialog(candidates As List(Of CatchUpCandidate), lastProcessed As DateTime) As List(Of String)
        ' Build ModelConfig list to feed into MultiModelSelectorForm
        ' We repurpose ModelDescription as the display label and Model as the EntryID
        Dim items As New List(Of ModelConfig)()
        For Each c In candidates
            items.Add(New ModelConfig() With {
                .ModelDescription = c.ToDisplayLabel(),
                .Model = c.EntryID
            })
        Next

        ' Pre-select all items by their display labels
        Dim allLabels = items.Select(Function(m) m.ModelDescription).ToList()

        Using dlg As New SharedLibrary.SharedLibrary.MultiModelSelectorForm(
            items, Nothing,
            title:=$"{AN6} AutoPilot — Catch-Up ({candidates.Count} mail(s) since {lastProcessed:yyyy-MM-dd HH:mm})",
            resetChecked:=False,
            preselectMany:=allLabels,
            instruction:=$"The following {candidates.Count} mail(s) arrived while AutoPilot was inactive. Select the mails to process:")

            ' Center on screen since there is no parent form in this context
            dlg.StartPosition = FormStartPosition.CenterScreen

            If dlg.ShowDialog() = DialogResult.OK Then
                Dim selected = dlg.SelectedModels
                If selected IsNot Nothing AndAlso selected.Count > 0 Then
                    Return selected.Select(Function(m) m.Model).ToList()
                End If
                Return New List(Of String)()
            End If
        End Using

        Return Nothing
    End Function

    ''' <summary>Queues newly received mail entry IDs for processing.</summary>
    Private Sub AutoPilot_NewMailEx(ByVal EntryIDCollection As String)
        If Not _apActive Then Return
        For Each id In EntryIDCollection.Split(","c)
            Dim trimmed = id.Trim()
            If trimmed.Length > 0 Then _apMailQueue.Enqueue(trimmed)
        Next
    End Sub

    ''' <summary>Processes queued messages until cancellation is requested.</summary>
    Private Async Function AutoPilotProcessingPump(ct As CancellationToken) As Task
        Try
            While Not ct.IsCancellationRequested
                Dim entryId As String = Nothing
                If _apMailQueue.TryDequeue(entryId) Then
                    Dim pending = _apMailQueue.Count
                    If pending > 0 Then
                        ApDashboardLog($"⏳ {pending} mail(s) queued behind current processing", "step")
                    End If
                    Await ProcessIncomingMailAsync(entryId, ct)
                Else
                    Await Task.Delay(1000, ct)
                End If
            End While
        Catch ex As OperationCanceledException
            ApDashboardLog("AutoPilot cancelled.", "step")
        Catch ex As System.Exception
            ApDashboardLog("AutoPilot pump error: " & ex.Message, "error")
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  AUTOPILOT USER PERMISSION CHECK
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Checks whether the current Windows user is permitted to run AutoPilot.
    ''' If INI_AutoPilot is empty or whitespace, any user is permitted.
    ''' Otherwise, INI_AutoPilot must contain a comma-separated list of usernames,
    ''' and the current %USERNAME% must appear in that list (case-insensitive).
    ''' </summary>
    ''' <returns>True if the current user is allowed; False otherwise.</returns>
    Public Shared Function IsAutoPilotPermitted() As Boolean
        Dim allowedUsers As String = INI_AutoPilot
        If String.IsNullOrWhiteSpace(allowedUsers) Then Return True

        Dim currentUser As String = Environment.GetEnvironmentVariable("USERNAME")
        If String.IsNullOrWhiteSpace(currentUser) Then Return False

        Dim users = allowedUsers.Split(","c)
        For Each entry In users
            Dim trimmed = entry.Trim()
            If trimmed.Length > 0 AndAlso trimmed.Equals(currentUser, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next

        Return False
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CORE PROCESSING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Processes a single incoming mail item by entry ID.</summary>
    Private Async Function ProcessIncomingMailAsync(entryId As String, ct As CancellationToken) As Task
        Dim mi As MailItem = Nothing

        Try
            mi = Await SwitchToUi(Function() As MailItem
                                      Try
                                          Dim ns = Application.GetNamespace("MAPI")
                                          Dim obj = ns.GetItemFromID(entryId)
                                          If TypeOf obj Is MailItem Then Return DirectCast(obj, MailItem)
                                      Catch
                                      End Try
                                      Return Nothing
                                  End Function)
            If mi Is Nothing Then
                ApDashboardLog("SKIP (could not resolve mail item from EntryID — may have been moved or is not a MailItem)", "warn")
                Return
            End If

            ' Mailbox filter
            If Not String.IsNullOrWhiteSpace(_apConfig.MonitoredMailbox) Then
                Dim recipientAddress As String = ""
                Try
                    recipientAddress = Await SwitchToUi(Function() As String
                                                            Try
                                                                Dim acct = mi.SendUsingAccount
                                                                If acct IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(acct.SmtpAddress) Then Return acct.SmtpAddress
                                                                For Each recip As Recipient In mi.Recipients
                                                                    If recip.Type = CInt(OlMailRecipientType.olTo) Then
                                                                        Dim addr = recip.Address
                                                                        If Not String.IsNullOrWhiteSpace(addr) Then Return addr
                                                                    End If
                                                                Next
                                                            Catch
                                                            End Try
                                                            Return ""
                                                        End Function)
                Catch
                End Try
                If Not _apConfig.MonitoredMailbox.Equals(recipientAddress, StringComparison.OrdinalIgnoreCase) Then
                    ApDashboardLog("SKIP (wrong mailbox: " & If(String.IsNullOrWhiteSpace(recipientAddress), "(empty)", recipientAddress) & ")", "step")
                    Return
                End If
            End If

            Dim mailInfo = Await SwitchToUi(Function() ExtractMailInfo(mi))
            If mailInfo Is Nothing Then
                ApDashboardLog("SKIP (ExtractMailInfo returned Nothing — COM error or unexpected mail format)", "warn")
                Return
            End If

            ' ── Filter checks ──

            ' Loop prevention: skip if this mail itself is an auto-reply from us
            If mailInfo.HasAutoReplyHeader Then : ApDashboardLog("SKIP (own auto-reply): " & mailInfo.Subject, "step") : Return : End If

            ' Conversation awareness: detect if this is a follow-up in an existing AutoPilot thread.
            ' This is needed so that second/subsequent user replies to an AutoPilot response are
            ' still processed, even though the user's mail itself has no X-RedInk-AutoReply header.
            Dim isExistingConversation As Boolean = False
            Try
                isExistingConversation = IsPartOfAutoPilotConversation(mi)
            Catch : End Try

            If IsAutoReplyOrOof(mailInfo) Then : ApDashboardLog("SKIP (auto-reply/OOF): " & mailInfo.Subject, "step") : Return : End If

            ' SECURITY: Filter rules (domain/sender) are NEVER bypassed — they are the trust boundary.
            ' An attacker could spoof a ConversationID (same subject line) to join an existing thread,
            ' so we must always verify the sender is permitted regardless of conversation membership.
            If Not MatchesFilterRules(mailInfo) Then : ApDashboardLog("SKIP (no filter match): " & mailInfo.SenderEmail & " — " & mailInfo.Subject, "step") : Return : End If
            If MatchesNegativeFilters(mailInfo) Then : ApDashboardLog("SKIP (negative filter): " & mailInfo.SenderEmail & " — " & mailInfo.Subject, "step") : Return : End If

            ' Subject trigger word CAN be bypassed for existing conversations — reply subjects
            ' often have "Re: Re: ..." prefixes or the trigger word gets stripped in threading.
            If Not String.IsNullOrWhiteSpace(_apConfig.SubjectTriggerWord) Then
                If Not mailInfo.Subject.IndexOf(_apConfig.SubjectTriggerWord, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    If isExistingConversation Then
                        ApDashboardLog("TRIGGER BYPASS (existing conversation): " & mailInfo.Subject, "step")
                    Else
                        ApDashboardLog("SKIP (no trigger word): " & mailInfo.Subject, "step") : Return
                    End If
                End If
            End If

            If IsSenderOnCooldown(mailInfo.SenderEmail) Then : ApDashboardLog("SKIP (cooldown): " & mailInfo.SenderEmail, "step") : Return : End If
            If _apSessionReplyCount >= _apConfig.MaxRepliesPerSession Then : ApDashboardLog("SKIP (session limit " & _apConfig.MaxRepliesPerSession.ToString() & " reached)", "warn") : Return : End If
            If mailInfo.ThreadAIReplyCount >= AP_MaxThreadDepth Then : ApDashboardLog("SKIP (thread depth " & mailInfo.ThreadAIReplyCount.ToString() & " >= " & AP_MaxThreadDepth.ToString() & "): " & mailInfo.Subject, "warn") : Return : End If

            Dim isWhitelisted As Boolean = IsSenderWhitelisted(mailInfo.SenderEmail)
            ' Existing conversations: skip approval only (sender already passed domain/sender filter above)
            Dim requiresApproval As Boolean = _apConfig.RequireApprovalForNonWhitelisted AndAlso Not isWhitelisted AndAlso Not isExistingConversation
            ApDashboardLog("━━━ PROCESSING ━━━", "info")
            ApDashboardLog($"From: {mailInfo.SenderName} <{mailInfo.SenderEmail}>", "info")
            ApDashboardLog($"Subject: {mailInfo.Subject}", "info")
            ApDashboardLog($"Attachments: {mailInfo.AttachmentCount}" & If(requiresApproval, " [approval required]", " [auto-send]"), "info")

            ' Check for #model: command
            Dim modelOverrideConfig As ModelConfig = Nothing
            Dim modelOverrideName As String = Nothing
            Dim strippedBody As String = mailInfo.Body
            TryParseModelOverride(mailInfo.Body, modelOverrideConfig, modelOverrideName, strippedBody)
            If modelOverrideConfig IsNot Nothing Then
                ApDashboardLog($"Model override: {modelOverrideName}", "info")
                mailInfo.Body = strippedBody ' Remove the #model: line from the body sent to the LLM
            End If

            ' ── Extract attachments to temp ──
            Dim tempDir As String = Path.Combine(Path.GetTempPath(), AP_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempDir)
            Dim attachmentPaths As New List(Of AutoPilotAttachmentInfo)()

            Try
                attachmentPaths = Await SwitchToUi(Function() SaveAttachmentsToTemp(mi, tempDir))

                If attachmentPaths.Count > 0 Then
                    ApDashboardLog($"Saved {attachmentPaths.Count} attachment(s) to temp:", "info")
                    For Each att In attachmentPaths
                        Dim sizeStr = If(att.SizeBytes > 0, $"{att.SizeBytes / 1024:F0} KB", "?")
                        Dim statusStr = If(att.IsOverSizeLimit, " [OVER LIMIT]", "")
                        ApDashboardLog($"  • {att.OriginalFileName} ({att.Extension}, {sizeStr}){statusStr}", "info")
                    Next
                End If

                ' Initialize tool call log for this e-mail
                _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()

                ' ── Build LLM prompt ──
                Dim userPrompt As String = BuildUserPromptFromMail(mailInfo, attachmentPaths)
                Dim systemPrompt As String = InterpolateAtRuntime(SP_AutoPilot)

                ' ── Set AutoPilot tool context ──
                _apCurrentTempDir = tempDir
                _apCurrentAttachments = attachmentPaths
                _apCurrentMailInfo = mailInfo

                Dim response As String

                ' ── Set AutoPilot-specific max tool iterations ──
                Dim previousMaxToolIterations = MaxToolIterations
                MaxToolIterations = AP_MaxToolIterations

                Try
                    ApDashboardLog("Calling AI model...", "llm")

                    ' Re-apply the base model config captured at startup.
                    ' This ensures all tooling templates, detection patterns, and API call
                    ' structures are pristine, regardless of what ExecuteToolingLoop's internal
                    ' ApplyModelConfig/RestoreDefaults cycles did during the previous mail.
                    ApplyModelConfig(_context, _apBaseModelConfig)

                    ' Apply model override on top if requested
                    Dim overrideSupportsTooling As Boolean = True
                    If modelOverrideConfig IsNot Nothing Then
                        ' Check if the override model supports tooling
                        overrideSupportsTooling = ModelSupportsTooling(modelOverrideConfig)

                        If overrideSupportsTooling Then
                            ' Override model has its own tooling config — apply the full config
                            ' so its detection patterns, extraction maps, and templates are used.
                            ApplyModelConfig(_context, modelOverrideConfig)
                            ApDashboardLog($"Applied full model override: {modelOverrideName} (tooling supported)", "info")
                        Else
                            ' Override model does NOT support tooling — apply only non-tooling
                            ' fields, and disable tools for this mail.
                            _context.INI_Model_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Model), modelOverrideConfig.Model, _context.INI_Model_2)
                            _context.INI_Endpoint_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Endpoint), modelOverrideConfig.Endpoint, _context.INI_Endpoint_2)
                            _context.INI_APIKey_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKey), modelOverrideConfig.APIKey, _context.INI_APIKey_2)
                            _context.INI_APIKeyBack_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKeyBack), modelOverrideConfig.APIKeyBack, _context.INI_APIKeyBack_2)
                            _context.INI_HeaderA_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.HeaderA), modelOverrideConfig.HeaderA, _context.INI_HeaderA_2)
                            _context.INI_HeaderB_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.HeaderB), modelOverrideConfig.HeaderB, _context.INI_HeaderB_2)
                            _context.INI_Temperature_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Temperature), modelOverrideConfig.Temperature, _context.INI_Temperature_2)
                            _context.INI_MaxOutputToken_2 = If(modelOverrideConfig.MaxOutputToken <> 0, modelOverrideConfig.MaxOutputToken, _context.INI_MaxOutputToken_2)
                            _context.INI_Timeout_2 = If(modelOverrideConfig.Timeout <> 0, modelOverrideConfig.Timeout, _context.INI_Timeout_2)
                            _context.INI_APIEncrypted_2 = modelOverrideConfig.APIEncrypted
                            _context.INI_APIKeyPrefix_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APIKeyPrefix), modelOverrideConfig.APIKeyPrefix, _context.INI_APIKeyPrefix_2)
                            _context.DecodedAPI_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.DecodedAPI), modelOverrideConfig.DecodedAPI, _context.DecodedAPI_2)
                            _context.INI_Response_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.Response), modelOverrideConfig.Response, _context.INI_Response_2)
                            _context.INI_APICall_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APICall), modelOverrideConfig.APICall, _context.INI_APICall_2)
                            _context.INI_APICall_Object_2 = If(Not String.IsNullOrEmpty(modelOverrideConfig.APICall_Object), modelOverrideConfig.APICall_Object, _context.INI_APICall_Object_2)
                            ApDashboardLog($"Applied model override: {modelOverrideName} (⚠ no tooling support — running without tools for this mail)", "warn")
                        End If
                    End If

                    ' Determine whether to use tooling for this specific mail
                    Dim useToolsForThisMail As Boolean = (modelOverrideConfig Is Nothing OrElse overrideSupportsTooling)

                    If useToolsForThisMail AndAlso _apSelectedTools IsNot Nothing AndAlso _apSelectedTools.Count > 0 Then
                        response = Await ExecuteToolingLoop(
                            systemPrompt, userPrompt,
                            _apSelectedTools, _apUseSecondApi,
                            hideSplash:=True, hideLogWindow:=True,
                            cancellationToken:=ct, binaryOutputDirectory:=_apCurrentTempDir)
                    Else
                        ' Use no-tools system prompt when tooling is disabled
                        Dim effectiveSystemPrompt = If(useToolsForThisMail, systemPrompt, InterpolateAtRuntime(SP_AutoPilot_NoTools))
                        response = Await LLM(effectiveSystemPrompt, userPrompt,
                                             UseSecondAPI:=_apUseSecondApi,
                                             HideSplash:=True, EnsureUI:=False,
                                             cancellationToken:=ct,
                                             binaryOutputDirectory:=_apCurrentTempDir)
                    End If
                Finally
                    _apCurrentTempDir = Nothing
                    _apCurrentAttachments = Nothing
                    _apCurrentMailInfo = Nothing
                    MaxToolIterations = previousMaxToolIterations
                    ClearAttachmentCaches()
                End Try

                If String.IsNullOrWhiteSpace(response) Then
                    ApDashboardLog("WARNING: LLM returned empty response for: " & mailInfo.Subject, "warn")
                    Dim errorMessage = $"I'm sorry, but {AN6} was unable to generate a response to your request. Please try again or rephrase your message."
                    Await SwitchToUi(Sub() SendReplyToSender(mi, errorMessage, Nothing, tagAsAutoReply:=True))
                    Interlocked.Increment(_apSessionReplyCount)
                    RecordSenderCooldown(mailInfo.SenderEmail)
                    ApDashboardLog("Sent error response to: " & mailInfo.SenderEmail, "warn")
                    Return
                End If

                ApDashboardLog($"AI response received ({response.Length} chars).", "info")

                ' Collect result attachments from OutputFiles
                Dim resultAttachments As List(Of String) = CollectResultAttachments(tempDir, attachmentPaths)
                If resultAttachments.Count > 0 Then
                    ApDashboardLog($"Result attachments to send: {resultAttachments.Count}", "info")
                    For Each ra In resultAttachments
                        ApDashboardLog($"  📎 {Path.GetFileName(ra)}", "info")
                    Next
                End If

                ' Build "Sources used:" footer
                Dim sourcesHtml = BuildSourcesUsedHtml(_apCurrentToolCallLog)

                ' ── Approval or auto-send ──
                If requiresApproval Then
                    Await SwitchToUi(Sub() SendReplyToSender(mi, SP_AutoPilot_HoldingResponse, Nothing, tagAsAutoReply:=True))
                    ApDashboardLog("Holding response sent to: " & mailInfo.SenderEmail, "step")

                    Dim approved As Boolean = Await SwitchToUi(Function() ShowApprovalDialog(mailInfo, response, resultAttachments))
                    If approved Then
                        Await SwitchToUi(Sub() SendReplyToSender(mi, response, resultAttachments, tagAsAutoReply:=True, sourcesHtml:=sourcesHtml))
                        Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
                        Interlocked.Increment(_apSessionReplyCount)
                        RecordSenderCooldown(mailInfo.SenderEmail)
                        RecordLastProcessedTime()
                        ApDashboardLog("✓ APPROVED & SENT reply to: " & mailInfo.SenderEmail, "info")
                    Else
                        ApDashboardLog("REJECTED reply for: " & mailInfo.Subject, "step")
                    End If
                Else
                    Await SwitchToUi(Sub() SendReplyToSender(mi, response, resultAttachments, tagAsAutoReply:=True, sourcesHtml:=sourcesHtml))
                    Await SwitchToUi(Sub() TagOriginalMailAsProcessed(mi))
                    Interlocked.Increment(_apSessionReplyCount)
                    RecordSenderCooldown(mailInfo.SenderEmail)
                    RecordLastProcessedTime()
                    TrackAutoPilotConversation(mi)
                    ApDashboardLog($"✓ SENT reply to: {mailInfo.SenderEmail} (session total: {_apSessionReplyCount})", "info")
                End If
            Finally
                ' ── SECURITY: Clean up temp directory ──
                Try
                    If Directory.Exists(tempDir) Then Directory.Delete(tempDir, recursive:=True)
                Catch
                End Try
                _apCurrentToolCallLog = Nothing
            End Try

        Catch ex As OperationCanceledException
            Throw
        Catch ex As System.Exception
            ApDashboardLog("ERROR: " & ex.Message, "error")
            Debug.WriteLine("AutoPilot ProcessIncomingMailAsync error: " & ex.ToString())
        Finally
            If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' Clears all cached text data from attachment info objects for security.
    ''' Must be called after each reply is sent or when processing is abandoned.
    ''' </summary>
    Private Sub ClearAttachmentCaches()
        If _apCurrentAttachments IsNot Nothing Then
            For Each att In _apCurrentAttachments
                att.CachedText = Nothing
                att.CachedDocxHint = Nothing
            Next
        End If
    End Sub

    ''' <summary>
    ''' Tags the original incoming mail with the AP_CategoryName category so that
    ''' the catch-up scan can identify it as already processed.
    ''' </summary>
    Private Sub TagOriginalMailAsProcessed(mi As MailItem)
        Try
            Dim existing = mi.Categories
            If String.IsNullOrWhiteSpace(existing) Then
                mi.Categories = AP_CategoryName
            ElseIf existing.IndexOf(AP_CategoryName, StringComparison.OrdinalIgnoreCase) < 0 Then
                mi.Categories = existing & ", " & AP_CategoryName
            Else
                Return ' Already tagged
            End If
            mi.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] TagOriginalMailAsProcessed error: {ex.Message}")
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MAIL INFO EXTRACTION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Extracts the core data needed for AutoPilot processing.</summary>
    Private Function ExtractMailInfo(mi As MailItem) As AutoPilotMailInfo
        Try
            Dim info As New AutoPilotMailInfo()
            info.EntryID = mi.EntryID
            info.Subject = If(mi.Subject, "")
            info.SenderName = If(mi.SenderName, "")
            info.ReceivedTime = mi.ReceivedTime
            Try
                If mi.SenderEmailType = "EX" Then
                    Dim sender As AddressEntry = mi.Sender
                    If sender IsNot Nothing Then
                        Dim exUser = sender.GetExchangeUser()
                        info.SenderEmail = If(exUser?.PrimarySmtpAddress, mi.SenderEmailAddress)
                    Else
                        info.SenderEmail = mi.SenderEmailAddress
                    End If
                Else
                    info.SenderEmail = If(mi.SenderEmailAddress, "")
                End If
            Catch
                info.SenderEmail = If(mi.SenderEmailAddress, "")
            End Try
            info.Body = If(mi.Body, "")
            Try
                Dim propVal = mi.PropertyAccessor.GetProperty(AP_LoopHeaderProperty)
                info.HasAutoReplyHeader = (propVal IsNot Nothing AndAlso propVal.ToString() = AP_LoopHeaderValue)
            Catch
                info.HasAutoReplyHeader = False
            End Try
            info.ThreadAIReplyCount = 0
            Try
                Dim conv = mi.GetConversation()
                If conv IsNot Nothing Then
                    Dim table = conv.GetTable()
                    Dim count As Integer = 0
                    While Not table.EndOfTable
                        Dim row = table.GetNextRow()
                        Try
                            Dim rowCategories = row("Categories")?.ToString()
                            If rowCategories IsNot Nothing AndAlso rowCategories.Contains(AP_CategoryName) Then count += 1
                        Catch
                        End Try
                    End While
                    info.ThreadAIReplyCount = count
                End If
            Catch
            End Try
            info.AttachmentCount = mi.Attachments.Count
            Dim attachNames As New List(Of String)()
            For i As Integer = 1 To mi.Attachments.Count
                attachNames.Add(mi.Attachments(i).FileName)
            Next
            info.AttachmentNames = attachNames
            Try
                Dim parentFolder As MAPIFolder = mi.Parent
                info.FolderPath = parentFolder.FolderPath
            Catch
                info.FolderPath = ""
            End Try
            info.MessageClass = If(mi.MessageClass, "")
            info.InternetHeaders = ""
            Try
                Dim headers = mi.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F")
                info.InternetHeaders = If(headers?.ToString(), "")
            Catch
            End Try
            Return info
        Catch ex As System.Exception
            Debug.WriteLine($"ExtractMailInfo error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  FILTER ENGINE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Determines whether a mail appears to be an automatic reply or OOF.</summary>
    Private Function IsAutoReplyOrOof(info As AutoPilotMailInfo) As Boolean
        Dim subjectLower = info.Subject.ToLowerInvariant()
        For Each pattern In AP_AutoReplySubjectPatterns
            If subjectLower.StartsWith(pattern) Then Return True
        Next
        If info.MessageClass.IndexOf("OofTemplate", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If info.MessageClass.IndexOf("Report", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If info.InternetHeaders.Length > 0 Then
            Dim headersLower = info.InternetHeaders.ToLowerInvariant()
            If headersLower.Contains("auto-submitted: auto-replied") Then Return True
            If headersLower.Contains("auto-submitted: auto-generated") Then Return True
            If headersLower.Contains("auto-submitted: auto-notified") Then Return True
            If Regex.IsMatch(headersLower, "precedence:\s*(bulk|junk|list)") Then Return True
        End If
        Return False
    End Function

    ''' <summary>Checks whether the mail matches any positive filter rule.</summary>
    Private Function MatchesFilterRules(info As AutoPilotMailInfo) As Boolean
        If _apConfig.FilterRules Is Nothing OrElse _apConfig.FilterRules.Count = 0 Then Return True
        For Each rule In _apConfig.FilterRules
            If rule.IsNegative Then Continue For
            If MatchesSingleRule(info, rule) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Checks whether the mail matches any negative filter rule.</summary>
    Private Function MatchesNegativeFilters(info As AutoPilotMailInfo) As Boolean
        If _apConfig.FilterRules Is Nothing Then Return False
        For Each rule In _apConfig.FilterRules
            If Not rule.IsNegative Then Continue For
            If MatchesSingleRule(info, rule) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Evaluates a single filter rule against the mail info.</summary>
    Private Shared Function MatchesSingleRule(info As AutoPilotMailInfo, rule As AutoPilotFilterRule) As Boolean
        Select Case rule.RuleType
            Case AutoPilotFilterRuleType.Domain
                Dim pattern = rule.Pattern.TrimStart("*"c)
                If info.SenderEmail.EndsWith(pattern, StringComparison.OrdinalIgnoreCase) Then Return True
            Case AutoPilotFilterRuleType.Sender
                If WildcardMatch(info.SenderEmail, rule.Pattern) Then Return True
                If WildcardMatch(info.SenderName, rule.Pattern) Then Return True
            Case AutoPilotFilterRuleType.Folder
                If info.FolderPath.IndexOf(rule.Pattern, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        End Select
        Return False
    End Function

    ''' <summary>Performs wildcard matching using * and ? patterns.</summary>
    Private Shared Function WildcardMatch(input As String, pattern As String) As Boolean
        If String.IsNullOrEmpty(input) OrElse String.IsNullOrEmpty(pattern) Then Return False
        Dim regexPattern = "^" & Regex.Escape(pattern).Replace("\*", ".*").Replace("\?", ".") & "$"
        Return Regex.IsMatch(input, regexPattern, RegexOptions.IgnoreCase)
    End Function

    ''' <summary>Determines whether the sender is in the whitelist.</summary>
    Private Function IsSenderWhitelisted(senderEmail As String) As Boolean
        If _apConfig.WhitelistedSenders Is Nothing OrElse _apConfig.WhitelistedSenders.Count = 0 Then Return False
        For Each pattern In _apConfig.WhitelistedSenders
            If WildcardMatch(senderEmail, pattern) Then Return True
        Next
        Return False
    End Function

    ''' <summary>Checks whether the sender is within the cooldown window.</summary>
    Private Function IsSenderOnCooldown(senderEmail As String) As Boolean
        Dim lastReply As DateTime
        If _apSenderCooldowns.TryGetValue(senderEmail, lastReply) Then
            Return (DateTime.UtcNow - lastReply).TotalSeconds < _apConfig.CooldownSeconds
        End If
        Return False
    End Function

    ''' <summary>Records a cooldown entry for the sender.</summary>
    Private Sub RecordSenderCooldown(senderEmail As String)
        _apSenderCooldowns(senderEmail) = DateTime.UtcNow
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  ATTACHMENT HANDLING (save + improved collect)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Saves attachments to a per-mail temporary directory.</summary>
    Private Function SaveAttachmentsToTemp(mi As MailItem, tempDir As String) As List(Of AutoPilotAttachmentInfo)
        Dim result As New List(Of AutoPilotAttachmentInfo)()
        For i As Integer = 1 To mi.Attachments.Count
            Dim att = mi.Attachments(i)
            Try
                If att.Type = OlAttachmentType.olEmbeddeditem Then Continue For
                Dim fileName As String = att.FileName
                If String.IsNullOrWhiteSpace(fileName) Then Continue For
                Dim size As Long = att.Size
                Dim info As New AutoPilotAttachmentInfo() With {
                    .OriginalFileName = fileName,
                    .Extension = Path.GetExtension(fileName).ToLowerInvariant(),
                    .SizeBytes = size,
                    .IsOverSizeLimit = (size > _apConfig.MaxAttachmentBytes)
                }
                ' Capture original file dates from MAPI properties (preserved from the original attachment,
                ' unlike file system timestamps which reset to current time on SaveAsFile)
                Try
                    Dim pa = att.PropertyAccessor
                    Try
                        info.CreatedTime = CDate(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x30070040"))
                    Catch : End Try
                    Try
                        info.LastModifiedTime = CDate(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x30080040"))
                    Catch : End Try
                Catch
                    ' Dates are optional; ignore failures
                End Try
                If info.IsOverSizeLimit Then
                    info.TempFilePath = Nothing
                    info.StatusMessage = $"Skipped (size {size / 1024 / 1024:F1} MB exceeds limit)"
                Else
                    Dim savePath = Path.Combine(tempDir, fileName)
                    Dim counter = 1
                    While File.Exists(savePath)
                        savePath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(fileName) & $"_{counter}" & Path.GetExtension(fileName))
                        counter += 1
                    End While
                    att.SaveAsFile(savePath)
                    info.TempFilePath = savePath
                    info.StatusMessage = "Saved"
                End If
                result.Add(info)
            Catch ex As System.Exception
                result.Add(New AutoPilotAttachmentInfo() With {.OriginalFileName = att.FileName, .StatusMessage = $"Error: {ex.Message}"})
            End Try
        Next
        Return result
    End Function

    ''' <summary>
    ''' Collects output files produced by tools. Uses the OutputFiles lists on each
    ''' attachment info AND falls back to scanning for new files in tempDir.
    ''' Only files within tempDir are eligible (security: path prefix check).
    ''' </summary>
    Private Shared Function CollectResultAttachments(tempDir As String, originalAttachments As List(Of AutoPilotAttachmentInfo)) As List(Of String)
        Dim results As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim tempDirFull = Path.GetFullPath(tempDir)

        ' 1. Collect from OutputFiles (registered by tools like process_word_document, merge_pdfs)
        If originalAttachments IsNot Nothing Then
            For Each att In originalAttachments
                If att.OutputFiles IsNot Nothing Then
                    For Each outputPath In att.OutputFiles
                        If Not String.IsNullOrEmpty(outputPath) AndAlso File.Exists(outputPath) Then
                            ' Security: only include files inside the per-mail temp dir
                            If Path.GetFullPath(outputPath).StartsWith(tempDirFull, StringComparison.OrdinalIgnoreCase) Then
                                results.Add(outputPath)
                            End If
                        End If
                    Next
                End If
            Next
        End If

        ' 2. Fallback: also scan for any new files in tempDir not in the original set
        If Directory.Exists(tempDir) Then
            Dim originalPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            If originalAttachments IsNot Nothing Then
                For Each att In originalAttachments
                    If att.TempFilePath IsNot Nothing Then originalPaths.Add(att.TempFilePath)
                Next
            End If
            For Each filePath In Directory.GetFiles(tempDir)
                If Not originalPaths.Contains(filePath) Then results.Add(filePath)
            Next
        End If

        Return results.ToList()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PROMPT BUILDING
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Builds the user prompt from the mail content and attachments.</summary>
    Private Shared Function BuildUserPromptFromMail(info As AutoPilotMailInfo, attachments As List(Of AutoPilotAttachmentInfo)) As String
        Dim sb As New StringBuilder()
        sb.AppendLine("[INCOMING EMAIL]")
        sb.AppendLine($"From: {info.SenderName} <{info.SenderEmail}>")
        sb.AppendLine($"Subject: {info.Subject}")
        sb.AppendLine($"Received: {info.ReceivedTime:yyyy-MM-dd HH:mm}")
        sb.AppendLine()
        sb.AppendLine("[EMAIL BODY]")
        Dim body = info.Body
        If body.Length > 15000 Then body = body.Substring(0, 15000) & vbCrLf & "[... truncated ...]"
        sb.AppendLine(body)
        sb.AppendLine("[/EMAIL BODY]")
        If attachments IsNot Nothing AndAlso attachments.Count > 0 Then
            sb.AppendLine()
            sb.AppendLine("[ATTACHMENTS]")
            For i As Integer = 0 To attachments.Count - 1
                Dim att = attachments(i)
                Dim sizeStr = If(att.SizeBytes > 0, $" ({att.SizeBytes / 1024:F0} KB)", "")
                Dim statusStr = If(att.IsOverSizeLimit, " [OVER SIZE LIMIT - cannot process]", "")
                Dim dateStr As String = ""
                If att.LastModifiedTime.HasValue Then
                    dateStr &= $", modified: {att.LastModifiedTime.Value:yyyy-MM-dd HH:mm:ss} UTC"
                End If
                If att.CreatedTime.HasValue Then
                    dateStr &= $", created: {att.CreatedTime.Value:yyyy-MM-dd HH:mm:ss} UTC"
                End If
                sb.AppendLine($"  {i + 1}. {att.OriginalFileName}{sizeStr}{dateStr}{statusStr}")
            Next
            sb.AppendLine("[/ATTACHMENTS]")
            sb.AppendLine()
            sb.AppendLine("You have tools available to process these attachments if the sender requests it.")
        End If
        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  REPLY SENDING (sourcesHtml parameter)
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Sends a reply to the original sender, optionally with attachments and sources.</summary>
    Private Sub SendReplyToSender(originalMail As MailItem, responseText As String,
                                  resultAttachments As List(Of String),
                                  Optional tagAsAutoReply As Boolean = True,
                                  Optional sourcesHtml As String = "")
        Dim reply As MailItem = Nothing
        Try
            reply = originalMail.Reply()
            reply.CC = ""
            reply.BCC = ""
            reply.To = If(originalMail.SenderEmailAddress, originalMail.SenderName)
            reply.BodyFormat = OlBodyFormat.olFormatHTML

            Dim originalThread As String = If(reply.HTMLBody, "")

            ' Build HTML body: response + sources used footer + branded footer
            Dim htmlBody = ConvertResponseToHtml(responseText)

            ' Append "Sources used:" section
            If Not String.IsNullOrWhiteSpace(sourcesHtml) Then
                htmlBody &= sourcesHtml
            End If

            htmlBody &= BuildAutoPilotFooter()
            reply.HTMLBody = htmlBody & originalThread

            ' Add result attachments
            If resultAttachments IsNot Nothing Then
                For Each attachPath In resultAttachments
                    If File.Exists(attachPath) Then
                        reply.Attachments.Add(attachPath, OlAttachmentType.olByValue, , Path.GetFileName(attachPath))
                    End If
                Next
            End If

            If tagAsAutoReply Then
                Try : reply.PropertyAccessor.SetProperty(AP_LoopHeaderProperty, AP_LoopHeaderValue) : Catch : End Try
            End If
            Try : reply.Categories = AP_CategoryName : Catch : End Try

            reply.Send()
            Try : MoveLastSentToInkyReplies() : Catch : End Try
        Catch ex As System.Exception
            ApDashboardLog($"ERROR sending reply: {ex.Message}", "error")
        Finally
            If reply IsNot Nothing Then Try : Marshal.ReleaseComObject(reply) : Catch : End Try
        End Try
    End Sub

    ''' <summary>Builds the HTML footer appended to AutoPilot replies.</summary>
    Private Function BuildAutoPilotFooter() As String
        Dim logoBase64 As String = ""
        Try
            Dim logoBytes = GetLogoPngBytes()
            If logoBytes IsNot Nothing AndAlso logoBytes.Length > 0 Then logoBase64 = System.Convert.ToBase64String(logoBytes)
        Catch
        End Try
        Dim sb As New StringBuilder()
        sb.AppendLine("<br/><hr style='border:none;border-top:1px solid #cccccc;margin:20px 0 10px 0;'/>")
        sb.AppendLine("<table cellpadding='0' cellspacing='0' border='0' style='font-family:Arial,sans-serif;font-size:9pt;color:#888888;'>")
        sb.AppendLine("<tr>")
        If logoBase64.Length > 0 Then
            sb.AppendLine($"<td style='padding-right:8px;vertical-align:middle;'><img src='data:image/png;base64,{logoBase64}' width='24' height='24' alt='{AN}' style='display:block;'/></td>")
        End If
        sb.AppendLine($"<td style='vertical-align:middle;'>")
        Dim customFooter = _apConfig?.FooterText
        If Not String.IsNullOrWhiteSpace(customFooter) Then
            sb.AppendLine($"<span style='font-size:9pt;color:#666666;'>{System.Net.WebUtility.HtmlEncode(customFooter)}</span><br/>")
        End If
        sb.AppendLine($"<span style='font-size:8pt;color:#aaaaaa;'>Powered by <a href='https://redink.ai' style='color:#aaaaaa;text-decoration:none;'>{AN}</a></span>")
        sb.AppendLine("</td></tr></table>")
        Return sb.ToString()
    End Function

    ''' <summary>Converts a Markdown response to HTML for Outlook.</summary>
    Private Function ConvertResponseToHtml(responseMarkdown As String) As String
        Dim builder As New Markdig.MarkdownPipelineBuilder()
        builder.UsePipeTables().UseGridTables().UseSoftlineBreakAsHardlineBreak()
        builder.UseListExtras().UseFootnotes().UseDefinitionLists()
        builder.UseAbbreviations().UseAutoLinks().UseTaskLists()
        builder.UseEmojiAndSmiley().UseMathematics().UseFigures()
        builder.UseAdvancedExtensions().UseGenericAttributes()
        Dim pipeline = builder.Build()
        Dim htmlBody As String = Markdig.Markdown.ToHtml(responseMarkdown, pipeline)
        Return "<div style='font-family:Arial,sans-serif;font-size:11pt;'>" & htmlBody & "</div>"
    End Function

    ''' <summary>Moves the most recent AutoPilot reply into the Inky Replies folder.</summary>
    Private Sub MoveLastSentToInkyReplies()
        Try
            Dim ns = Application.GetNamespace("MAPI")
            Dim sentFolder = ns.GetDefaultFolder(OlDefaultFolders.olFolderSentMail)
            Dim inkyFolder As MAPIFolder = Nothing
            Try : inkyFolder = sentFolder.Folders(AP_SentSubfolder) : Catch : inkyFolder = sentFolder.Folders.Add(AP_SentSubfolder, OlDefaultFolders.olFolderSentMail) : End Try
            Dim items = sentFolder.Items
            items.Sort("[ReceivedTime]", True)
            For Each item In items
                If TypeOf item Is MailItem Then
                    Dim mi2 = DirectCast(item, MailItem)
                    If mi2.Categories IsNot Nothing AndAlso mi2.Categories.Contains(AP_CategoryName) Then
                        mi2.Move(inkyFolder) : Exit For
                    End If
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"MoveLastSentToInkyReplies error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>Shows the approval dialog and returns the user's choice.</summary>
    Private Function ShowApprovalDialog(info As AutoPilotMailInfo, draftResponse As String, resultAttachments As List(Of String)) As Boolean
        Dim attachmentInfo As String = ""
        If resultAttachments IsNot Nothing AndAlso resultAttachments.Count > 0 Then
            attachmentInfo = vbCrLf & vbCrLf & $"Attachments to include: {resultAttachments.Count}" & vbCrLf &
                String.Join(vbCrLf, resultAttachments.Select(Function(p) "  • " & Path.GetFileName(p)))
        End If
        Dim displayText As String =
            $"From: {info.SenderName} <{info.SenderEmail}>" & vbCrLf &
            $"Subject: {info.Subject}" & vbCrLf & vbCrLf &
            "── Draft Reply ──" & vbCrLf & draftResponse & attachmentInfo
        Return ShowCustomYesNoBox(displayText, "Send Reply", "Discard", header:=$"{AN6} AutoPilot — Approve Reply") = 1
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  DASHBOARD LOG — date+time, no duplicate timestamp
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Appends a log entry to the AutoPilot dashboard.
    ''' Does NOT prepend a timestamp here because LogWindow.AppendLogInternal
    ''' already adds [HH:mm:ss.fff]. Instead we prepend the date-only portion
    ''' so the dashboard shows [HH:mm:ss.fff] [dd-MMM] message.
    ''' </summary>
    Private Sub ApDashboardLog(message As String, level As String)
        Debug.WriteLine($"[AutoPilot] [{level}] {message}")
        Try
            If _apDashboard IsNot Nothing Then
                ' Prepend date portion only — LogWindow adds the time
                Dim dateTag = DateTime.Now.ToString("dd-MMM", Globalization.CultureInfo.InvariantCulture)
                Dim taggedMessage = $"[{dateTag}] {message}"
                If _apDashboard.InvokeRequired Then
                    _apDashboard.BeginInvoke(New MethodInvoker(Sub() _apDashboard.AppendLog(taggedMessage, level)))
                Else
                    _apDashboard.AppendLog(taggedMessage, level)
                End If
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Marks the dashboard operation as complete.</summary>
    Private Sub ApDashboardMarkComplete()
        Try
            If _apDashboard IsNot Nothing Then
                If _apDashboard.InvokeRequired Then
                    _apDashboard.BeginInvoke(New MethodInvoker(Sub() _apDashboard.MarkComplete()))
                Else
                    _apDashboard.MarkComplete()
                End If
            End If
        Catch
        End Try
    End Sub


    ''' <summary>
    ''' Records the current UTC time as the last successfully processed timestamp.
    ''' Uses wall-clock time (not the mail's ReceivedTime) to avoid losing mails
    ''' received at the exact same second as the last processed mail.
    ''' </summary>
    Private Sub RecordLastProcessedTime()
        Try
            My.Settings.AP_LastProcessedTime = DateTime.UtcNow
            My.Settings.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] Failed to save last processed time: {ex.Message}")
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SOURCES USED HTML FOOTER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Builds an HTML "Sources used:" section from tool call entries.
    ''' Only external tools (not internal document/attachment tools) are shown.
    ''' Web retriever URLs are rendered as clickable hyperlinks.
    ''' </summary>
    Private Shared Function BuildSourcesUsedHtml(toolCallLog As List(Of AutoPilotToolCallEntry)) As String
        If toolCallLog Is Nothing OrElse toolCallLog.Count = 0 Then Return ""

        Dim externalCalls = toolCallLog.Where(Function(e) Not e.IsInternalTool).ToList()
        If externalCalls.Count = 0 Then Return ""

        Dim sb As New StringBuilder()
        sb.AppendLine("<div style='font-size:8pt;color:#999999;margin-top:16px;border-top:1px solid #eeeeee;padding-top:8px;'>")
        sb.AppendLine("<b style='color:#888888;'>Sources used:</b><br/>")

        For Each entry In externalCalls
            Dim toolLabel = If(Not String.IsNullOrEmpty(entry.ToolDisplayName), entry.ToolDisplayName, entry.ToolName)
            Dim icon = If(entry.WasSuccessful, "✓", "✗")

            ' Check if this tool call has URLs (web retriever)
            If entry.Urls IsNot Nothing AndAlso entry.Urls.Count > 0 Then
                ' Render each URL as a clickable link
                For Each url In entry.Urls
                    Dim encodedUrl = System.Net.WebUtility.HtmlEncode(url)
                    sb.AppendLine($"&bull; {icon} <a href='{encodedUrl}' style='color:#4A90D9;text-decoration:underline;'>{encodedUrl}</a><br/>")
                Next
            Else
                ' Non-web tool: render as before with param summary
                Dim paramInfo = ""
                If Not String.IsNullOrEmpty(entry.ParamSummary) Then
                    paramInfo = $" — {System.Net.WebUtility.HtmlEncode(entry.ParamSummary)}"
                End If
                sb.AppendLine($"&bull; {icon} {System.Net.WebUtility.HtmlEncode(toolLabel)}{paramInfo}<br/>")
            End If
        Next

        sb.AppendLine("</div>")
        Return sb.ToString()
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  TOOL CALL LOGGING HELPER
    '  Called from ExecuteToolCall in ThisAddin.Tooling.vb and from
    '  TryExecuteAutoPilotTool in ThisAddIn.AutoPilot.Tools.vb
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Records a tool call for dashboard logging and the "Sources used:" footer.</summary>
    Friend Sub RecordAutoPilotToolCall(toolName As String, toolDisplayName As String,
                                       paramSummary As String, isInternalTool As Boolean,
                                       wasSuccessful As Boolean, resultExcerpt As String,
                                       elapsed As TimeSpan,
                                       Optional urls As List(Of String) = Nothing)
        If _apCurrentToolCallLog Is Nothing Then Return

        _apCurrentToolCallLog.Add(New AutoPilotToolCallEntry() With {
            .ToolName = toolName,
            .ToolDisplayName = toolDisplayName,
            .ParamSummary = paramSummary,
            .IsInternalTool = isInternalTool,
            .WasSuccessful = wasSuccessful,
            .ResultExcerpt = resultExcerpt,
            .Elapsed = elapsed,
            .Urls = urls
        })
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  Conversation Detection
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Checks whether the incoming mail is part of an existing AutoPilot conversation,
    ''' even if it is a second (or subsequent) user reply to a previous AutoPilot response.
    ''' This fixes the issue where a user replies to an already-responded-to mail and the
    ''' AutoPilot does not detect it as part of an active conversation.
    ''' </summary>
    ''' <param name="mi">The incoming MailItem to check.</param>
    ''' <returns>True if this mail is part of an AutoPilot-managed conversation.</returns>
    Private Function IsPartOfAutoPilotConversation(mi As MailItem) As Boolean
        Try
            ' 1. Check the incoming mail itself for the header
            If HasAutoReplyHeader(mi) Then Return True

            ' 2. Walk the conversation thread to find any mail with the AutoPilot header.
            '    This handles the case where the user replies to a previous AutoPilot reply:
            '    the user's new mail won't have the header, but the previous reply does.
            Dim conversation As Outlook.Conversation = Nothing
            Try
                conversation = mi.GetConversation()
            Catch
                ' GetConversation can fail on some store types
            End Try

            If conversation IsNot Nothing Then
                Try
                    Dim rootItems As SimpleItems = conversation.GetRootItems()
                    If rootItems IsNot Nothing Then
                        For Each rootItem As Object In rootItems
                            If CheckConversationNodeForAutoReply(conversation, rootItem, 0) Then
                                Return True
                            End If
                        Next
                    End If
                Catch
                    ' Conversation API can throw on certain store types
                End Try
            End If

            ' 3. Fallback: check the ConversationID against recently processed conversations.
            '    This handles edge cases where the Conversation API doesn't work but we
            '    have already processed a mail in this thread during this session.
            If Not String.IsNullOrWhiteSpace(mi.ConversationID) Then
                If _apProcessedConversations.ContainsKey(mi.ConversationID) Then
                    Return True
                End If
            End If

        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] IsPartOfAutoPilotConversation error: {ex.Message}")
        End Try

        Return False
    End Function

    ''' <summary>Recursively checks conversation tree nodes for the AutoPilot reply header.</summary>
    Private Function CheckConversationNodeForAutoReply(
            conv As Outlook.Conversation,
            item As Object,
            depth As Integer) As Boolean

        If depth > AP_MaxThreadDepth Then Return False

        Try
            Dim nodeMail = TryCast(item, MailItem)
            If nodeMail IsNot Nothing Then
                If HasAutoReplyHeader(nodeMail) Then Return True
            End If

            ' Check child nodes
            Dim children As SimpleItems = conv.GetChildren(item)
            If children IsNot Nothing Then
                For Each child As Object In children
                    If CheckConversationNodeForAutoReply(conv, child, depth + 1) Then
                        Return True
                    End If
                Next
            End If
        Catch
            ' Ignore individual node errors
        End Try

        Return False
    End Function

    ''' <summary>Checks whether a mail item has the X-RedInk-AutoReply MAPI property set.</summary>
    Private Function HasAutoReplyHeader(mi As MailItem) As Boolean
        Try
            Dim prop As Object = Nothing
            Try
                Dim pa As PropertyAccessor = mi.PropertyAccessor
                prop = pa.GetProperty(AP_LoopHeaderProperty)
            Catch
                ' Property not set — this is normal for user-sent mails
                Return False
            End Try
            If prop IsNot Nothing AndAlso CStr(prop).Equals(AP_LoopHeaderValue, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Catch
            ' Ignore errors
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Tracks conversation IDs that AutoPilot has already processed during this session.
    ''' Used as a fallback when the Conversation API is unavailable.
    ''' Key = ConversationID, Value = UTC timestamp of last processing.
    ''' </summary>
    Private ReadOnly _apProcessedConversations As New ConcurrentDictionary(Of String, DateTime)()

    ''' <summary>Records that a conversation has been processed by AutoPilot.</summary>
    Private Sub TrackAutoPilotConversation(mi As MailItem)
        Try
            If Not String.IsNullOrWhiteSpace(mi.ConversationID) Then
                _apProcessedConversations(mi.ConversationID) = DateTime.UtcNow
            End If
        Catch
            ' Ignore
        End Try
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  MODEL OVERRIDE COMMAND
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Scans the first few lines of the latest e-mail body for "#model: ModelName".
    ''' Only the latest/top-most portion is scanned (not quoted replies).
    ''' The command must appear in the first AP_ModelCommandScanLines non-blank lines.
    ''' </summary>
    Private Sub TryParseModelOverride(body As String,
                                      ByRef modelConfig As ModelConfig,
                                      ByRef modelName As String,
                                      ByRef strippedBody As String)
        modelConfig = Nothing
        modelName = Nothing
        strippedBody = body

        If String.IsNullOrWhiteSpace(body) Then Return

        ' Use the existing helper from ThisAddIn.Commands.vb which handles
        ' more separator patterns (From:, On ... wrote:, etc.)
        Dim latestBody = GetLatestMailBody(body)
        Dim lines = latestBody.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
        Dim nonBlankCount = 0
        Dim commandLineIndex = -1

        For i As Integer = 0 To lines.Length - 1
            Dim line = lines(i).Trim()
            If String.IsNullOrWhiteSpace(line) Then Continue For

            nonBlankCount += 1
            If nonBlankCount > AP_ModelCommandScanLines Then Exit For

            If line.StartsWith(AP_ModelCommandPrefix, StringComparison.OrdinalIgnoreCase) Then
                Dim requestedModel = line.Substring(AP_ModelCommandPrefix.Length).Trim()
                If String.IsNullOrWhiteSpace(requestedModel) Then Continue For

                ' Try to find a matching model from available alternates
                Try
                    ' Use INI_AlternateModelPath (same as the config dialog and model selector),
                    ' NOT INI_LocalModelPath (which is for speech/local models).
                    Dim iniPath = INI_AlternateModelPath
                    If String.IsNullOrWhiteSpace(iniPath) Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — no alternate model path configured.", "warn")
                        Continue For
                    End If

                    iniPath = ExpandEnvironmentVariables(iniPath)
                    If Not IO.File.Exists(iniPath) Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — alternate model file not found: {iniPath}", "warn")
                        Continue For
                    End If

                    Dim alternates = LoadAlternativeModels(iniPath, _context)
                    If alternates Is Nothing OrElse alternates.Count = 0 Then
                        ApDashboardLog($"Model override '{requestedModel}' ignored — no models loaded from INI.", "warn")
                        Continue For
                    End If

                    ApDashboardLog($"Resolving model override '{requestedModel}' against {alternates.Count} available model(s)...", "step")

                    ' Match by description (friendly name) first, then by model ID
                    Dim match = alternates.FirstOrDefault(
                        Function(m) m.ModelDescription IsNot Nothing AndAlso
                                    m.ModelDescription.IndexOf(requestedModel, StringComparison.OrdinalIgnoreCase) >= 0)
                    If match Is Nothing Then
                        match = alternates.FirstOrDefault(
                            Function(m) m.Model IsNot Nothing AndAlso
                                        m.Model.IndexOf(requestedModel, StringComparison.OrdinalIgnoreCase) >= 0)
                    End If

                    ' Fallback: keyword matching — all words in the request must appear
                    ' somewhere in ModelDescription (order-independent)
                    If match Is Nothing Then
                        Dim keywords = requestedModel.Split({" "c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                        If keywords.Length > 1 Then
                            match = alternates.FirstOrDefault(
                                Function(m) m.ModelDescription IsNot Nothing AndAlso
                                            keywords.All(Function(kw) m.ModelDescription.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                        End If
                        If match Is Nothing AndAlso keywords.Length > 1 Then
                            match = alternates.FirstOrDefault(
                                Function(m) m.Model IsNot Nothing AndAlso
                                            keywords.All(Function(kw) m.Model.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0))
                        End If
                    End If

                    If match IsNot Nothing Then
                        modelConfig = match
                        modelName = If(match.ModelDescription, match.Model)
                        commandLineIndex = i
                        ApDashboardLog($"✓ Model override resolved: '{requestedModel}' → {modelName}", "info")
                    Else
                        ApDashboardLog($"Model override '{requestedModel}' not found in {alternates.Count} available model(s).", "warn")
                        ' Log available model names to help the user
                        For Each alt In alternates
                            Dim altLabel = If(Not String.IsNullOrWhiteSpace(alt.ModelDescription), alt.ModelDescription, alt.Model)
                            ApDashboardLog($"  available: {altLabel}", "step")
                        Next
                    End If
                Catch ex As System.Exception
                    ApDashboardLog($"Error resolving model override: {ex.Message}", "warn")
                End Try

                Exit For ' Only process the first #model: found
            End If
        Next

        ' Strip the #model: line from the body
        If commandLineIndex >= 0 Then
            Dim sbStripped As New StringBuilder()
            For j As Integer = 0 To lines.Length - 1
                If j <> commandLineIndex Then sbStripped.AppendLine(lines(j))
            Next
            ' Reconstruct: stripped latest portion + any remaining quoted thread
            If latestBody.Length < body.Length Then
                strippedBody = sbStripped.ToString().TrimEnd() & body.Substring(latestBody.Length)
            Else
                strippedBody = sbStripped.ToString().TrimEnd()
            End If
        End If
    End Sub


    ' ═══════════════════════════════════════════════════════════════════════════
    '  DATA CLASSES
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Represents extracted mail metadata used by AutoPilot.</summary>
    Private Class AutoPilotMailInfo
        Public Property EntryID As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Body As String
        Public Property ReceivedTime As DateTime
        Public Property HasAutoReplyHeader As Boolean
        Public Property ThreadAIReplyCount As Integer
        Public Property AttachmentCount As Integer
        Public Property AttachmentNames As List(Of String)
        Public Property FolderPath As String
        Public Property MessageClass As String
        Public Property InternetHeaders As String
    End Class

    ''' <summary>Describes an attachment and its per-mail processing data.</summary>
    Friend Class AutoPilotAttachmentInfo
        Public Property OriginalFileName As String
        Public Property TempFilePath As String
        Public Property Extension As String
        Public Property SizeBytes As Long
        Public Property IsOverSizeLimit As Boolean
        Public Property StatusMessage As String
        Public Property CreatedTime As DateTime?
        Public Property LastModifiedTime As DateTime?
        Public Property OutputFiles As New List(Of String)()
        Property CachedText As String
        Property CachedDocxHint As String
        Public Property IsToolOutput As Boolean = False
    End Class

    ''' <summary>Defines supported filter rule types.</summary>
    Public Enum AutoPilotFilterRuleType
        Domain
        Sender
        Folder
    End Enum

    ''' <summary>Represents a single sender or folder filter rule.</summary>
    Public Class AutoPilotFilterRule
        Public Property RuleType As AutoPilotFilterRuleType
        Public Property Pattern As String
        Public Property IsNegative As Boolean
    End Class

    ''' <summary>Tracks a single tool call for logging and the "Sources used:" footer.</summary>
    Friend Class AutoPilotToolCallEntry
        Public Property ToolName As String
        Public Property ToolDisplayName As String
        Public Property ParamSummary As String
        Public Property IsInternalTool As Boolean
        Public Property WasSuccessful As Boolean
        Public Property ResultExcerpt As String
        Public Property Elapsed As TimeSpan

        Public Property Urls As List(Of String)
    End Class

End Class