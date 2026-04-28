' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365SearchForm.vb
' Purpose: Lightweight harness for SharedLibrary.M365Service. Lets the user
'          enter a query, lists matching emails, and opens the selected
'          message inside Outlook itself (or, as a last resort, in the
'          browser via webLink).
'
' Wire-up: From any ribbon button (or Immediate window), call:
'              M365SearchTest.Show(_context)
' =============================================================================

Option Explicit On
Option Strict On

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext
Imports Outlook = Microsoft.Office.Interop.Outlook

''' <summary>
''' Modeless Red-Ink-styled M365 mail search window with open-in-Outlook.
''' </summary>
Public Class M365SearchTestForm
    Inherits Form
    Implements IProgress(Of M365SearchProgress)

    Private ReadOnly _context As ISharedContext
    Private _hits As New List(Of M365SearchHit)()
    Private _cts As CancellationTokenSource

    ' ── UI fields ───────────────────────────────────────────────────────
    Private WithEvents txtQuery As TextBox
    Private WithEvents btnSearch As Button
    Private WithEvents btnOpen As Button
    Private WithEvents btnSignIn As Button
    Private WithEvents btnSignOut As Button
    Private WithEvents numMax As NumericUpDown
    Private WithEvents lvResults As ListView
    Private WithEvents btnClose As Button
    Private headerPanel As Panel
    Private footerPanel As Panel
    Private splitMain As SplitContainer
    Private lblQuery As Label
    Private lblStatus As Label
    Private lblUser As Label
    Private lblMax As Label
    Private WithEvents btnAISearch As Button
    Private txtSummary As TextBox
    Private lblSummary As Label
    Private pb As ProgressBar

    ' Standard Red Ink button padding/margin (matches ShowSelectionForm etc.).
    Private Shared ReadOnly StdButtonPadding As New Padding(8, 4, 5, 4)
    Private Shared ReadOnly StdButtonMargin As New Padding(0, 0, 10, 0)

    Public Sub New(context As ISharedContext)
        If context Is Nothing Then Throw New ArgumentNullException(NameOf(context))
        _context = context
        BuildUi()
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  UI construction
    ' ════════════════════════════════════════════════════════════════════
    Private Sub BuildUi()
        Me.Text = "Search Mails"
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.MinimumSize = New Size(820, 540)
        Me.ClientSize = New Size(1060, 660)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

        ' Red Ink logo as window icon.
        Try
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        ' ── Header (search bar + user label) ─────────────────────────────
        headerPanel = New Panel() With {
            .Dock = DockStyle.Top,
            .Height = 100
        }

        lblQuery = New Label() With {
            .Text = "Search query:", .AutoSize = True,
            .Location = New Point(12, 16)}
        headerPanel.Controls.Add(lblQuery)

        txtQuery = New TextBox() With {
            .Location = New Point(95, 12),
            .Size = New Size(380, 23),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left}
        headerPanel.Controls.Add(txtQuery)

        lblMax = New Label() With {
            .Text = "Max:", .AutoSize = True,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(485, 16)}
        headerPanel.Controls.Add(lblMax)

        numMax = New NumericUpDown() With {
            .Minimum = 1, .Maximum = 500, .Value = 25,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(520, 12), .Size = New Size(60, 23)}
        headerPanel.Controls.Add(numMax)

        btnSearch = New Button() With {
            .Text = "Search",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(590, 9)}
        headerPanel.Controls.Add(btnSearch)
        Me.AcceptButton = btnSearch

        btnAISearch = New Button() With {
            .Text = "AI Search",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(660, 9)}
        headerPanel.Controls.Add(btnAISearch)

        btnSignIn = New Button() With {
            .Text = "Sign in",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = StdButtonMargin,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(730, 9)}
        headerPanel.Controls.Add(btnSignIn)

        btnSignOut = New Button() With {
            .Text = "Sign out",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(800, 9)}
        headerPanel.Controls.Add(btnSignOut)

        ' "(not signed in)" — 20 px padding above and below.
        lblUser = New Label() With {
            .Text = "(not signed in)", .AutoSize = True, .ForeColor = Color.DimGray,
            .Location = New Point(12, 60)}
        headerPanel.Controls.Add(lblUser)

        ' ── Footer (progress, status, Open + Close) ──────────────────────
        footerPanel = New Panel() With {
            .Dock = DockStyle.Bottom,
            .Height = 80
        }

        pb = New ProgressBar() With {
            .Location = New Point(12, 10),
            .Size = New Size(footerPanel.ClientSize.Width - 260, 18),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right,
            .Style = ProgressBarStyle.Continuous}
        footerPanel.Controls.Add(pb)

        lblStatus = New Label() With {
            .Text = "Ready.",
            .AutoEllipsis = True,
            .Location = New Point(12, 36),
            .Size = New Size(footerPanel.ClientSize.Width - 260, 22),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right}
        footerPanel.Controls.Add(lblStatus)

        btnOpen = New Button() With {
            .Text = "Open in Outlook",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(footerPanel.ClientSize.Width - 160, 8),
            .Enabled = False}
        footerPanel.Controls.Add(btnOpen)

        btnClose = New Button() With {
            .Text = "Close",
            .AutoSize = True,
            .Padding = StdButtonPadding,
            .Margin = New Padding(0),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .Location = New Point(footerPanel.ClientSize.Width - 70, 8)}
        footerPanel.Controls.Add(btnClose)
        Me.CancelButton = btnClose

        ' ── Center (results + resizable AI summary) ──────────────────────
        splitMain = New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .SplitterWidth = 6,
            .Panel1MinSize = 140,
            .Panel2MinSize = 90,
            .FixedPanel = FixedPanel.None
        }
        splitMain.Panel1.Padding = New Padding(12, 6, 12, 2)
        splitMain.Panel2.Padding = New Padding(12, 4, 12, 6)

        lvResults = New ListView() With {
            .Dock = DockStyle.Fill,
            .View = View.Details,
            .FullRowSelect = True,
            .GridLines = True,
            .HideSelection = False,
            .MultiSelect = False}
        lvResults.Columns.Add("Date", 95)
        lvResults.Columns.Add("Time", 55)
        lvResults.Columns.Add("From", 200)
        lvResults.Columns.Add("To", 200)
        lvResults.Columns.Add("Subject", 470)
        splitMain.Panel1.Controls.Add(lvResults)

        ' Order: textbox first (Fill), label second (Top), so the label
        ' docks above the textbox and never overlaps it.
        txtSummary = New TextBox() With {
            .Multiline = True,
            .ReadOnly = True,
            .ScrollBars = ScrollBars.Vertical,
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .WordWrap = True}
        splitMain.Panel2.Controls.Add(txtSummary)

        lblSummary = New Label() With {
            .Text = "AI Summary:",
            .Dock = DockStyle.Top,
            .Height = 28,
            .Padding = New Padding(0, 6, 0, 6),
            .TextAlign = ContentAlignment.MiddleLeft}
        splitMain.Panel2.Controls.Add(lblSummary)

        ' Add docked containers in the order Top → Bottom → Fill so each
        ' takes its slice and Fill consumes what's left.
        Me.Controls.Add(splitMain)
        Me.Controls.Add(footerPanel)
        Me.Controls.Add(headerPanel)

        ' Initial split: list ~60%, summary ~40% so the summary is tall
        ' enough to read several paragraphs without scrolling.
        Try
            Dim h As Integer = Math.Max(splitMain.Panel1MinSize + splitMain.Panel2MinSize + splitMain.SplitterWidth + 20,
                                        Me.ClientSize.Height - headerPanel.Height - footerPanel.Height)
            splitMain.SplitterDistance = CInt(h * 0.58)
        Catch
        End Try

        ' Pack the right-aligned button cluster + footer buttons after
        ' the controls have their preferred sizes.
        LayoutTopRowRight()
        LayoutFooterRight()
    End Sub

    ''' <summary>
    ''' Re-positions the top-row right-aligned controls (Search / AI Search
    ''' / Sign in / Sign out plus Max label + NumericUpDown) using each
    ''' button's actual PreferredSize, so AutoSize + Padding never causes
    ''' overlap.
    ''' </summary>
    Private Sub LayoutTopRowRight()
        If headerPanel Is Nothing OrElse btnSearch Is Nothing OrElse
           btnAISearch Is Nothing OrElse btnSignIn Is Nothing OrElse
           btnSignOut Is Nothing Then Return

        Const rightMargin As Integer = 12
        Const buttonGap As Integer = 6
        Const groupGap As Integer = 12
        Dim topY As Integer = 9
        Dim labelY As Integer = 16
        Dim numY As Integer = 12

        Dim x As Integer = headerPanel.ClientSize.Width - rightMargin

        For Each b As Button In New Button() {btnSignOut, btnSignIn, btnAISearch, btnSearch}
            Dim w As Integer = b.PreferredSize.Width
            x -= w
            b.SetBounds(x, topY, w, b.Height)
            x -= buttonGap
        Next

        x -= (groupGap - buttonGap)
        x -= numMax.Width
        numMax.SetBounds(x, numY, numMax.Width, numMax.Height)

        Dim lblW As Integer = lblMax.PreferredSize.Width
        x -= 4
        x -= lblW
        lblMax.SetBounds(x, labelY, lblW, lblMax.Height)

        ' Place txtQuery just to the right of lblQuery and stretch it so
        ' it stops a fixed gap before lblMax, regardless of font/DPI.
        Const labelTextGap As Integer = 8
        Dim queryLeft As Integer = lblQuery.Left + lblQuery.PreferredSize.Width + labelTextGap
        Dim queryRight As Integer = x - groupGap
        Dim newWidth As Integer = Math.Max(120, queryRight - queryLeft)
        txtQuery.SetBounds(queryLeft, txtQuery.Top, newWidth, txtQuery.Height)

    End Sub

    ''' <summary>Right-aligns Open + Close in the footer using PreferredSize.</summary>
    Private Sub LayoutFooterRight()
        If footerPanel Is Nothing OrElse btnOpen Is Nothing OrElse btnClose Is Nothing Then Return

        Const rightMargin As Integer = 12
        Const buttonGap As Integer = 6
        Dim topY As Integer = 8

        Dim x As Integer = footerPanel.ClientSize.Width - rightMargin

        Dim wClose = btnClose.PreferredSize.Width
        x -= wClose
        btnClose.SetBounds(x, topY, wClose, btnClose.Height)
        x -= buttonGap

        Dim wOpen = btnOpen.PreferredSize.Width
        x -= wOpen
        btnOpen.SetBounds(x, topY, wOpen, btnOpen.Height)

        ' Resize progress + status to leave room for the button cluster.
        Dim leftAreaWidth As Integer = x - 12 - 8
        If leftAreaWidth > 80 Then
            pb.Size = New Size(leftAreaWidth, pb.Height)
            lblStatus.Size = New Size(leftAreaWidth, lblStatus.Height)
        End If
    End Sub


    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        LayoutTopRowRight()
        LayoutFooterRight()
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  UI-thread marshalling helpers
    ' ════════════════════════════════════════════════════════════════════
    ''' <summary>
    ''' Runs <paramref name="action"/> on the form's UI thread. Safe to call
    ''' from any thread; no-ops if the form is gone. Use this for every UI
    ''' touch that follows an Await — Outlook's main thread does not have a
    ''' WindowsFormsSynchronizationContext, so ConfigureAwait(True) may
    ''' resume on a thread-pool thread.
    ''' </summary>
    Private Sub UiPost(action As Action)
        If action Is Nothing Then Return
        If Not IsFormUsable() Then Return
        Try
            If Me.InvokeRequired Then
                Me.BeginInvoke(action)
            Else
                action()
            End If
        Catch
            ' Form vanished between the check and the invoke — ignore.
        End Try
    End Sub

    ''' <summary>Synchronous variant of <see cref="UiPost"/>.</summary>
    Private Sub UiInvoke(action As Action)
        If action Is Nothing Then Return
        If Not IsFormUsable() Then Return
        Try
            If Me.InvokeRequired Then
                Me.Invoke(action)
            Else
                action()
            End If
        Catch
        End Try
    End Sub


    ' ════════════════════════════════════════════════════════════════════
    '  Lifecycle
    ' ════════════════════════════════════════════════════════════════════
    Protected Overrides Async Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        Try
            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = If(String.IsNullOrEmpty(user),
                                         "(not signed in)",
                                         "Signed in as: " & user)
                   End Sub)
        Catch
        End Try
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Try : _cts?.Cancel() : Catch : End Try
        MyBase.OnFormClosing(e)
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  Sign in / out
    ' ════════════════════════════════════════════════════════════════════


    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Try : _cts?.Cancel() : Catch : End Try
        Me.Close()
    End Sub


    Private Async Sub btnSignIn_Click(sender As Object, e As EventArgs) Handles btnSignIn.Click
        SetBusy(True, "Signing in…")
        Try
            Await M365Service.SignInAsync(_context).ConfigureAwait(False)
            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = "Signed in as: " & user
                       lblStatus.Text = "Signed in."
                   End Sub)
        Catch ex As Exception When IsCancellation(ex)
            ' User aborted the auth dialog — treat as benign no-op.
            UiPost(Sub() lblStatus.Text = "Sign-in cancelled.")
        Catch ex As Exception
            Dim msg As String = ex.Message
            UiPost(Sub()
                       SharedMethods.ShowCustomMessageBox("Sign-in failed: " & msg, "Microsoft 365")
                       lblStatus.Text = "Sign-in failed."
                   End Sub)
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    Private Async Sub btnSignOut_Click(sender As Object, e As EventArgs) Handles btnSignOut.Click
        SetBusy(True, "Signing out…")
        Try
            Await M365Service.SignOutAsync(_context).ConfigureAwait(False)
            UiPost(Sub()
                       lblUser.Text = "(not signed in)"
                       lblStatus.Text = "Signed out."
                   End Sub)
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Sign-out cancelled.")
        Catch ex As Exception
            Dim msg As String = ex.Message
            UiPost(Sub() lblStatus.Text = "Sign-out failed: " & msg)
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub


    ''' <summary>True if the form is alive and safe to touch from a continuation.</summary>
    Private Function IsFormUsable() As Boolean
        Return Not Me.IsDisposed AndAlso Not Me.Disposing AndAlso Me.IsHandleCreated
    End Function

    ''' <summary>
    ''' Returns True for any flavour of "user cancelled" — covers
    ''' OperationCanceledException, MSAL's authentication_canceled, and
    ''' WebView2/WAM cancel paths that surface as plain Exceptions.
    ''' </summary>
    Private Shared Function IsCancellation(ex As Exception) As Boolean
        If ex Is Nothing Then Return False
        If TypeOf ex Is OperationCanceledException Then Return True

        Dim t As Type = ex.GetType()
        Dim typeName As String = If(t?.FullName, "")
        If typeName.IndexOf("Msal", StringComparison.OrdinalIgnoreCase) >= 0 Then
            Try
                Dim codeProp = t.GetProperty("ErrorCode")
                If codeProp IsNot Nothing Then
                    Dim code As String = TryCast(codeProp.GetValue(ex), String)
                    If Not String.IsNullOrEmpty(code) AndAlso
                       (code.IndexOf("cancel", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                        code.Equals("authentication_canceled", StringComparison.OrdinalIgnoreCase)) Then
                        Return True
                    End If
                End If
            Catch
            End Try
        End If

        Dim msg As String = If(ex.Message, "")
        Return msg.IndexOf("cancel", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               msg.IndexOf("user_cancel", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' ════════════════════════════════════════════════════════════════════
    '  Search
    ' ════════════════════════════════════════════════════════════════════
    Private Async Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim q = txtQuery.Text.Trim()
        If String.IsNullOrEmpty(q) Then
            SharedMethods.ShowCustomMessageBox("Enter a search term first.", "Microsoft 365")
            Return
        End If

        _cts?.Cancel()
        _cts = New CancellationTokenSource()

        SetBusy(True, "Searching…")
        lvResults.Items.Clear()
        _hits.Clear()
        btnOpen.Enabled = False

        Try
            Dim opts As New M365SearchOptions() With {
                .MaxPerSource = CInt(numMax.Value),
                .Parallel = True}
            Dim res = Await M365Service.SearchAsync(
                _context, q, M365SearchSources.Mail,
                opts, Me, _cts.Token).ConfigureAwait(False)

            Dim user = Await M365Service.GetSignedInUserAsync(_context).ConfigureAwait(False)

            UiPost(Sub()
                       _hits = res.Hits
                       PopulateResults(_hits)
                       If Not String.IsNullOrEmpty(user) Then lblUser.Text = "Signed in as: " & user
                       lblStatus.Text = $"{_hits.Count} hit(s)." &
                           If(res.HasErrors, " Errors: " & String.Join("; ", res.ErrorsBySource.Values), "")
                   End Sub)
        Catch ex As OperationCanceledException
            UiPost(Sub() lblStatus.Text = "Cancelled.")
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "Cancelled.")
        Catch ex As Exception
            Dim msg As String = ex.Message
            UiPost(Sub()
                       lblStatus.Text = "Search failed: " & msg
                       SharedMethods.ShowCustomMessageBox("Microsoft 365 search failed: " & msg, "Microsoft 365")
                   End Sub)
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    Private Async Sub btnAISearch_Click(sender As Object, e As EventArgs) Handles btnAISearch.Click
        ' 1) Capture the user's free-form prompt.
        Dim userPrompt As String = ""
        Try
            userPrompt = SharedMethods.ShowCustomInputBox(
                "Describe the e-mails you are looking for. The AI will search " &
                "your Microsoft 365 mailbox, list the relevant messages and " &
                "summarise them.",
                "AI Mail Search",
                SimpleInput:=False)
        Catch
        End Try
        If String.IsNullOrWhiteSpace(userPrompt) OrElse userPrompt = "ESC" Then Return

        _cts?.Cancel()
        _cts = New CancellationTokenSource()
        SetBusy(True, "AI is searching your mailbox…")
        lvResults.Items.Clear()
        _hits.Clear()
        btnOpen.Enabled = False
        txtSummary.Text = ""

        Dim addIn = Globals.ThisAddIn
        Dim chatScope As IDisposable = Nothing
        Dim previousOtherPrompt As String = addIn.OtherPrompt

        Try
            addIn.OtherPrompt = userPrompt
            ' 2) Make {OtherPrompt} resolve to the user's text inside SP_AIMailSearch.
            Globals.ThisAddIn.OtherPrompt = userPrompt

            ' 3) Pick only the M365 mail-search tool from the loaded tooling services.
            Dim allTools = addIn.LoadToolingServices(_context.INI_SpecialServicePath, toolsOnly:=True)

            Dim mailTool As ModelConfig = allTools.FirstOrDefault(
            Function(t) Not String.IsNullOrEmpty(t?.ToolName) AndAlso
                        SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(t.ToolName) AndAlso
                        t.ToolName.IndexOf("mail", StringComparison.OrdinalIgnoreCase) >= 0)

            If mailTool Is Nothing Then
                ' Fall back: any single M365 tool, in case the naming convention
                ' differs (e.g. m365_search, m365_messages, etc.).
                mailTool = allTools.FirstOrDefault(
                Function(t) Not String.IsNullOrEmpty(t?.ToolName) AndAlso
                            SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(t.ToolName))
            End If

            If mailTool Is Nothing Then
                SharedMethods.ShowCustomMessageBox(
                "The Microsoft 365 mail search tool is not available. Please ensure it is enabled and configured.",
                "AI Mail Search")
                Return
            End If
            Dim selectedTools As New List(Of ModelConfig) From {mailTool}

            ' 4) Allow internal M365 tool dispatch for the duration of this call.
            chatScope = addIn.EnterChatAgentScope()

            ' 5) Run the tool-enabled LLM loop. SP_AIMailSearch is the system
            '    prompt; {OtherPrompt} inside it is auto-interpolated.
            '    Do NOT force hideLogWindow — the tooling log window already
            '    honors the user's INI setting (INI_APIDebug / log toggles).
            Dim finalText = Await addIn.ExecuteToolingLoop(
                sysCommand:=_context.SP_AIMailSearch,
                userText:="",
                selectedTools:=selectedTools,
                useSecondAPI:=False,
                otherPrompt:=userPrompt,
                hideSplash:=True,
                cancellationToken:=_cts.Token).ConfigureAwait(False)

            If Not IsFormUsable() Then Return

            ' 6) Parse <summary> and <email_ids> from the model output.
            Dim summary As String = ExtractTaggedBlock(finalText, "summary")
            Dim ids = ExtractEmailIds(finalText)

            ' 7) Resolve IDs to displayable hits via Graph (so the listview is
            '    correct even if the model abbreviated the tool result).
            Dim resolved As New List(Of M365SearchHit)()
            For Each id In ids
                If Not IsFormUsable() Then Return
                _cts.Token.ThrowIfCancellationRequested()
                Try
                    Dim full = Await M365Service.GetMessageAsync(
                        _context, id,
                        M365MessageFields.Headers).ConfigureAwait(False)
                    If full IsNot Nothing Then resolved.Add(MailToHit(full))
                Catch ex As OperationCanceledException
                    Throw
                Catch ex As Exception
                    Debug.WriteLine("[AISearch] Resolve failed for " & id & ": " & ex.Message)
                End Try
            Next

            UiPost(Sub()
                       _hits = resolved
                       PopulateResults(_hits)
                       txtSummary.Text = If(String.IsNullOrWhiteSpace(summary),
                                            "(no summary returned)",
                                            summary.Trim())
                       lblStatus.Text = $"AI search: {_hits.Count} mail(s) selected by the model."
                   End Sub)

        Catch ex As OperationCanceledException
            UiPost(Sub() lblStatus.Text = "AI search cancelled.")
        Catch ex As Exception When IsCancellation(ex)
            UiPost(Sub() lblStatus.Text = "AI search cancelled.")
        Catch ex As Exception
            Dim msg As String = ex.Message
            UiPost(Sub()
                       lblStatus.Text = "AI search failed: " & msg
                       SharedMethods.ShowCustomMessageBox("AI mail search failed: " & msg, "AI Mail Search")
                   End Sub)
        Finally
            Try : chatScope?.Dispose() : Catch : End Try
            Try : addIn.OtherPrompt = previousOtherPrompt : Catch : End Try
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    Private Shared Function ExtractTaggedBlock(text As String, tag As String) As String
        If String.IsNullOrEmpty(text) Then Return ""
        Dim pattern As String = "<" & tag & ">(?<v>[\s\S]*?)</" & tag & ">"
        Dim m = System.Text.RegularExpressions.Regex.Match(
            text, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups("v").Value
        Return ""
    End Function

    Private Shared Function ExtractEmailIds(text As String) As List(Of String)
        Dim block = ExtractTaggedBlock(text, "email_ids")
        Dim list As New List(Of String)()
        If String.IsNullOrWhiteSpace(block) Then Return list
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each line In block.Split({ControlChars.Lf, ControlChars.Cr}, StringSplitOptions.RemoveEmptyEntries)
            Dim id = line.Trim().TrimStart("-"c, "*"c, " "c).Trim()
            If id.Length > 0 AndAlso seen.Add(id) Then list.Add(id)
        Next
        Return list
    End Function

    ''' <summary>Adapts a fetched M365 message to the existing M365SearchHit shape used by the listview.</summary>
    Private Shared Function MailToHit(m As Object) As M365SearchHit
        Dim h As New M365SearchHit()
        Try
            Dim t = m.GetType()
            h.Id = CStr(t.GetProperty("Id")?.GetValue(m))
            h.Title = CStr(t.GetProperty("Subject")?.GetValue(m))
            h.Author = CStr(t.GetProperty("From")?.GetValue(m))
            Dim dt = t.GetProperty("ReceivedDateTimeUtc")?.GetValue(m)
            If TypeOf dt Is DateTime Then h.LastModifiedUtc = CType(dt, DateTime)
            Dim web = TryCast(t.GetProperty("WebLink")?.GetValue(m), String)
            If Not String.IsNullOrEmpty(web) Then h.WebUrl = web
        Catch
        End Try
        Return h
    End Function

    Private Sub PopulateResults(hits As List(Of M365SearchHit))
        lvResults.BeginUpdate()
        Try
            Dim ordered = hits.
                OrderByDescending(Function(h) GetMailDate(h)).
                ToList()
            _hits = ordered
            For i = 0 To ordered.Count - 1
                Dim h = ordered(i)
                Dim dtUtc As DateTime? = GetMailDate(h)
                Dim datePart As String = ""
                Dim timePart As String = ""
                If dtUtc.HasValue Then
                    Dim dtLocal = dtUtc.Value.ToLocalTime()
                    datePart = dtLocal.ToString("yyyy-MM-dd")
                    timePart = dtLocal.ToString("HH:mm")
                End If
                Dim it As New ListViewItem(datePart)
                it.SubItems.Add(timePart)
                it.SubItems.Add(If(h.Author, ""))
                it.SubItems.Add(GetMailTo(h))
                it.SubItems.Add(If(h.Title, "(no subject)"))
                it.Tag = i
                lvResults.Items.Add(it)
            Next
        Finally
            lvResults.EndUpdate()
        End Try
    End Sub

    ''' <summary>
    ''' Builds a compact "To" string from the raw Graph payload's
    ''' toRecipients[].emailAddress.{name|address}. Returns "" if absent.
    ''' </summary>
    Private Function GetMailTo(h As M365SearchHit) As String
        Try
            Dim r = TryCast(h?.RawJson?("resource"), JObject)
            If r Is Nothing Then Return ""
            Dim arr = TryCast(r("toRecipients"), JArray)
            If arr Is Nothing OrElse arr.Count = 0 Then Return ""

            Dim names As New List(Of String)()
            For Each tok In arr
                Dim ea = TryCast(tok("emailAddress"), JObject)
                If ea Is Nothing Then Continue For
                Dim nm As String = If(ea("name")?.ToString(), "")
                Dim ad As String = If(ea("address")?.ToString(), "")
                Dim disp As String = If(Not String.IsNullOrWhiteSpace(nm), nm, ad)
                If Not String.IsNullOrWhiteSpace(disp) Then names.Add(disp.Trim())
            Next
            Return String.Join("; ", names)
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>Prefers SentDateTime from the raw Graph payload, falls back to LastModifiedUtc.</summary>
    Private Function GetMailDate(h As M365SearchHit) As DateTime?
        Try
            Dim r = TryCast(h?.RawJson?("resource"), JObject)
            If r IsNot Nothing Then
                Dim sentTok = r("sentDateTime")
                If sentTok IsNot Nothing AndAlso sentTok.Type <> JTokenType.Null Then
                    Dim dt As DateTime
                    If DateTime.TryParse(sentTok.ToString(), Nothing,
                                         Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                                         dt) Then
                        Return dt
                    End If
                End If
                Dim recvTok = r("receivedDateTime")
                If recvTok IsNot Nothing AndAlso recvTok.Type <> JTokenType.Null Then
                    Dim dt As DateTime
                    If DateTime.TryParse(recvTok.ToString(), Nothing,
                                         Globalization.DateTimeStyles.AdjustToUniversal Or Globalization.DateTimeStyles.AssumeUniversal,
                                         dt) Then
                        Return dt
                    End If
                End If
            End If
        Catch
        End Try
        Return h?.LastModifiedUtc
    End Function

    Private Sub lvResults_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles lvResults.SelectedIndexChanged
        btnOpen.Enabled = lvResults.SelectedItems.Count > 0
    End Sub

    Private Sub lvResults_DoubleClick(sender As Object, e As EventArgs) _
        Handles lvResults.DoubleClick
        OpenSelectedHit()
    End Sub

    Private Sub btnOpen_Click(sender As Object, e As EventArgs) Handles btnOpen.Click
        OpenSelectedHit()
    End Sub

    ' ════════════════════════════════════════════════════════════════════
    '  Open hit in Outlook (or browser as fallback)
    ' ════════════════════════════════════════════════════════════════════
    Private Async Sub OpenSelectedHit()
        If lvResults.SelectedItems.Count = 0 Then Return
        Dim idx = CInt(lvResults.SelectedItems(0).Tag)
        If idx < 0 OrElse idx >= _hits.Count Then Return
        Dim hit = _hits(idx)

        SetBusy(True, "Opening message…")
        Try
            ' 1) Get the InternetMessageID — sometimes present in the search hit, otherwise fetch.
            Dim imid As String = TryGetInternetMessageId(hit)
            If String.IsNullOrEmpty(imid) Then
                Try
                    Dim full = Await M365Service.GetMessageAsync(_context, hit.Id, M365MessageFields.Headers).ConfigureAwait(False)
                    imid = If(full?.InternetMessageId, "")
                Catch ex As Exception
                    Debug.WriteLine("[M365Search] GetMessageAsync failed: " & ex.Message)
                End Try
            End If

            ' Steps 2 & 3 touch Outlook COM + UI — must run on the UI thread.
            Dim capturedImid As String = imid
            UiInvoke(Sub()
                         Dim opened As Boolean = False
                         If Not String.IsNullOrEmpty(capturedImid) Then
                             Try
                                 Dim app = Globals.ThisAddIn.Application
                                 Dim mail = FindMailItemByInternetMessageId(app, capturedImid)
                                 If mail IsNot Nothing Then
                                     mail.Display(False)
                                     opened = True
                                 End If
                             Catch ex As Exception
                                 Debug.WriteLine("[M365Search] Outlook lookup failed: " & ex.Message)
                             End Try
                         End If

                         If Not opened Then
                             If Not String.IsNullOrEmpty(hit.WebUrl) Then
                                 Try
                                     Process.Start(New ProcessStartInfo(hit.WebUrl) With {.UseShellExecute = True})
                                     lblStatus.Text = "Opened web link (message not found in local store)."
                                 Catch ex As Exception
                                     SharedMethods.ShowCustomMessageBox("Could not open message: " & ex.Message, "Microsoft 365")
                                 End Try
                             Else
                                 SharedMethods.ShowCustomMessageBox(
                                     "Message could not be located in Outlook and no web link is available.",
                                     "Microsoft 365")
                             End If
                         Else
                             lblStatus.Text = "Opened in Outlook."
                         End If
                     End Sub)
        Finally
            UiPost(Sub() SetBusy(False))
        End Try
    End Sub

    ''' <summary>Reads internetMessageId from the raw Graph payload if Graph returned it on the search hit.</summary>
    Private Function TryGetInternetMessageId(hit As M365SearchHit) As String
        Try
            Dim r = TryCast(hit?.RawJson?("resource"), JObject)
            If r Is Nothing Then Return ""
            Return If(r("internetMessageId")?.ToString(), "")
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Walks every Outlook store and folder looking for a mail item whose
    ''' InternetMessageID matches. Stops at the first match.
    ''' </summary>
    Private Function FindMailItemByInternetMessageId(app As Outlook.Application, imid As String) As Outlook.MailItem
        If app Is Nothing OrElse String.IsNullOrEmpty(imid) Then Return Nothing
        Dim filter = "[InternetMessageID] = '" & imid.Replace("'", "''") & "'"
        For Each store As Outlook.Store In app.Session.Stores
            Try
                Dim root As Outlook.Folder = CType(store.GetRootFolder(), Outlook.Folder)
                Dim found = SearchFolderRecursive(root, filter)
                If found IsNot Nothing Then Return found
            Catch ex As Exception
                Debug.WriteLine("[M365Search] Store skipped: " & ex.Message)
            End Try
        Next
        Return Nothing
    End Function

    Private Function SearchFolderRecursive(folder As Outlook.Folder, filter As String) As Outlook.MailItem
        If folder Is Nothing Then Return Nothing
        Try
            If folder.DefaultItemType = Outlook.OlItemType.olMailItem Then
                Dim found = TryCast(folder.Items.Find(filter), Outlook.MailItem)
                If found IsNot Nothing Then Return found
            End If
        Catch ex As Exception
            Debug.WriteLine("[M365Search] Folder skipped '" & folder.Name & "': " & ex.Message)
        End Try
        Try
            For Each sub_ As Outlook.Folder In folder.Folders
                Dim r = SearchFolderRecursive(sub_, filter)
                If r IsNot Nothing Then Return r
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ' ════════════════════════════════════════════════════════════════════
    '  IProgress<M365SearchProgress>
    ' ════════════════════════════════════════════════════════════════════
    Public Sub Report(value As M365SearchProgress) Implements IProgress(Of M365SearchProgress).Report
        If InvokeRequired Then
            Try
                BeginInvoke(New Action(Sub() ApplyProgress(value)))
            Catch
            End Try
        Else
            ApplyProgress(value)
        End If
    End Sub

    Private Sub ApplyProgress(p As M365SearchProgress)
        If Not IsFormUsable() Then Return
        If p.Percent < 0 Then
            pb.Style = ProgressBarStyle.Marquee
        Else
            pb.Style = ProgressBarStyle.Continuous
            pb.Value = Math.Max(0, Math.Min(100, p.Percent))
        End If
        lblStatus.Text = $"{p.Stage}{If(p.Source = M365SearchSources.None, "", " — " & p.Source.ToString())}: {p.Message}  ({p.HitsSoFar} hit(s))"
    End Sub

    Private Sub SetBusy(busy As Boolean, Optional status As String = Nothing)
        btnSearch.Enabled = Not busy
        btnSignIn.Enabled = Not busy
        btnSignOut.Enabled = Not busy
        txtQuery.Enabled = Not busy
        numMax.Enabled = Not busy
        Cursor = If(busy, Cursors.AppStarting, Cursors.Default)
        If status IsNot Nothing Then lblStatus.Text = status
        If Not busy Then
            pb.Style = ProgressBarStyle.Continuous
            pb.Value = 0
        End If
    End Sub

End Class

''' <summary>Convenience launcher — call from any ribbon button or the Immediate window.</summary>
Public Module M365SearchTest
    Public Sub Show(context As ISharedContext)
        Dim f As New M365SearchTestForm(context)
        f.Show()
    End Sub
End Module