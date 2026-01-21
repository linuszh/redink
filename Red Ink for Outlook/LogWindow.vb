' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' File: LogWindow.vb
' Purpose: Provides a log window for displaying tooling operations in real-time.

Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

''' <summary>
''' A modeless log window that displays tooling operations and allows cancellation.
''' </summary>
Public Class LogWindow
    Inherits Form

    Private ReadOnly rtbLog As RichTextBox
    Private ReadOnly btnCancel As Button
    Private ReadOnly btnCopy As Button
    Private ReadOnly btnClear As Button

    Private _initialPositionSet As Boolean = False
    Private _autoCloseTimer As Timer = Nothing
    Private _autoCloseSecondsRemaining As Integer = 0
    Private _closeButtonBaseText As String = "Close"

    Public Event CancelRequested As EventHandler

    Public Sub StartAutoCloseCountdown(Optional seconds As Integer = 30)
        If seconds <= 0 Then seconds = 1

        If Me.InvokeRequired Then
            If Me.IsHandleCreated Then Me.BeginInvoke(Sub() StartAutoCloseCountdown(seconds))
            Return
        End If

        _autoCloseSecondsRemaining = seconds

        If _autoCloseTimer Is Nothing Then
            _autoCloseTimer = New Timer() With {.Interval = 1000}
            AddHandler _autoCloseTimer.Tick, AddressOf AutoCloseTimerTick
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
        _autoCloseTimer.Stop()
        _autoCloseTimer.Start()
    End Sub

    Private Sub AutoCloseTimerTick(sender As Object, e As EventArgs)
        _autoCloseSecondsRemaining -= 1

        If _autoCloseSecondsRemaining <= 0 Then
            _autoCloseTimer.Stop()
            Try
                Me.Close()
            Catch
            End Try
            Return
        End If

        btnCancel.Text = $"{_closeButtonBaseText} ({_autoCloseSecondsRemaining}s)"
    End Sub

    Private Sub StopAutoCloseCountdown()
        If _autoCloseTimer IsNot Nothing Then
            _autoCloseTimer.Stop()
        End If
        _autoCloseSecondsRemaining = 0
        If btnCancel IsNot Nothing Then
            btnCancel.Text = _closeButtonBaseText
        End If
    End Sub

    Public Sub New()
        Me.Text = $"{ThisAddIn.AN} Tooling Log"
        Me.Width = 600
        Me.Height = 400
        Me.StartPosition = FormStartPosition.Manual
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(450, 300)
        Me.TopMost = False
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.AutoScaleMode = AutoScaleMode.Font

        ' FIX 1: Double buffering prevents rendering artifacts on buttons/text
        Me.DoubleBuffered = True

        Try
            Me.Icon = Icon.FromHandle((New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))).GetHicon())
        Catch
        End Try

        Dim mainPanel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Padding = New Padding(10)
        }
        mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        mainPanel.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        rtbLog = New RichTextBox() With {
            .Dock = DockStyle.Fill, .ReadOnly = True, .BorderStyle = BorderStyle.None,
            .ScrollBars = RichTextBoxScrollBars.Vertical, .WordWrap = False, .DetectUrls = False,
            .BackColor = Me.BackColor, .HideSelection = False, .ShortcutsEnabled = True,
            .Font = New Font("Consolas", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        }
        mainPanel.Controls.Add(rtbLog, 0, 0)

        Dim buttonPanel As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.RightToLeft, .AutoSize = True
        }

        btnCancel = New Button() With {.Text = "Cancel", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnCancel.Click, AddressOf OnCancelClick

        btnCopy = New Button() With {.Text = "Copy Log", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnCopy.Click, AddressOf OnCopyClick

        btnClear = New Button() With {.Text = "Clear", .AutoSize = True, .Padding = New Padding(10, 5, 10, 5)}
        AddHandler btnClear.Click, AddressOf OnClearClick

        buttonPanel.Controls.Add(btnCancel)
        buttonPanel.Controls.Add(btnCopy)
        buttonPanel.Controls.Add(btnClear)
        mainPanel.Controls.Add(buttonPanel, 0, 1)
        Me.Controls.Add(mainPanel)

        ' FIX 2: Force handle creation immediately. This ensures BeginInvoke works 
        ' even if called before the window is fully Shown.
        Me.CreateControl()
    End Sub

    Protected Overrides Sub OnShown(e As EventArgs)
        MyBase.OnShown(e)
        If Not _initialPositionSet Then
            Dim wa = Screen.PrimaryScreen.WorkingArea
            Me.Location = New Point(wa.Right - Me.Width - 40, wa.Bottom - Me.Height - 40)
            _initialPositionSet = True
        End If
        Me.BringToFront()
        Me.Refresh() ' Ensure initial paint
    End Sub

    Public Sub AppendLog(message As String, Optional level As String = "info")
        If Me.IsDisposed OrElse rtbLog.IsDisposed Then Return

        ' FIX 3: Safe invoke pattern
        If Me.InvokeRequired Then
            ' Check handle again before invoking to avoid race conditions during close
            If Me.IsHandleCreated Then
                Me.BeginInvoke(Sub() AppendLogInternal(message, level))
            End If
        Else
            AppendLogInternal(message, level)
        End If
    End Sub

    Private Sub AppendLogInternal(message As String, level As String)
        If String.IsNullOrEmpty(message) Then Return
        If rtbLog.IsDisposed Then Return

        Dim timestamp = DateTime.Now.ToString("HH:mm:ss.fff")
        Dim textColor As Color = Color.Black
        Select Case (If(level, "info")).ToLowerInvariant()
            Case "error", "err", "fail" : textColor = Color.DarkRed
            Case "warn", "warning" : textColor = Color.DarkOrange
            Case "success", "ok" : textColor = Color.DarkGreen
            Case "step" : textColor = Color.DarkBlue
            Case "llm" : textColor = Color.DarkMagenta
        End Select

        ' Pause layout to prevent flicker during append
        rtbLog.SuspendLayout()
        Try
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionColor = Color.Gray
            rtbLog.AppendText($"[{timestamp}] ")
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.SelectionColor = textColor
            rtbLog.AppendText(message & Environment.NewLine)
            rtbLog.SelectionStart = rtbLog.TextLength
            rtbLog.ScrollToCaret()
        Finally
            rtbLog.ResumeLayout()
        End Try
    End Sub

    Private Sub OnCancelClick(sender As Object, e As EventArgs)
        btnCancel.Enabled = False
        btnCancel.Text = "Cancelling..."
        RaiseEvent CancelRequested(Me, EventArgs.Empty)
    End Sub

    Private Sub OnCopyClick(sender As Object, e As EventArgs)
        If Not String.IsNullOrEmpty(rtbLog.Text) Then Clipboard.SetText(rtbLog.Text)
    End Sub

    Private Sub OnClearClick(sender As Object, e As EventArgs)
        rtbLog.Clear()
    End Sub

    Public Sub MarkComplete()
        If Me.IsDisposed Then Return

        If Me.InvokeRequired Then
            If Me.IsHandleCreated Then Me.BeginInvoke(Sub() MarkCompleteInternal())
        Else
            MarkCompleteInternal()
        End If
    End Sub

    Private Sub MarkCompleteInternal()
        If btnCancel.IsDisposed Then Return

        StopAutoCloseCountdown()

        btnCancel.Text = "Close"
        _closeButtonBaseText = "Close"
        btnCancel.Enabled = True
        RemoveHandler btnCancel.Click, AddressOf OnCancelClick
        AddHandler btnCancel.Click, Sub()
                                        StopAutoCloseCountdown()
                                        Me.Close()
                                    End Sub

        Me.Text = $"{ThisAddIn.AN} Tooling Log (Complete)"

        StartAutoCloseCountdown(ThisAddIn.ToolingLog_AutoCloseDefaultSeconds)
    End Sub

End Class