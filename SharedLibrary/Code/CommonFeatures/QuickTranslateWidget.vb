' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: QuickTranslateWidget.vb
' Purpose: Provides a non-modal, resizable translation widget allowing users to
'          quickly translate text using the LLM. Features debounced auto-translate
'          on typing pause, Enter key translation, and clipboard copy support.
'
' Architecture:
'  - Non-modal Form with two side-by-side panels (input TextBox, output RichTextBox)
'  - Source language input field with "(auto)" placeholder
'  - Target language input field with persistence via My.Settings
'  - Debounce timer (1 second) triggers translation after user stops typing
'  - Enter key also triggers immediate translation
'  - Buttons: Clear, [Collapse], [Clear 2], [Copy 2], Copy, Close
'  - Keyboard shortcuts: Escape closes, Ctrl+C copies result, Ctrl+L clears
'  - Spinner indicator during LLM call
'  - Window position/size persisted via My.Settings
'  - Click on word selects it and copies to clipboard
'  - Click on selected text triggers translation/action
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Namespace SharedLibrary
    ''' <summary>
    ''' A non-modal translation widget that provides quick LLM-powered translation.
    ''' </summary>
    Public Class QuickTranslateWidget
        Inherits Form

        Private Const DEBOUNCE_MS As Integer = 1000
        Private Const COLLAPSED_INPUT_PERCENT As Single = 40.0F
        Private Const COLLAPSED_OUTPUT_PERCENT As Single = 60.0F
        Private Const EXPANDED_INPUT_PERCENT As Single = 25.0F
        Private Const EXPANDED_OUTPUT1_PERCENT As Single = 37.5F
        Private Const EXPANDED_OUTPUT2_PERCENT As Single = 37.5F

        ' Win32 API for cue banner (placeholder text)
        Private Const EM_SETCUEBANNER As Integer = &H1501

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
        End Function

        ' Controls
        Private WithEvents txtInput As TextBox
        Private WithEvents rtbOutput As RichTextBox
        Private WithEvents rtbOutput2 As RichTextBox
        Private WithEvents txtSourceLanguage As TextBox
        Private WithEvents txtLanguage As TextBox
        Private WithEvents btnClear As Button
        Private WithEvents btnCopy As Button
        Private WithEvents btnClose As Button
        Private WithEvents btnCollapse As Button
        Private WithEvents btnClear2 As Button
        Private WithEvents btnCopy2 As Button
        Private lblSpinner As Label
        Private lblSpinner2 As Label
        Private mainTable As TableLayoutPanel
        Private toolTip As ToolTip

        ' Debounce timer
        Private debounceTimer As System.Windows.Forms.Timer

        ' Selection copy timer (for auto-copy after selection)
        Private _selectionCopyTimer As System.Windows.Forms.Timer
        Private _selectionCopyRtb As RichTextBox = Nothing
        Private Const SELECTION_COPY_DELAY_MS As Integer = 500

        ' Track mouse drag state
        Private _mouseDownPoint As Point = Point.Empty
        Private _isDragging As Boolean = False
        Private Const DRAG_THRESHOLD As Integer = 3

        ' Track selection state for click-on-selection detection
        Private _rtbOutput_SelectionBeforeClick As Tuple(Of Integer, Integer) = Nothing
        Private _rtbOutput2_SelectionBeforeClick As Tuple(Of Integer, Integer) = Nothing

        ' Flag to suppress SelectionChanged during programmatic selection
        Private _suppressSelectionChanged As Boolean = False

        ' Cancellation for ongoing translations
        Private _cts As CancellationTokenSource
        Private _cts2 As CancellationTokenSource

        ' Callback to perform translation (now with sourceLanguage parameter)
        Private ReadOnly _translateFunc As Func(Of String, String, String, CancellationToken, Task(Of String))

        ' Default language from context
        Private ReadOnly _defaultLanguage As String

        ' Expansion state
        Private _isExpanded As Boolean = False
        Private _collapsedWidth As Integer
        Private _collapsedX As Integer

        ' Store calculated minimum width for collapsed state
        Private _minWidthCollapsed As Integer

        ' Track if form is closing
        Private _isClosing As Boolean = False

        ''' <summary>
        ''' Creates a new QuickTranslateWidget.
        ''' </summary>
        ''' <param name="translateFunc">
        ''' Async function that takes (textToTranslate, targetLanguage, sourceLanguage, cancellationToken) and returns the translated text.
        ''' </param>
        ''' <param name="defaultLanguage">The default target language (from INI_Language1).</param>
        Public Sub New(translateFunc As Func(Of String, String, String, CancellationToken, Task(Of String)),
                       defaultLanguage As String)
            _translateFunc = translateFunc
            _defaultLanguage = If(defaultLanguage, "English")
            InitializeComponent()
            RestoreSettings()
        End Sub

        Private Sub InitializeComponent()
            Me.Text = $"{SharedMethods.AN} Translate on-the-fly"
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.MinimizeBox = True
            Me.MaximizeBox = False
            Me.ShowInTaskbar = True
            Me.TopMost = True
            Me.StartPosition = FormStartPosition.Manual
            Me.KeyPreview = True

            ' Set icon
            Try
                Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
            Catch
            End Try

            ' Font
            Dim stdFont As New Font("Segoe UI", 10.0F, FontStyle.Regular, GraphicsUnit.Point)
            Me.Font = stdFont

            ' Initialize tooltip
            toolTip = New ToolTip() With {
                .AutoPopDelay = 5000,
                .InitialDelay = 500,
                .ReshowDelay = 200,
                .ShowAlways = True
            }

            ' Main layout: 2 rows, 3 columns (3rd column for expanded view)
            mainTable = New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 3,
                .RowCount = 2,
                .Padding = New Padding(10, 10, 10, 5),
                .Margin = New Padding(0)
            }
            mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, COLLAPSED_INPUT_PERCENT))
            mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, COLLAPSED_OUTPUT_PERCENT))
            mainTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 0.0F)) ' Hidden initially
            mainTable.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
            mainTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))

            ' Input TextBox (left)
            txtInput = New TextBox() With {
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .Dock = DockStyle.Fill,
                .Font = stdFont,
                .AcceptsReturn = False
            }
            mainTable.Controls.Add(txtInput, 0, 0)

            ' Output RichTextBox (middle)
            rtbOutput = New RichTextBox() With {
                .Dock = DockStyle.Fill,
                .Font = stdFont,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = SystemColors.Window,
                .Cursor = Cursors.IBeam,
                .DetectUrls = False
            }
            mainTable.Controls.Add(rtbOutput, 1, 0)

            ' Second Output RichTextBox (right) - hidden initially
            rtbOutput2 = New RichTextBox() With {
                .Dock = DockStyle.Fill,
                .Font = stdFont,
                .ReadOnly = True,
                .BorderStyle = BorderStyle.FixedSingle,
                .BackColor = SystemColors.Info,
                .Cursor = Cursors.IBeam,
                .DetectUrls = False,
                .Visible = False
            }
            mainTable.Controls.Add(rtbOutput2, 2, 0)

            ' Set tooltips for output fields
            Dim outputTooltipText As String = "Click on word: select and copy to clipboard." & vbCrLf &
                                               "Select text: auto-copies after brief pause." & vbCrLf &
                                               "Click on selection: translate to source language."
            toolTip.SetToolTip(rtbOutput, outputTooltipText)
            toolTip.SetToolTip(rtbOutput2, "Click on word: select and copy to clipboard." & vbCrLf &
                                            "Select text: auto-copies after brief pause." & vbCrLf &
                                            "Click on selection: use as new input text.")

            ' Bottom row: use a TableLayoutPanel with improved layout
            ' Col 0 (Left): 100% width (takes usually remaining space)
            ' Col 1 (Right): AutoSize (ensures buttons are never covered)
            Dim bottomTable As New TableLayoutPanel() With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 2,
                .RowCount = 1,
                .Margin = New Padding(0),
                .Padding = New Padding(0, 5, 0, 0),
                .AutoSize = True
            }
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            bottomTable.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            mainTable.Controls.Add(bottomTable, 0, 1)
            mainTable.SetColumnSpan(bottomTable, 3)

            ' Left side: Source language + arrow + Target language + spinners
            Dim leftFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            bottomTable.Controls.Add(leftFlow, 0, 0)

            ' Source language textbox
            txtSourceLanguage = New TextBox() With {
                .Width = 100,
                .Font = stdFont,
                .Margin = New Padding(0, 2, 0, 0)
            }
            leftFlow.Controls.Add(txtSourceLanguage)
            toolTip.SetToolTip(txtSourceLanguage, "Source language (leave empty for auto-detect)")

            ' Arrow label
            Dim lblArrow As New Label() With {
                .Text = "▶",
                .AutoSize = True,
                .Margin = New Padding(5, 4, 5, 0),
                .TextAlign = ContentAlignment.MiddleCenter
            }
            leftFlow.Controls.Add(lblArrow)

            ' Target language textbox
            txtLanguage = New TextBox() With {
                .Width = 150,
                .Font = stdFont,
                .Margin = New Padding(0, 2, 0, 0)
            }
            leftFlow.Controls.Add(txtLanguage)
            toolTip.SetToolTip(txtLanguage, "Target language")

            ' Spinner label (needed for calculation even if hidden)
            lblSpinner = New Label() With {
                .Text = "⏳",
                .AutoSize = True,
                .Visible = False,
                .Margin = New Padding(10, 5, 0, 0)
            }
            leftFlow.Controls.Add(lblSpinner)

            ' Second spinner
            lblSpinner2 = New Label() With {
                .Text = "⏳",
                .AutoSize = True,
                .Visible = False,
                .Margin = New Padding(5, 5, 0, 0)
            }
            leftFlow.Controls.Add(lblSpinner2)

            ' Right side: Buttons right-aligned
            Dim rightFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .Dock = DockStyle.Fill,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .Padding = New Padding(0)
            }
            bottomTable.Controls.Add(rightFlow, 1, 0)

            ' Buttons (added in reverse order due to RightToLeft flow)
            btnClose = New Button() With {
                .Text = "Close",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnClose)

            btnCopy2 = New Button() With {
                .Text = "Copy 2",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0),
                .Visible = False
            }
            rightFlow.Controls.Add(btnCopy2)

            btnClear2 = New Button() With {
                .Text = "Clear 2",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0),
                .Visible = False
            }
            rightFlow.Controls.Add(btnClear2)

            btnCopy = New Button() With {
                .Text = "Copy",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnCopy)

            btnClear = New Button() With {
                .Text = "Clear",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0)
            }
            rightFlow.Controls.Add(btnClear)

            btnCollapse = New Button() With {
                .Text = "◀ Collapse",
                .AutoSize = True,
                .Margin = New Padding(5, 0, 0, 0),
                .Visible = False
            }
            rightFlow.Controls.Add(btnCollapse)

            Me.Controls.Add(mainTable)

            ' === Dynamic Size Calculation ===
            ' Calculate Heights
            Dim oneLineHeight As Integer = txtInput.Font.Height + 6
            Dim bottomRowHeight As Integer = btnClose.Height + 10
            Dim chromeHeight As Integer = Me.Height - Me.ClientSize.Height
            Dim minHeight As Integer = chromeHeight + oneLineHeight + bottomRowHeight + 25

            ' Calculate Widths for Collapse Mode
            ' Left side: Input + Arrow + Target + Spinner + Margins (approx manually summed to be safe)
            Dim leftContentWidth As Integer = txtSourceLanguage.Width + 5 + lblArrow.PreferredWidth + 10 + txtLanguage.Width + 5 + lblSpinner.PreferredWidth + 15
            ' Right side: Clear + Copy + Close + Margins
            Dim rightContentWidth As Integer = btnClear.PreferredSize.Width + 5 + btnCopy.PreferredSize.Width + 5 + btnClose.PreferredSize.Width + 5

            ' Total width required (content + padding)
            Dim totalContentWidth As Integer = leftContentWidth + rightContentWidth + mainTable.Padding.Horizontal + 20

            ' Store calculated min width for logic usage
            _minWidthCollapsed = CInt(Math.Max(520, totalContentWidth) * Me.DeviceDpi / 96.0F)

            Me.MinimumSize = New Size(_minWidthCollapsed, minHeight)
            Me.Size = New Size(CInt(Math.Max(600, _minWidthCollapsed + 50) * Me.DeviceDpi / 96.0F), minHeight)


            ' Timers
            debounceTimer = New System.Windows.Forms.Timer() With {.Interval = DEBOUNCE_MS}
            AddHandler debounceTimer.Tick, AddressOf OnDebounceTimerTick

            _selectionCopyTimer = New System.Windows.Forms.Timer() With {.Interval = SELECTION_COPY_DELAY_MS}
            AddHandler _selectionCopyTimer.Tick, AddressOf OnSelectionCopyTimerTick

            ' Event Handlers
            AddHandler rtbOutput.MouseDown, AddressOf RtbOutput_MouseDown
            AddHandler rtbOutput.MouseMove, AddressOf RtbOutput_MouseMove
            AddHandler rtbOutput.MouseUp, AddressOf RtbOutput_MouseUp
            AddHandler rtbOutput.SelectionChanged, AddressOf RtbOutput_SelectionChanged

            AddHandler rtbOutput2.MouseDown, AddressOf RtbOutput2_MouseDown
            AddHandler rtbOutput2.MouseMove, AddressOf RtbOutput2_MouseMove
            AddHandler rtbOutput2.MouseUp, AddressOf RtbOutput2_MouseUp
            AddHandler rtbOutput2.SelectionChanged, AddressOf RtbOutput2_SelectionChanged
        End Sub

        Protected Overrides Sub OnHandleCreated(e As EventArgs)
            MyBase.OnHandleCreated(e)
            ' Set cue banner (placeholder) for source language field
            SendMessage(txtSourceLanguage.Handle, EM_SETCUEBANNER, IntPtr.Zero, "(auto)")
        End Sub

        Private Sub RestoreSettings()
            Try
                Dim savedLang As String = My.Settings.QuickTranslateLanguage
                txtLanguage.Text = If(String.IsNullOrWhiteSpace(savedLang), _defaultLanguage, savedLang)

                Dim x As Integer = My.Settings.QuickTranslateX
                Dim y As Integer = My.Settings.QuickTranslateY
                Dim w As Integer = My.Settings.QuickTranslateWidth
                Dim h As Integer = My.Settings.QuickTranslateHeight

                If w > 0 AndAlso h > 0 Then
                    Me.Size = New Size(w, h)
                End If

                If x > 0 AndAlso y > 0 Then
                    Dim wa As Rectangle = Screen.FromPoint(New Point(x, y)).WorkingArea
                    If x >= wa.Left AndAlso x < wa.Right - 50 AndAlso
                       y >= wa.Top AndAlso y < wa.Bottom - 50 Then
                        Me.Location = New Point(x, y)
                    Else
                        PositionOnScreen()
                    End If
                Else
                    PositionOnScreen()
                End If
            Catch
                txtLanguage.Text = _defaultLanguage
                PositionOnScreen()
            End Try
        End Sub

        Private Sub PositionOnScreen()
            Dim wa As Rectangle = Screen.FromPoint(Cursor.Position).WorkingArea
            Const MARGIN As Integer = 30
            Me.Location = New Point(wa.Right - Me.Width - MARGIN, wa.Top + MARGIN)
        End Sub

        Private Sub SaveSettings()
            Try
                My.Settings.QuickTranslateLanguage = txtLanguage.Text
                If Not _isExpanded Then
                    My.Settings.QuickTranslateX = Me.Location.X
                    My.Settings.QuickTranslateY = Me.Location.Y
                    My.Settings.QuickTranslateWidth = Me.Size.Width
                    My.Settings.QuickTranslateHeight = Me.Size.Height
                End If
                My.Settings.Save()
            Catch
            End Try
        End Sub

        Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
            _isClosing = True
            SaveSettings()
            CancelOngoingTranslation()
            CancelOngoingTranslation2()
            debounceTimer?.Stop()
            _selectionCopyTimer?.Stop()
            toolTip?.Dispose()
            MyBase.OnFormClosing(e)
        End Sub

        Private Sub txtInput_TextChanged(sender As Object, e As EventArgs) Handles txtInput.TextChanged
            debounceTimer.Stop()
            If Not String.IsNullOrWhiteSpace(txtInput.Text) Then
                debounceTimer.Start()
            End If
        End Sub

        Private Sub txtInput_KeyDown(sender As Object, e As KeyEventArgs) Handles txtInput.KeyDown
            If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
                e.SuppressKeyPress = True
                debounceTimer.Stop()
                PerformTranslationAsync()
            End If
        End Sub

        Private Sub OnDebounceTimerTick(sender As Object, e As EventArgs)
            debounceTimer.Stop()
            PerformTranslationAsync()
        End Sub

        Private Async Sub PerformTranslationAsync()
            Dim textToTranslate As String = txtInput.Text.Trim()
            Dim targetLanguage As String = txtLanguage.Text.Trim()
            Dim sourceLanguage As String = txtSourceLanguage.Text.Trim()

            If String.IsNullOrWhiteSpace(textToTranslate) Then
                rtbOutput.Text = ""
                Return
            End If

            If String.IsNullOrWhiteSpace(targetLanguage) Then
                targetLanguage = _defaultLanguage
            End If

            CancelOngoingTranslation()
            _cts = New CancellationTokenSource()
            Dim token As CancellationToken = _cts.Token

            lblSpinner.Visible = True
            rtbOutput.Text = ""

            Try
                Dim result As String = Await Task.Run(
                    Async Function()
                        Return Await _translateFunc(textToTranslate, targetLanguage, sourceLanguage, token)
                    End Function, token).ConfigureAwait(False)

                If Not token.IsCancellationRequested AndAlso Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then
                                               rtbOutput.Text = If(result, "")
                                               lblSpinner.Visible = False
                                           End If
                                       End Sub)
                    Else
                        rtbOutput.Text = If(result, "")
                        lblSpinner.Visible = False
                    End If
                End If
            Catch ex As OperationCanceledException
                ' Cancelled - ignore
            Catch ex As Exception
                If Not token.IsCancellationRequested AndAlso Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then
                                               rtbOutput.Text = $"Error: {ex.Message}"
                                               lblSpinner.Visible = False
                                           End If
                                       End Sub)
                    Else
                        rtbOutput.Text = $"Error: {ex.Message}"
                        lblSpinner.Visible = False
                    End If
                End If
            Finally
                If Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then lblSpinner.Visible = False
                                       End Sub)
                    Else
                        lblSpinner.Visible = False
                    End If
                End If
            End Try
        End Sub

        Private Async Sub PerformWordTranslationAsync(word As String)
            If String.IsNullOrWhiteSpace(word) Then Return

            CancelOngoingTranslation2()
            _cts2 = New CancellationTokenSource()
            Dim token As CancellationToken = _cts2.Token

            ' Determine target and source language based on source language field
            Dim userSourceLanguage As String = txtSourceLanguage.Text.Trim()
            Dim userTargetLanguage As String = txtLanguage.Text.Trim()

            Dim targetLanguage As String
            Dim sourceLanguage As String

            If Not String.IsNullOrWhiteSpace(userSourceLanguage) Then
                ' Source language is specified: swap roles for word translation
                ' Target becomes the original source language
                ' Source becomes the original target language
                targetLanguage = userSourceLanguage
                sourceLanguage = userTargetLanguage
            Else
                ' Source language is empty: detect from input text
                targetLanguage = $"the language of this text ""{Me.txtInput.Text}"""
                sourceLanguage = userTargetLanguage
            End If

            lblSpinner2.Visible = True
            rtbOutput2.Text = ""

            Try
                Dim result As String = Await Task.Run(
                    Async Function()
                        Return Await _translateFunc(word, targetLanguage, sourceLanguage, token)
                    End Function, token).ConfigureAwait(False)

                If Not token.IsCancellationRequested AndAlso Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then
                                               rtbOutput2.Text = If(result, "")
                                               lblSpinner2.Visible = False
                                           End If
                                       End Sub)
                    Else
                        rtbOutput2.Text = If(result, "")
                        lblSpinner2.Visible = False
                    End If
                End If
            Catch ex As OperationCanceledException
                ' Cancelled - ignore
            Catch ex As Exception
                If Not token.IsCancellationRequested AndAlso Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then
                                               rtbOutput2.Text = $"Error: {ex.Message}"
                                               lblSpinner2.Visible = False
                                           End If
                                       End Sub)
                    Else
                        rtbOutput2.Text = $"Error: {ex.Message}"
                        lblSpinner2.Visible = False
                    End If
                End If
            Finally
                If Not _isClosing Then
                    If Me.InvokeRequired Then
                        Me.BeginInvoke(Sub()
                                           If Not _isClosing Then lblSpinner2.Visible = False
                                       End Sub)
                    Else
                        lblSpinner2.Visible = False
                    End If
                End If
            End Try
        End Sub

        Private Sub CancelOngoingTranslation()
            Try
                _cts?.Cancel()
                _cts?.Dispose()
                _cts = Nothing
            Catch
            End Try
        End Sub

        Private Sub CancelOngoingTranslation2()
            Try
                _cts2?.Cancel()
                _cts2?.Dispose()
                _cts2 = Nothing
            Catch
            End Try
        End Sub

        Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
            ClearAll()
        End Sub

        Private Sub ClearAll()
            CancelOngoingTranslation()
            CancelOngoingTranslation2()
            txtInput.Text = ""
            rtbOutput.Text = ""
            rtbOutput2.Text = ""
            txtInput.Focus()
        End Sub

        Private Sub btnCopy_Click(sender As Object, e As EventArgs) Handles btnCopy.Click
            CopyResult()
        End Sub

        Private Sub CopyResult()
            Dim text As String = rtbOutput.Text
            If Not String.IsNullOrEmpty(text) Then
                Try
                    Clipboard.SetText(text.Trim())
                    FlashControl(rtbOutput)
                Catch
                End Try
            End If
        End Sub

        Private Sub btnClear2_Click(sender As Object, e As EventArgs) Handles btnClear2.Click
            CancelOngoingTranslation2()
            rtbOutput2.Text = ""
        End Sub

        Private Sub btnCopy2_Click(sender As Object, e As EventArgs) Handles btnCopy2.Click
            Dim text As String = rtbOutput2.Text
            If Not String.IsNullOrEmpty(text) Then
                Try
                    Clipboard.SetText(text.Trim())
                    FlashControl(rtbOutput2)
                Catch
                End Try
            End If
        End Sub

        Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
            Me.Close()
        End Sub

        Private Sub btnCollapse_Click(sender As Object, e As EventArgs) Handles btnCollapse.Click
            CollapseWindow()
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            MyBase.OnKeyDown(e)
            If e.KeyCode = Keys.Escape Then
                Me.Close()
                e.Handled = True
            ElseIf e.Control AndAlso e.KeyCode = Keys.C AndAlso Not txtInput.Focused Then
                CopyResult()
                e.Handled = True
            ElseIf e.Control AndAlso e.KeyCode = Keys.L Then
                ClearAll()
                e.Handled = True
            End If
        End Sub

#Region "Output RichTextBox Click Handling"

        Private Sub RtbOutput_MouseDown(sender As Object, e As MouseEventArgs)
            If e.Button <> MouseButtons.Left Then Return
            _mouseDownPoint = e.Location
            _isDragging = False
            _rtbOutput_SelectionBeforeClick = Tuple.Create(rtbOutput.SelectionStart, rtbOutput.SelectionLength)
            _selectionCopyTimer.Stop()
        End Sub

        Private Sub RtbOutput_MouseMove(sender As Object, e As MouseEventArgs)
            If e.Button = MouseButtons.Left AndAlso _mouseDownPoint <> Point.Empty Then
                If Math.Abs(e.X - _mouseDownPoint.X) > DRAG_THRESHOLD OrElse
                   Math.Abs(e.Y - _mouseDownPoint.Y) > DRAG_THRESHOLD Then
                    _isDragging = True
                End If
            End If
        End Sub

        Private Sub RtbOutput_MouseUp(sender As Object, e As MouseEventArgs)
            If e.Button <> MouseButtons.Left Then Return

            If _isDragging Then
                ' User was selecting text - don't trigger click action
                ' Selection will be auto-copied via SelectionChanged timer
                _isDragging = False
                _mouseDownPoint = Point.Empty
                _rtbOutput_SelectionBeforeClick = Nothing
                Return
            End If

            HandleOutputClick(rtbOutput, e.Location, _rtbOutput_SelectionBeforeClick, isSecondOutput:=False)
            _mouseDownPoint = Point.Empty
            _rtbOutput_SelectionBeforeClick = Nothing
        End Sub

        Private Sub RtbOutput_SelectionChanged(sender As Object, e As EventArgs)
            If _suppressSelectionChanged Then Return
            HandleSelectionChanged(rtbOutput)
        End Sub

        Private Sub RtbOutput2_MouseDown(sender As Object, e As MouseEventArgs)
            If e.Button <> MouseButtons.Left Then Return
            _mouseDownPoint = e.Location
            _isDragging = False
            _rtbOutput2_SelectionBeforeClick = Tuple.Create(rtbOutput2.SelectionStart, rtbOutput2.SelectionLength)
            _selectionCopyTimer.Stop()
        End Sub

        Private Sub RtbOutput2_MouseMove(sender As Object, e As MouseEventArgs)
            If e.Button = MouseButtons.Left AndAlso _mouseDownPoint <> Point.Empty Then
                If Math.Abs(e.X - _mouseDownPoint.X) > DRAG_THRESHOLD OrElse
                   Math.Abs(e.Y - _mouseDownPoint.Y) > DRAG_THRESHOLD Then
                    _isDragging = True
                End If
            End If
        End Sub

        Private Sub RtbOutput2_MouseUp(sender As Object, e As MouseEventArgs)
            If e.Button <> MouseButtons.Left Then Return

            If _isDragging Then
                _isDragging = False
                _mouseDownPoint = Point.Empty
                _rtbOutput2_SelectionBeforeClick = Nothing
                Return
            End If

            HandleOutputClick(rtbOutput2, e.Location, _rtbOutput2_SelectionBeforeClick, isSecondOutput:=True)
            _mouseDownPoint = Point.Empty
            _rtbOutput2_SelectionBeforeClick = Nothing
        End Sub

        Private Sub RtbOutput2_SelectionChanged(sender As Object, e As EventArgs)
            If _suppressSelectionChanged Then Return
            HandleSelectionChanged(rtbOutput2)
        End Sub

        Private Sub HandleOutputClick(rtb As RichTextBox, location As Point,
                                       selectionBefore As Tuple(Of Integer, Integer),
                                       isSecondOutput As Boolean)
            Dim charIndex As Integer = rtb.GetCharIndexFromPosition(location)

            ' Check if click was within the previous selection
            If selectionBefore IsNot Nothing AndAlso selectionBefore.Item2 > 0 Then
                Dim selStart As Integer = selectionBefore.Item1
                Dim selLength As Integer = selectionBefore.Item2

                If charIndex >= selStart AndAlso charIndex < selStart + selLength Then
                    ' Click was within the selection - trigger action
                    Dim selectedText As String = rtb.Text.Substring(selStart, selLength).Trim()

                    If Not String.IsNullOrWhiteSpace(selectedText) Then
                        ' Restore selection visually
                        _suppressSelectionChanged = True
                        rtb.Select(selStart, selLength)
                        _suppressSelectionChanged = False

                        If isSecondOutput Then
                            txtInput.Text = selectedText
                            txtInput.Focus()
                        Else
                            ExpandWindow()
                            PerformWordTranslationAsync(selectedText)
                        End If
                    End If
                    Return
                End If
            End If

            ' Click was outside selection or no selection - select word at position and copy
            Dim wordText As String = SelectWordAtPositionSuppressed(rtb, location)
            If Not String.IsNullOrWhiteSpace(wordText) Then
                Try
                    Clipboard.SetText(wordText)
                    FlashControl(rtb)
                Catch
                End Try
            End If
        End Sub

        Private Sub HandleSelectionChanged(rtb As RichTextBox)
            If rtb.SelectionLength > 0 Then
                _selectionCopyTimer.Stop()
                _selectionCopyRtb = rtb
                _selectionCopyTimer.Start()
            Else
                If _selectionCopyRtb Is rtb Then
                    _selectionCopyTimer.Stop()
                    _selectionCopyRtb = Nothing
                End If
            End If
        End Sub

        Private Sub OnSelectionCopyTimerTick(sender As Object, e As EventArgs)
            _selectionCopyTimer.Stop()

            If _selectionCopyRtb IsNot Nothing AndAlso _selectionCopyRtb.SelectionLength > 0 Then
                Dim selectedText As String = _selectionCopyRtb.SelectedText.Trim()
                If Not String.IsNullOrWhiteSpace(selectedText) Then
                    Try
                        Clipboard.SetText(selectedText)
                        FlashControl(_selectionCopyRtb)
                    Catch
                    End Try
                End If
            End If

            _selectionCopyRtb = Nothing
        End Sub

#End Region

        ''' <summary>
        ''' Selects the word at the given position, suppressing SelectionChanged event.
        ''' </summary>
        Private Function SelectWordAtPositionSuppressed(rtb As RichTextBox, location As Point) As String
            Dim charIndex As Integer = rtb.GetCharIndexFromPosition(location)
            Dim text As String = rtb.Text

            If String.IsNullOrEmpty(text) OrElse charIndex < 0 OrElse charIndex >= text.Length Then
                Return Nothing
            End If

            Dim wordStart As Integer = charIndex
            Dim wordEnd As Integer = charIndex

            While wordStart > 0 AndAlso Not Char.IsWhiteSpace(text(wordStart - 1))
                wordStart -= 1
            End While

            While wordEnd < text.Length AndAlso Not Char.IsWhiteSpace(text(wordEnd))
                wordEnd += 1
            End While

            If wordEnd > wordStart Then
                _suppressSelectionChanged = True
                rtb.Select(wordStart, wordEnd - wordStart)
                _suppressSelectionChanged = False
                Return rtb.SelectedText.Trim()
            End If

            Return Nothing
        End Function

        Private Sub ExpandWindow()
            If _isExpanded Then Return

            _collapsedWidth = Me.Width
            _collapsedX = Me.Left

            Dim expandedWidth As Integer = CInt(_collapsedWidth * 1.6)
            Dim rightEdge As Integer = Me.Right
            Dim newLeft As Integer = rightEdge - expandedWidth

            Dim wa As Rectangle = Screen.FromControl(Me).WorkingArea
            If newLeft < wa.Left Then
                newLeft = wa.Left
                expandedWidth = rightEdge - newLeft
            End If

            Me.MinimumSize = New Size(CInt(750 * Me.DeviceDpi / 96.0F), Me.MinimumSize.Height)

            mainTable.ColumnStyles(0) = New ColumnStyle(SizeType.Percent, EXPANDED_INPUT_PERCENT)
            mainTable.ColumnStyles(1) = New ColumnStyle(SizeType.Percent, EXPANDED_OUTPUT1_PERCENT)
            mainTable.ColumnStyles(2) = New ColumnStyle(SizeType.Percent, EXPANDED_OUTPUT2_PERCENT)

            rtbOutput2.Visible = True
            btnCollapse.Visible = True
            btnClear2.Visible = True
            btnCopy2.Visible = True

            Me.SuspendLayout()
            Me.SetBounds(newLeft, Me.Top, expandedWidth, Me.Height)
            Me.ResumeLayout()

            _isExpanded = True
        End Sub

        Private Sub CollapseWindow()
            If Not _isExpanded Then Return

            CancelOngoingTranslation2()

            rtbOutput2.Visible = False
            rtbOutput2.Text = ""
            btnCollapse.Visible = False
            btnClear2.Visible = False
            btnCopy2.Visible = False

            ' Restore the correctly calculated minimum width instead of hardcoded 500
            Me.MinimumSize = New Size(_minWidthCollapsed, Me.MinimumSize.Height)

            mainTable.ColumnStyles(0) = New ColumnStyle(SizeType.Percent, COLLAPSED_INPUT_PERCENT)
            mainTable.ColumnStyles(1) = New ColumnStyle(SizeType.Percent, COLLAPSED_OUTPUT_PERCENT)
            mainTable.ColumnStyles(2) = New ColumnStyle(SizeType.Percent, 0.0F)

            Me.SuspendLayout()
            Me.SetBounds(_collapsedX, Me.Top, _collapsedWidth, Me.Height)
            Me.ResumeLayout()

            _isExpanded = False
        End Sub

        Private Sub FlashControl(ctrl As Control)
            If _isClosing OrElse ctrl Is Nothing OrElse ctrl.IsDisposed Then Return

            Dim original As Color = ctrl.BackColor
            ctrl.BackColor = Color.LightGreen

            Dim flashTimer As New System.Windows.Forms.Timer() With {.Interval = 200}
            AddHandler flashTimer.Tick, Sub(s, ev)
                                            flashTimer.Stop()
                                            flashTimer.Dispose()
                                            Try
                                                If Not _isClosing AndAlso ctrl IsNot Nothing AndAlso Not ctrl.IsDisposed Then
                                                    ctrl.BackColor = original
                                                End If
                                            Catch
                                            End Try
                                        End Sub
            flashTimer.Start()
        End Sub

        Public Sub ShowWidget()
            If Me.Visible Then
                Me.BringToFront()
                Me.Activate()
            Else
                Me.Show()
            End If
            txtInput.Focus()
        End Sub
    End Class
End Namespace