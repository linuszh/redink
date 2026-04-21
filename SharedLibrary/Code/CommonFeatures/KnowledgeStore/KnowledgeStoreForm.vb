' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreForm.vb
' Purpose:
'   WinForms administration dialog for configuring and maintaining Knowledge
'   Stores.
'
' Responsibilities:
'   - List all configured stores discovered from `KnowledgeStoreCatalog`.
'   - Edit local store metadata such as name, path, owner, role, and flags.
'   - Display per-store statistics including indexed documents, wiki pages,
'     embeddings, and writeability state.
'   - Trigger maintenance actions such as index, re-index, health check,
'     repair, revectorization, and schema editing.
'   - Run long-running maintenance work asynchronously and surface progress via
'     status text and tray notifications.
'
' Notes:
'   - This dialog is shared between Word and Outlook.
'   - Health Check, Lint Wiki, and Repair delegate into `KnowledgeWikiService`.
'   - Foreground indexing delegates into `KnowledgeStoreForegroundIndexer`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Threading.Tasks

Namespace SharedLibrary

    ''' <summary>
    ''' Modal dialog for full Knowledge Store administration: listing, editing,
    ''' statistics, and maintenance operations.
    ''' </summary>
    Public Class KnowledgeStoreForm
        Inherits Form

        Private ReadOnly _context As ISharedContext
        Private _stores As New List(Of KnowledgeStoreCatalog.KnowledgeStoreDefinition)()

        ' ── Left panel: Store list ──
        Private ReadOnly _lstStores As New ListBox()
        Private ReadOnly _btnAddStore As New Button() With {.Text = "New Store", .Margin = New Padding(4)}
        Private ReadOnly _btnDeleteStore As New Button() With {.Text = "Delete Store", .Margin = New Padding(4)}

        ' ── Right panel: Detail editing ──
        Private ReadOnly _txtName As New TextBox() With {.Width = 320}
        Private ReadOnly _txtSourcePath As New TextBox() With {.Width = 420}
        Private ReadOnly _btnBrowsePath As New Button() With {.Text = "Browse...", .Margin = New Padding(4)}
        Private ReadOnly _txtOwner As New TextBox() With {.Width = 320}
        Private ReadOnly _cboRole As New ComboBox() With {.Width = 180, .DropDownStyle = ComboBoxStyle.DropDownList}
        Private ReadOnly _chkActive As New CheckBox() With {.Text = "Active", .AutoSize = True}
        Private ReadOnly _chkScanSub As New CheckBox() With {.Text = "Scan subdirectories", .AutoSize = True}
        Private ReadOnly _btnSave As New Button() With {.Text = "Save Changes", .Margin = New Padding(4)}

        ' ── Statistics ──
        Private ReadOnly _lblStats As New Label() With {
            .AutoSize = False,
            .Width = 420,
            .Height = 124,
            .BorderStyle = BorderStyle.FixedSingle,
            .Padding = New Padding(6),
            .BackColor = SystemColors.Info,
            .ForeColor = SystemColors.InfoText
        }

        ' ── Operation buttons ──
        Private ReadOnly _btnIndex As New Button() With {.Text = "Index", .Margin = New Padding(4)}
        Private ReadOnly _btnReindex As New Button() With {.Text = "Full Re-Index", .Margin = New Padding(4)}
        Private ReadOnly _btnHealthCheck As New Button() With {.Text = "Health Check", .Margin = New Padding(4)}
        Private ReadOnly _btnRepair As New Button() With {.Text = "Repair", .Margin = New Padding(4)}
        Private ReadOnly _btnRevectorize As New Button() With {.Text = "Rebuild Embeddings", .Margin = New Padding(4)}
        Private ReadOnly _btnLint As New Button() With {.Text = "Lint Wiki", .Margin = New Padding(4)}
        Private ReadOnly _btnEditSchema As New Button() With {.Text = "Edit Schema", .Margin = New Padding(4)}
        Private ReadOnly _btnClose As New Button() With {.Text = "Close", .Margin = New Padding(4)}

        ' ── Status bar ──
        Private ReadOnly _lblStatus As New Label() With {.AutoSize = True, .Dock = DockStyle.Bottom, .Padding = New Padding(8)}

        ''' <summary>
        ''' Initializes the Knowledge Store administration form.
        ''' </summary>
        Public Sub New(context As ISharedContext)
            _context = context
            InitializeForm()
            LoadStores()
        End Sub

#Region "Form Initialization"

        Private Sub InitializeForm()
            Me.SuspendLayout()

            Me.Text = $"{AN} — Knowledge Store Administration"
            Me.AutoScaleMode = AutoScaleMode.Dpi
            Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
            Me.Size = New Size(1120, 780)
            Me.MinimumSize = New Size(980, 700)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.KeyPreview = True
            Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)

            ' Set icon
            Try
                Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
                Me.Icon = Icon.FromHandle(bmp.GetHicon())
                bmp.Dispose()
            Catch
            End Try

            ConfigureStandardButton(_btnAddStore)
            ConfigureStandardButton(_btnDeleteStore)
            ConfigureStandardButton(_btnSave)
            ConfigureStandardButton(_btnIndex)
            ConfigureStandardButton(_btnReindex)
            ConfigureStandardButton(_btnHealthCheck)
            ConfigureStandardButton(_btnRepair)
            ConfigureStandardButton(_btnRevectorize)
            ConfigureStandardButton(_btnLint)
            ConfigureStandardButton(_btnEditSchema)
            ConfigureStandardButton(_btnClose)

            ConfigureBrowseButton(_btnBrowsePath, _txtSourcePath)

            ' ── Role combo items ──
            _cboRole.Items.AddRange(New Object() {"personal", "shared", "readonly"})
            _cboRole.SelectedIndex = 0

            ' ── Left panel: store list ──
            Dim leftPanel As New Panel() With {
                .Dock = DockStyle.Left,
                .Width = 260,
                .Padding = New Padding(10)
            }

            Dim lblStores As New Label() With {
                .Text = "Knowledge Stores",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 0, 0, 4)
            }

            _lstStores.Dock = DockStyle.Fill
            _lstStores.IntegralHeight = False

            Dim leftBtnPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .Padding = New Padding(0, 4, 0, 0)
            }
            leftBtnPanel.Controls.AddRange(New Control() {_btnAddStore, _btnDeleteStore})

            leftPanel.Controls.Add(_lstStores)
            leftPanel.Controls.Add(leftBtnPanel)
            leftPanel.Controls.Add(lblStores)

            ' ── Right panel: details + operations ──
            Dim rightPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10),
                .AutoScroll = True
            }

            Dim detailLayout As New TableLayoutPanel() With {
                .Dock = DockStyle.Top,
                .AutoSize = True,
                .ColumnCount = 2,
                .Padding = New Padding(0)
            }
            detailLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            detailLayout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))

            Dim row As Integer = 0
            AddLabeledRow(detailLayout, row, "Name:", _txtName) : row += 1

            ' Source path with browse button
            Dim pathPanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0),
                .AutoSizeMode = AutoSizeMode.GrowAndShrink
            }
            pathPanel.Controls.Add(_txtSourcePath)
            pathPanel.Controls.Add(_btnBrowsePath)
            AddLabeledRow(detailLayout, row, "Source Path:", pathPanel) : row += 1

            AddLabeledRow(detailLayout, row, "Owner:", _txtOwner) : row += 1
            AddLabeledRow(detailLayout, row, "Role:", _cboRole) : row += 1

            Dim checkPanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = False,
                .Margin = New Padding(0)
            }
            checkPanel.Controls.Add(_chkActive)
            checkPanel.Controls.Add(_chkScanSub)
            AddLabeledRow(detailLayout, row, "", checkPanel) : row += 1

            ' Save button
            Dim savePanel As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 6, 0, 6)
            }
            savePanel.Controls.Add(_btnSave)

            ' Statistics
            Dim lblStatsHeader As New Label() With {
                .Text = "Statistics",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 8, 0, 4)
            }

            Dim statsPanel As New Panel() With {
                .Dock = DockStyle.Top,
                .Height = 132,
                .Padding = New Padding(0, 0, 0, 4)
            }
            _lblStats.Dock = DockStyle.Fill
            statsPanel.Controls.Add(_lblStats)

            ' Operations
            Dim lblOpsHeader As New Label() With {
                .Text = "Operations",
                .Font = New Font(Me.Font, FontStyle.Bold),
                .AutoSize = True,
                .Dock = DockStyle.Top,
                .Padding = New Padding(0, 6, 0, 4)
            }

            Dim opsPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Top,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .WrapContents = True,
                .Padding = New Padding(0, 0, 0, 4)
            }
            opsPanel.Controls.AddRange(New Control() {
                _btnIndex, _btnReindex, _btnHealthCheck, _btnRepair,
                _btnRevectorize, _btnLint, _btnEditSchema
            })

            ' Bottom close button
            Dim bottomPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.RightToLeft,
                .AutoSize = True,
                .Padding = New Padding(8, 4, 8, 8)
            }
            bottomPanel.Controls.Add(_btnClose)

            ' Add right panel children in reverse dock order (bottom-up for Top docking)
            rightPanel.Controls.Add(opsPanel)
            rightPanel.Controls.Add(lblOpsHeader)
            rightPanel.Controls.Add(statsPanel)
            rightPanel.Controls.Add(lblStatsHeader)
            rightPanel.Controls.Add(savePanel)
            rightPanel.Controls.Add(detailLayout)

            ' Splitter
            Dim splitter As New Splitter() With {.Dock = DockStyle.Left, .Width = 4}

            ' Add to form (order matters for docking)
            Me.Controls.Add(rightPanel)
            Me.Controls.Add(splitter)
            Me.Controls.Add(leftPanel)
            Me.Controls.Add(bottomPanel)
            Me.Controls.Add(_lblStatus)

            ' Wire events
            AddHandler _lstStores.SelectedIndexChanged, AddressOf OnStoreSelected
            AddHandler _btnAddStore.Click, AddressOf OnAddStore
            AddHandler _btnDeleteStore.Click, AddressOf OnDeleteStore
            AddHandler _btnBrowsePath.Click, AddressOf OnBrowsePath
            AddHandler _btnSave.Click, AddressOf OnSaveChanges
            AddHandler _btnIndex.Click, AddressOf OnIndex
            AddHandler _btnReindex.Click, AddressOf OnReindex
            AddHandler _btnHealthCheck.Click, AddressOf OnHealthCheck
            AddHandler _btnRepair.Click, AddressOf OnRepair
            AddHandler _btnRevectorize.Click, AddressOf OnRevectorize
            AddHandler _btnLint.Click, AddressOf OnLint
            AddHandler _btnEditSchema.Click, AddressOf OnEditSchema
            AddHandler _btnClose.Click, Sub(s, ev) Me.Close()
            AddHandler Me.KeyDown, AddressOf OnKeyDown
            Me.ResumeLayout(performLayout:=True)
        End Sub

        Private Shared Sub AddLabeledRow(table As TableLayoutPanel, row As Integer, labelText As String, control As Control)
            table.RowCount = row + 1
            table.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            If Not String.IsNullOrEmpty(labelText) Then
                Dim lbl As New Label() With {
                    .Text = labelText,
                    .AutoSize = True,
                    .Anchor = AnchorStyles.Left,
                    .Margin = New Padding(0, 6, 8, 2)
                }
                table.Controls.Add(lbl, 0, row)
            End If
            control.Margin = New Padding(0, 4, 0, 2)
            table.Controls.Add(control, 1, row)
        End Sub

        Private Shared Sub ConfigureStandardButton(button As Button)
            button.AutoSize = True
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 4, 10, 4)
            button.MinimumSize = New Size(0, 0)
        End Sub

        Private Shared Sub ConfigureBrowseButton(button As Button, relatedTextBox As TextBox)
            button.AutoSize = False
            button.UseVisualStyleBackColor = True
            button.Padding = New Padding(10, 0, 10, 0)
            button.Height = relatedTextBox.PreferredHeight + 2
            button.Width = Math.Max(90, TextRenderer.MeasureText(button.Text, button.Font).Width + 24)
        End Sub

#End Region

#Region "Store Loading"

        Private Sub LoadStores()
            _stores = KnowledgeStoreCatalog.LoadAll(_context)
            _lstStores.Items.Clear()

            For Each store In _stores
                _lstStores.Items.Add(KnowledgeStoreCatalog.GetDisplayLabel(store))
            Next

            UpdateStatus()

            If _lstStores.Items.Count > 0 Then
                _lstStores.SelectedIndex = 0
            Else
                ClearDetails()
            End If
        End Sub

        Private Sub UpdateStatus()
            Dim total = _stores.Count
            Dim active = _stores.Where(Function(s) s.Active).Count()
            _lblStatus.Text = $"{total} store(s) configured, {active} active"
        End Sub

#End Region

#Region "Store Selection & Details"

        Private Sub OnStoreSelected(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then
                ClearDetails()
                Return
            End If

            Dim store = _stores(idx)
            _txtName.Text = store.Name
            _txtSourcePath.Text = store.SourcePath
            _txtOwner.Text = store.Owner
            _chkActive.Checked = store.Active
            _chkScanSub.Checked = store.ScanSubdirectories

            Dim roleIdx = _cboRole.Items.IndexOf(If(store.Role, "personal").ToLowerInvariant())
            _cboRole.SelectedIndex = If(roleIdx >= 0, roleIdx, 0)

            ' Enable/disable editing for central stores
            Dim isLocal = Not store.IsFromCentralCatalog
            _txtName.ReadOnly = Not isLocal
            _txtSourcePath.ReadOnly = Not isLocal
            _txtOwner.ReadOnly = Not isLocal
            _cboRole.Enabled = isLocal
            _chkActive.Enabled = isLocal
            _chkScanSub.Enabled = isLocal
            _btnBrowsePath.Enabled = isLocal
            _btnSave.Enabled = isLocal
            _btnDeleteStore.Enabled = isLocal

            LoadStatistics(store)
        End Sub

        Private Sub ClearDetails()
            _txtName.Text = ""
            _txtSourcePath.Text = ""
            _txtOwner.Text = ""
            _cboRole.SelectedIndex = 0
            _chkActive.Checked = False
            _chkScanSub.Checked = False
            _lblStats.Text = "No store selected."
            _btnSave.Enabled = False
            _btnDeleteStore.Enabled = False
        End Sub

        Private Sub LoadStatistics(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition)
            Try
                Dim manifest = KnowledgeStoreManifest.Load(store)
                Dim docCount = manifest.Entries.Count
                Dim lastIndexed As String = "Never"
                If docCount > 0 Then
                    Dim maxDate = manifest.Entries.Max(Function(e) e.IndexedDate)
                    If maxDate <> Date.MinValue Then
                        lastIndexed = maxDate.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture)
                    End If
                End If

                Dim wikiPageCount As Integer = 0
                Dim embeddingCount As Integer = 0
                Dim sourceLabel = If(store.IsFromCentralCatalog, "Central", "Local")
                Dim resolvedPath = If(store.ResolvedSourcePath, ExpandEnvironmentVariables(store.SourcePath))
                Dim pathExists = Not String.IsNullOrWhiteSpace(resolvedPath) AndAlso Directory.Exists(resolvedPath)

                If pathExists Then
                    Dim wikiRoot = Path.Combine(resolvedPath, ".redink", KnowledgeStoreCatalog.WikiFolder)
                    If Directory.Exists(wikiRoot) Then
                        wikiPageCount = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
                            Where(Function(f)
                                      Dim n = Path.GetFileName(f)
                                      Return Not n.Equals("index.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("log.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("health_report.md", StringComparison.OrdinalIgnoreCase) AndAlso
                                             Not n.Equals("review_queue.md", StringComparison.OrdinalIgnoreCase)
                                  End Function).Count()
                    End If

                    Dim embPath = Path.Combine(resolvedPath, ".redink", ".embeddings.json")
                    If File.Exists(embPath) Then
                        Try
                            Dim embJson = File.ReadAllText(embPath, System.Text.Encoding.UTF8)
                            Dim arr = Newtonsoft.Json.Linq.JArray.Parse(embJson)
                            embeddingCount = arr.Count
                        Catch
                        End Try
                    End If
                End If

                Dim canWrite = KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context)

                _lblStats.Text =
                    $"Source: {sourceLabel}  |  Path exists: {If(pathExists, "Yes", "No")}" & vbCrLf &
                    $"Documents indexed: {docCount}" & vbCrLf &
                    $"Wiki pages: {wikiPageCount}" & vbCrLf &
                    $"Embedding vectors: {embeddingCount}" & vbCrLf &
                    $"Last indexed: {lastIndexed}" & vbCrLf &
                    $"Writable: {If(canWrite, "Yes", "No")}"
            Catch ex As Exception
                _lblStats.Text = $"Error loading statistics: {ex.Message}"
            End Try
        End Sub

#End Region

#Region "Add / Delete Store"

        Private Sub OnAddStore(sender As Object, e As EventArgs)
            If String.IsNullOrWhiteSpace(_context.INI_KnowledgeStorePathLocal) Then
                ShowCustomMessageBox(
                    "No local Knowledge Store catalog path is configured. " &
                    "Please set 'KnowledgeStorePathLocal' in your configuration file.", AN)
                Return
            End If

            Dim storeName = ShowCustomInputBox(
                "Enter a name for the new Knowledge Store (e.g., 'Contracts', 'Research'):",
                $"{AN} — New Knowledge Store", True, "")
            If String.IsNullOrWhiteSpace(storeName) Then Return

            Dim existing = KnowledgeStoreCatalog.GetStoreByName(storeName.Trim(), _context)
            If existing IsNot Nothing Then
                ShowCustomMessageBox($"A Knowledge Store named '{storeName.Trim()}' already exists.", AN)
                Return
            End If

            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select the root directory for this Knowledge Store"
                fbd.ShowNewFolderButton = True
                If fbd.ShowDialog() <> DialogResult.OK Then Return

                Try
                    Dim resolvedPath = fbd.SelectedPath
                    If Not Directory.Exists(resolvedPath) Then
                        Directory.CreateDirectory(resolvedPath)
                    End If

                    KnowledgeWikiService.InitializeWikiStructure(resolvedPath)

                    Dim def = KnowledgeStoreCatalog.CreateDefinition(
                        storeName.Trim(), fbd.SelectedPath, _context)

                    Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
                    allDefs.Add(def)
                    KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                    ShowCustomMessageBox(
                        $"Knowledge Store '{def.Name}' created at:{vbCrLf}{def.ResolvedSourcePath}",
                        $"{AN} Knowledge Store")

                    LoadStores()

                    ' Select the newly added store
                    For i = 0 To _stores.Count - 1
                        If _stores(i).Name.Equals(storeName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            _lstStores.SelectedIndex = i
                            Exit For
                        End If
                    Next

                Catch ex As Exception
                    ShowCustomMessageBox($"Error creating Knowledge Store: {ex.Message}", AN)
                End Try
            End Using
        End Sub

        Private Sub OnDeleteStore(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then Return

            Dim store = _stores(idx)
            If store.IsFromCentralCatalog Then
                ShowCustomMessageBox("Central stores cannot be deleted from here. Contact your administrator.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                $"Delete the Knowledge Store definition '{store.Name}'?{vbCrLf}" &
                $"(Source files at '{store.ResolvedSourcePath}' will NOT be deleted.)",
                "Yes, delete", "Cancel")
            If answer <> 1 Then Return

            Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
            allDefs.RemoveAll(Function(d)
                                  Return Not d.IsFromCentralCatalog AndAlso
                                         String.Equals(d.StoreId, store.StoreId, StringComparison.OrdinalIgnoreCase)
                              End Function)
            KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

            LoadStores()
        End Sub

#End Region

#Region "Edit & Save"

        Private Sub OnBrowsePath(sender As Object, e As EventArgs)
            Using fbd As New FolderBrowserDialog()
                fbd.Description = "Select the root directory for this Knowledge Store"
                fbd.ShowNewFolderButton = True
                If Not String.IsNullOrWhiteSpace(_txtSourcePath.Text) Then
                    Dim expanded = ExpandEnvironmentVariables(_txtSourcePath.Text)
                    If Directory.Exists(expanded) Then
                        fbd.SelectedPath = expanded
                    End If
                End If
                If fbd.ShowDialog() = DialogResult.OK Then
                    _txtSourcePath.Text = fbd.SelectedPath
                End If
            End Using
        End Sub

        Private Sub OnSaveChanges(sender As Object, e As EventArgs)
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then Return

            Dim store = _stores(idx)
            If store.IsFromCentralCatalog Then
                ShowCustomMessageBox("Central store definitions cannot be modified here.", AN)
                Return
            End If

            If String.IsNullOrWhiteSpace(_txtName.Text) Then
                ShowCustomMessageBox("Store name cannot be empty.", AN)
                Return
            End If

            If String.IsNullOrWhiteSpace(_txtSourcePath.Text) Then
                ShowCustomMessageBox("Source path cannot be empty.", AN)
                Return
            End If

            ' Update the in-memory definition
            store.Name = _txtName.Text.Trim()
            store.SourcePath = KnowledgeStoreCatalog.StripQuotes(_txtSourcePath.Text)
            store.Owner = _txtOwner.Text.Trim()
            store.Role = If(_cboRole.SelectedItem IsNot Nothing, _cboRole.SelectedItem.ToString(), "personal")
            store.Active = _chkActive.Checked
            store.ScanSubdirectories = _chkScanSub.Checked
            store.ResolvedSourcePath = ExpandEnvironmentVariables(store.SourcePath)

            ' Save back to local catalog
            Try
                Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)

                ' Find and replace the matching local definition
                Dim found = False
                For i = 0 To allDefs.Count - 1
                    If Not allDefs(i).IsFromCentralCatalog AndAlso
                       String.Equals(allDefs(i).StoreId, store.StoreId, StringComparison.OrdinalIgnoreCase) Then
                        allDefs(i) = store
                        found = True
                        Exit For
                    End If
                Next

                If Not found Then
                    ' Match by old name if StoreId changed
                    For i = 0 To allDefs.Count - 1
                        If Not allDefs(i).IsFromCentralCatalog AndAlso
                           String.Equals(allDefs(i).Name, _stores(idx).Name, StringComparison.OrdinalIgnoreCase) Then
                            allDefs(i) = store
                            found = True
                            Exit For
                        End If
                    Next
                End If

                KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                ShowCustomMessageBox($"Store '{store.Name}' saved successfully.", AN)

                ' Reload to reflect changes
                Dim selectedName = store.Name
                LoadStores()
                For i = 0 To _stores.Count - 1
                    If _stores(i).Name.Equals(selectedName, StringComparison.OrdinalIgnoreCase) Then
                        _lstStores.SelectedIndex = i
                        Exit For
                    End If
                Next
            Catch ex As Exception
                ShowCustomMessageBox($"Error saving store: {ex.Message}", AN)
            End Try
        End Sub

#End Region

#Region "Operations"

        Private Async Function RunMaintenanceJobAsync(store As KnowledgeStoreCatalog.KnowledgeStoreDefinition,
                                                      operationName As String,
                                                      work As Func(Of Action(Of String), Task(Of String))) As Task(Of String)
            Dim trayIcon As NotifyIcon = Nothing
            Dim lastBalloonKey As String = ""

            Try
                Try
                    trayIcon = New NotifyIcon() With {
                        .Icon = SystemIcons.Information,
                        .Visible = True
                    }
                Catch
                    trayIcon = Nothing
                End Try

                UpdateMaintenanceStatusSafe(store.Name, operationName, "Queued as a background job.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' started.")

                Dim progressCallback As Action(Of String) =
                    Sub(statusText)
                        UpdateMaintenanceStatusSafe(store.Name, operationName, statusText, trayIcon)

                        Dim nextBalloonKey As String = ""
                        Dim nextBalloonText As String = ""

                        If statusText.IndexOf("waiting for", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            nextBalloonKey = "wait"
                            nextBalloonText = $"{operationName} for '{store.Name}' is waiting for {KnowledgeStoreHostGate.HostDisplayName} to become idle."
                        ElseIf statusText.IndexOf("starting ai step", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               statusText.IndexOf("continuing", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            nextBalloonKey = "run"
                            nextBalloonText = $"{operationName} for '{store.Name}' resumed."
                        End If

                        If nextBalloonKey <> "" AndAlso
                           Not String.Equals(lastBalloonKey, nextBalloonKey, StringComparison.Ordinal) Then
                            lastBalloonKey = nextBalloonKey
                            ShowMaintenanceBalloonSafe(trayIcon, $"{AN} Knowledge Store", nextBalloonText)
                        End If
                    End Sub

                Dim result = Await Task.Run(
                    Async Function()
                        Return Await work(progressCallback).ConfigureAwait(False)
                    End Function)

                UpdateMaintenanceStatusSafe(store.Name, operationName, "Completed.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' completed.")

                Return result
            Catch ex As Exception
                UpdateMaintenanceStatusSafe(store.Name, operationName, "Failed.", trayIcon)
                ShowMaintenanceBalloonSafe(
                    trayIcon,
                    $"{AN} Knowledge Store",
                    $"{operationName} for '{store.Name}' failed: {ex.Message}")
                Throw
            Finally
                If trayIcon IsNot Nothing Then
                    Try
                        trayIcon.Visible = False
                        trayIcon.Dispose()
                    Catch
                    End Try
                End If
            End Try
        End Function

        Private Sub UpdateMaintenanceStatusSafe(storeName As String,
                                                operationName As String,
                                                statusText As String,
                                                trayIcon As NotifyIcon)
            If Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(
                        New Action(Of String, String, String, NotifyIcon)(AddressOf UpdateMaintenanceStatusSafe),
                        storeName,
                        operationName,
                        statusText,
                        trayIcon)
                Catch
                End Try
                Return
            End If

            Dim fullText = $"{operationName} — {storeName}: {statusText}"
            _lblStatus.Text = fullText

            If trayIcon IsNot Nothing Then
                Dim toolTipText = $"{operationName}: {statusText}"
                If String.IsNullOrWhiteSpace(toolTipText) Then
                    toolTipText = $"{operationName}: working..."
                End If

                If toolTipText.Length > 63 Then
                    toolTipText = toolTipText.Substring(0, 60).TrimEnd() & "..."
                End If

                trayIcon.Text = toolTipText
            End If
        End Sub


        Private Sub ShowMaintenanceBalloonSafe(trayIcon As NotifyIcon, title As String, text As String)
            If trayIcon Is Nothing OrElse String.IsNullOrWhiteSpace(text) OrElse Me.IsDisposed Then
                Return
            End If

            If Me.InvokeRequired Then
                Try
                    Me.BeginInvoke(
                        New Action(Of NotifyIcon, String, String)(AddressOf ShowMaintenanceBalloonSafe),
                        trayIcon,
                        title,
                        text)
                Catch
                End Try
                Return
            End If

            Try
                trayIcon.BalloonTipIcon = ToolTipIcon.Info
                trayIcon.BalloonTipTitle = title
                trayIcon.BalloonTipText = text
                trayIcon.ShowBalloonTip(3000)
            Catch
            End Try
        End Sub

        Private Function GetSelectedStore() As KnowledgeStoreCatalog.KnowledgeStoreDefinition
            Dim idx = _lstStores.SelectedIndex
            If idx < 0 OrElse idx >= _stores.Count Then
                ShowCustomMessageBox("Please select a Knowledge Store first.", AN)
                Return Nothing
            End If
            Return _stores(idx)
        End Function

        Private Async Sub OnIndex(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            SetOperationButtonsEnabled(False)
            Try
                Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(
                    _context, storeName:=store.StoreId, forceReindex:=False)

                ShowCustomMessageBox(
                    $"Indexing complete for '{store.Name}':{vbCrLf}" &
                    $"Total: {result.TotalFiles}, Indexed: {result.IndexedFiles}, " &
                    $"Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}" &
                    If(result.WasCancelled, $"{vbCrLf}(Cancelled by user)", ""),
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
            Catch ex As Exception
                ShowCustomMessageBox($"Error during indexing: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
            End Try
        End Sub

        Private Async Sub OnReindex(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                $"This will force a full re-index of '{store.Name}', regenerating all metadata and Wiki summaries. " &
                "This may take a while and use API credits. Continue?",
                "Yes, re-index", "No, cancel")
            If answer <> 1 Then Return

            SetOperationButtonsEnabled(False)
            Try
                Dim result = Await KnowledgeStoreForegroundIndexer.RunAsync(
                    _context, storeName:=store.StoreId, forceReindex:=True)

                ShowCustomMessageBox(
                    $"Re-indexing complete for '{store.Name}':{vbCrLf}" &
                    $"Total: {result.TotalFiles}, Indexed: {result.IndexedFiles}, " &
                    $"Skipped: {result.SkippedFiles}, Failed: {result.FailedFiles}" &
                    If(result.WasCancelled, $"{vbCrLf}(Cancelled by user)", ""),
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
            Catch ex As Exception
                ShowCustomMessageBox($"Error during re-indexing: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
            End Try
        End Sub

        Private Async Sub OnHealthCheck(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            SetOperationButtonsEnabled(False)
            Try
                Dim report = Await RunMaintenanceJobAsync(
                    store:=store,
                    operationName:="Health Check",
                    work:=Function(progressCallback)
                              Return KnowledgeWikiService.LintWikiAsync(
                                  kbRootPath:=store.ResolvedSourcePath,
                                  context:=_context,
                                  autoApply:=False,
                                  progressCallback:=progressCallback)
                          End Function)

                ShowCustomWindow(
                    $"Health check report for '{store.Name}':",
                    report,
                    "You can copy this report to the clipboard.",
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                ShowCustomMessageBox($"Error during health check: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
                UpdateStatus()
            End Try
        End Sub

        Private Async Sub OnRepair(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            If Not KnowledgeStoreCatalog.CanCurrentUserWrite(store, _context) Then
                ShowCustomMessageBox($"You do not have write permission to '{store.Name}'.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                $"This will run automatic repairs on '{store.Name}', including LLM-assisted fixes. Continue?",
                "Yes, repair", "No, cancel")
            If answer <> 1 Then Return

            SetOperationButtonsEnabled(False)
            Try
                Dim summary = Await RunMaintenanceJobAsync(
                    store:=store,
                    operationName:="Repair",
                    work:=Function(progressCallback)
                              Return KnowledgeWikiService.ApplyWikiHealthFixesAsync(
                                  kbRootPath:=store.ResolvedSourcePath,
                                  context:=_context,
                                  includeLlmRepairs:=True,
                                  progressCallback:=progressCallback,
                                  operationName:="Repair")
                          End Function)

                ShowCustomWindow(
                    $"Repair summary for '{store.Name}':",
                    summary,
                    "You can copy this summary to the clipboard.",
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                ShowCustomMessageBox($"Error during repair: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
                UpdateStatus()
            End Try
        End Sub

        Private Async Sub OnLint(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            SetOperationButtonsEnabled(False)
            Try
                Dim report = Await RunMaintenanceJobAsync(
                    store:=store,
                    operationName:="Lint Wiki",
                    work:=Function(progressCallback)
                              Return KnowledgeWikiService.LintWikiAsync(
                                  kbRootPath:=store.ResolvedSourcePath,
                                  context:=_context,
                                  autoApply:=True,
                                  progressCallback:=progressCallback)
                          End Function)

                ShowCustomWindow(
                    $"Lint report for '{store.Name}':",
                    report,
                    "Auto-fixes have been applied. You can copy this report to the clipboard.",
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
                UpdateStatus()
            Catch ex As Exception
                ShowCustomMessageBox($"Error during wiki lint: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
                UpdateStatus()
            End Try
        End Sub

        Private Async Sub OnRevectorize(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            ' Count embeddable wiki pages
            Dim wikiRoot = Path.Combine(store.ResolvedSourcePath, ".redink", KnowledgeStoreCatalog.WikiFolder)
            If Not Directory.Exists(wikiRoot) Then
                ShowCustomMessageBox("No wiki pages found to rebuild embeddings for.", AN)
                Return
            End If

            Dim wikiPages = Directory.GetFiles(wikiRoot, "*.md", SearchOption.AllDirectories).
                Where(Function(f)
                          Dim n = Path.GetFileName(f)
                          Return Not n.Equals(KnowledgeStoreCatalog.IndexFile, StringComparison.OrdinalIgnoreCase) AndAlso
                                 Not n.Equals(KnowledgeStoreCatalog.LogFile, StringComparison.OrdinalIgnoreCase) AndAlso
                                 Not n.Equals("health_report.md", StringComparison.OrdinalIgnoreCase)
                      End Function).ToList()

            If wikiPages.Count = 0 Then
                ShowCustomMessageBox("No wiki pages found to rebuild embeddings for.", AN)
                Return
            End If

            Dim answer = ShowCustomYesNoBox(
                $"This will rebuild the embedding index for '{store.Name}' from {wikiPages.Count} existing wiki page(s) " &
                "using the currently configured embedding model. Continue?",
                "Yes, rebuild embeddings", "No, cancel")
            If answer <> 1 Then Return

            SetOperationButtonsEnabled(False)
            Try
                ProgressBarModule.GlobalProgressValue = 0
                ProgressBarModule.GlobalProgressMax = wikiPages.Count
                ProgressBarModule.GlobalProgressLabel = "Preparing embedding rebuild..."
                ProgressBarModule.CancelOperation = False
                ProgressBarModule.ShowProgressBarInSeparateThread(
                    $"{AN} Knowledge Store — Embeddings",
                    "Refreshing embeddings...")

                Dim rebuiltPages = Await KnowledgeEmbeddingService.RebuildAllWikiEmbeddingsAsync(
                    kbRootPath:=store.ResolvedSourcePath,
                    context:=_context,
                    progressPrefix:=store.Name,
                    progressOffset:=0,
                    progressTotal:=wikiPages.Count)

                Dim wasCancelled = ProgressBarModule.CancelOperation
                ProgressBarModule.CancelOperation = True

                ShowCustomMessageBox(
                    If(wasCancelled,
                       $"Embedding rebuild cancelled. Refreshed {rebuiltPages} wiki page embedding set(s).",
                       $"Embedding rebuild complete. Refreshed {rebuiltPages} wiki page embedding set(s)."),
                    $"{AN} Knowledge Store")

                LoadStatistics(store)
            Catch ex As Exception
                ProgressBarModule.CancelOperation = True
                ShowCustomMessageBox($"Error during embedding rebuild: {ex.Message}", AN)
            Finally
                SetOperationButtonsEnabled(True)
            End Try
        End Sub



        Private Sub OnEditSchema(sender As Object, e As EventArgs)
            Dim store = GetSelectedStore()
            If store Is Nothing Then Return

            If String.IsNullOrWhiteSpace(store.ResolvedSourcePath) Then
                ShowCustomMessageBox("The selected store does not have a valid source path.", AN)
                Return
            End If

            Try
                KnowledgeStoreSchema.LoadOrCreate(store.ResolvedSourcePath)

                Dim schemaPath = KnowledgeStoreSchema.GetSchemaPath(store.ResolvedSourcePath)
                If String.IsNullOrWhiteSpace(schemaPath) Then
                    ShowCustomMessageBox("Could not resolve the schema path for the selected store.", AN)
                    Return
                End If

                ShowTextFileEditor(
                    schemaPath,
                    $"Edit schema for Knowledge Store '{store.Name}'.",
                    ForceJson:=True,
                    _context:=_context)
            Catch ex As Exception
                ShowCustomMessageBox($"Error opening schema: {ex.Message}", AN)
            End Try
        End Sub

        Private Sub SetOperationButtonsEnabled(enabled As Boolean)
            _btnIndex.Enabled = enabled
            _btnReindex.Enabled = enabled
            _btnHealthCheck.Enabled = enabled
            _btnRepair.Enabled = enabled
            _btnRevectorize.Enabled = enabled
            _btnLint.Enabled = enabled
            _btnEditSchema.Enabled = enabled
            _btnAddStore.Enabled = enabled
            _btnDeleteStore.Enabled = enabled
            _btnSave.Enabled = enabled
            _lstStores.Enabled = enabled
        End Sub

#End Region

#Region "Key Handling"

        Private Sub OnKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Escape Then
                e.Handled = True
                Me.Close()
            End If
        End Sub

#End Region

    End Class

End Namespace