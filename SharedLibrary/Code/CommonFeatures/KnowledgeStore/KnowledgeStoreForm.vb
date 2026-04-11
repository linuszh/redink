' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: KnowledgeStoreForm.vb
' Purpose: WinForms modal dialog for managing Knowledge Store index entries.
'          Displays documents from catalog-based stores via per-store manifests.
'          Supports add store/remove/tag/refresh operations.
'
' Architecture / How it works:
'  - Reads active stores via KnowledgeStoreCatalog.GetActiveStores().
'  - For each store, reads its manifest via KnowledgeStoreManifest.Load().
'  - Displays entries in a DataGridView with columns: Title, Path, Store, Tags, Date.
'  - "New Store" creates a new catalog entry.
'  - "Add Document(s)" indexes selected files into the manifest of their store.
'  - "Remove" deletes selected manifest entries (local stores only).
'  - "Edit Tags" allows tagging entries for filtered retrieval via (kb:tag:...).
'  - "Refresh" re-indexes a document whose content may have changed on disk.
'
' External Dependencies:
'  - KnowledgeStoreCatalog for store definitions.
'  - KnowledgeStoreManifest for per-store document manifests.
'  - KnowledgeIndexer for document indexing.
'  - ISharedContext for path configuration.
'  - SharedMethods for UI helpers (ShowCustomMessageBox, GetLogoBitmap).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Modal dialog for browsing, adding, removing, and tagging Knowledge Store documents.
    ''' </summary>
    Public Class KnowledgeStoreForm
        Inherits Form

        Private ReadOnly _context As ISharedContext
        Private _entries As New List(Of ManifestRow)()
        Private ReadOnly _grid As New DataGridView()
        Private ReadOnly _btnAdd As New Button() With {.Text = "Add Document(s)", .AutoSize = True, .Margin = New Padding(4)}
        Private ReadOnly _btnRemove As New Button() With {.Text = "Remove", .AutoSize = True, .Margin = New Padding(4)}
        Private ReadOnly _btnEditTags As New Button() With {.Text = "Edit Tags", .AutoSize = True, .Margin = New Padding(4)}
        Private ReadOnly _btnRefresh As New Button() With {.Text = "Refresh", .AutoSize = True, .Margin = New Padding(4)}
        Private ReadOnly _btnClose As New Button() With {.Text = "Close", .AutoSize = True, .Margin = New Padding(4)}
        Private ReadOnly _lblStatus As New Label() With {.AutoSize = True, .Dock = DockStyle.Bottom, .Padding = New Padding(8)}
        Private ReadOnly _btnAddStore As New Button() With {.Text = "New Store", .AutoSize = True, .Margin = New Padding(4)}
        Private _isDirty As Boolean = False

        ''' <summary>
        ''' Couples a manifest entry with the store it belongs to.
        ''' </summary>
        Private Class ManifestRow
            Public Property Store As KnowledgeStoreCatalog.KnowledgeStoreDefinition
            Public Property Entry As KnowledgeStoreManager.KnowledgeEntry
        End Class

        ''' <summary>
        ''' Initializes the Knowledge Store management form.
        ''' </summary>
        Public Sub New(context As ISharedContext)
            _context = context
            InitializeForm()
            LoadEntries()
        End Sub

#Region "Form Initialization"

        ''' <summary>
        ''' Builds the form layout with DataGridView and action buttons.
        ''' </summary>
        Private Sub InitializeForm()
            Me.Text = $"{AN} — Knowledge Store"
            Me.Size = New Size(960, 580)
            Me.MinimumSize = New Size(700, 400)
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

            ' Configure DataGridView
            _grid.ReadOnly = True
            _grid.AllowUserToAddRows = False
            _grid.AllowUserToDeleteRows = False
            _grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            _grid.MultiSelect = True
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            _grid.RowHeadersVisible = False
            _grid.BackgroundColor = SystemColors.Window
            _grid.BorderStyle = BorderStyle.FixedSingle
            _grid.AllowUserToResizeRows = False
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            _grid.ColumnHeadersHeight = CInt(Me.Font.Height * 2.0)
            _grid.ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle() With {
                .Font = New Font(Me.Font, FontStyle.Bold),
                .Padding = New Padding(4, 2, 4, 2)
            }
            _grid.DefaultCellStyle = New DataGridViewCellStyle() With {
                .Padding = New Padding(4, 1, 4, 1)
            }
            _grid.RowTemplate.Height = CInt(Me.Font.Height * 1.5)

            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Title", .HeaderText = "Title", .FillWeight = 25
            })
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Path", .HeaderText = "Path", .FillWeight = 30
            })
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Store", .HeaderText = "Store", .FillWeight = 12
            })
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Tags", .HeaderText = "Tags", .FillWeight = 13
            })
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Indexed", .HeaderText = "Indexed", .FillWeight = 10
            })
            _grid.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Source", .HeaderText = "Source", .FillWeight = 10
            })

            ' Grid panel with padding
            Dim gridPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(10, 10, 10, 6)
            }
            gridPanel.Controls.Add(_grid)
            _grid.Dock = DockStyle.Fill

            ' Button panel
            Dim buttonPanel As New FlowLayoutPanel() With {
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.LeftToRight,
                .AutoSize = True,
                .Padding = New Padding(8, 4, 8, 8)
            }
            buttonPanel.Controls.AddRange(New Control() {_btnAddStore, _btnAdd, _btnRemove, _btnEditTags, _btnRefresh, _btnClose})

            ' Wire events
            AddHandler _btnAddStore.Click, AddressOf OnAddStore
            AddHandler _btnAdd.Click, AddressOf OnAddDocuments
            AddHandler _btnRemove.Click, AddressOf OnRemoveSelected
            AddHandler _btnEditTags.Click, AddressOf OnEditTags
            AddHandler _btnRefresh.Click, AddressOf OnRefreshSelected
            AddHandler _btnClose.Click, Sub(s, e) Me.Close()
            AddHandler Me.KeyDown, AddressOf OnKeyDown
            AddHandler Me.FormClosing, AddressOf OnFormClosing

            ' Add controls (order matters: bottom first, then fill)
            Me.Controls.Add(gridPanel)
            Me.Controls.Add(buttonPanel)
            Me.Controls.Add(_lblStatus)
        End Sub

#End Region

#Region "Add Store"

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
                    Dim def = KnowledgeStoreCatalog.CreateDefinition(
                        storeName.Trim(), fbd.SelectedPath, _context)

                    Dim allDefs = KnowledgeStoreCatalog.LoadAll(_context)
                    allDefs.Add(def)
                    KnowledgeStoreCatalog.SaveLocalCatalog(allDefs, _context)

                    ShowCustomMessageBox(
                        $"Knowledge Store '{def.Name}' created at:{vbCrLf}{def.ResolvedSourcePath}{vbCrLf}{vbCrLf}" &
                        "Use 'kbindex' in Freestyle to index its documents, or enable background indexing in Settings.",
                        $"{AN} Knowledge Store")

                    LoadEntries()

                Catch ex As Exception
                    ShowCustomMessageBox($"Error creating Knowledge Store: {ex.Message}", AN)
                End Try
            End Using
        End Sub

#End Region

#Region "Data Loading"

        ''' <summary>
        ''' Loads entries from all active stores' manifests.
        ''' </summary>
        Private Sub LoadEntries()
            _entries.Clear()

            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context)
            For Each store In stores
                Try
                    Dim manifest = KnowledgeStoreManifest.Load(store)
                    For Each entry In manifest.Entries
                        _entries.Add(New ManifestRow() With {
                            .Store = store,
                            .Entry = entry
                        })
                    Next
                Catch ex As Exception
                    Debug.WriteLine($"KnowledgeStoreForm: Error loading manifest for '{store.Name}': {ex.Message}")
                End Try
            Next

            PopulateGrid()
            UpdateStatus()
        End Sub

        Private Sub PopulateGrid()
            _grid.Rows.Clear()
            For Each row In _entries
                Dim tags As String = If(row.Entry.Tags IsNot Nothing, String.Join(", ", row.Entry.Tags), "")
                Dim indexed As String = If(row.Entry.IndexedDate <> Date.MinValue,
                                           row.Entry.IndexedDate.ToString("yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture),
                                           "")
                Dim source As String = If(row.Store.IsFromCentralCatalog, "Central", "Local")
                _grid.Rows.Add(row.Entry.Title, row.Entry.FilePath, row.Store.Name, tags, indexed, source)
            Next
        End Sub

        Private Sub UpdateStatus()
            Dim total = _entries.Count
            Dim storeCount = _entries.Select(Function(r) r.Store.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count()
            _lblStatus.Text = $"{total} document(s) across {storeCount} store(s)"
        End Sub

#End Region

#Region "Add Documents"

        Private Async Sub OnAddDocuments(sender As Object, e As EventArgs)
            ' Pick which store to add to
            Dim stores = KnowledgeStoreCatalog.GetActiveStores(_context).
                Where(Function(s) KnowledgeStoreCatalog.CanCurrentUserWrite(s, _context)).ToList()

            If stores.Count = 0 Then
                ShowCustomMessageBox("No writable Knowledge Stores found.", AN)
                Return
            End If

            Dim targetStore As KnowledgeStoreCatalog.KnowledgeStoreDefinition = stores(0)
            If stores.Count > 1 Then
                Dim storeNames = stores.Select(Function(s) s.Name).ToList()
                Dim chosen = ShowSelectionForm(
                    "Select the store to add documents to:",
                    $"{AN} — Add Documents", storeNames)
                If String.IsNullOrWhiteSpace(chosen) Then Return
                targetStore = stores.FirstOrDefault(
                    Function(s) s.Name.Equals(chosen, StringComparison.OrdinalIgnoreCase))
                If targetStore Is Nothing Then Return
            End If

            ' Use OpenFileDialog
            Dim filesToIndex As New List(Of String)
            Using ofd As New OpenFileDialog()
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.json;*.xml;*.html;*.htm;*.md;*.csv;*.yaml;*.yml;*.ini;*.log|All Files|*.*"
                ofd.Multiselect = True
                ofd.Title = $"Select documents to add to '{targetStore.Name}'"
                If ofd.ShowDialog() <> DialogResult.OK Then Return
                filesToIndex.AddRange(ofd.FileNames)
            End Using

            If filesToIndex.Count = 0 Then Return

            Dim addedCount As Integer = 0
            Dim manifest = KnowledgeStoreManifest.Load(targetStore)

            For Each filePath In filesToIndex
                Try
                    Dim entry = Await KnowledgeIndexer.IndexDocumentAsync(
                        filePath, _context, _context.INI_KnowledgeStoreUseLLMIndex)
                    If entry IsNot Nothing Then
                        manifest.AddOrUpdate(entry)
                        addedCount += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error indexing '{filePath}': {ex.Message}")
                End Try
            Next

            If addedCount > 0 Then
                manifest.Save(targetStore)
                LoadEntries()
            End If
        End Sub

#End Region

#Region "Remove Documents"

        Private Sub OnRemoveSelected(sender As Object, e As EventArgs)
            If _grid.SelectedRows.Count = 0 Then
                ShowCustomMessageBox("Please select one or more entries to remove.", AN)
                Return
            End If

            Dim hasProtected As Boolean = False
            Dim toRemove As New List(Of ManifestRow)()

            For Each gridRow As DataGridViewRow In _grid.SelectedRows
                Dim idx = gridRow.Index
                If idx >= 0 AndAlso idx < _entries.Count Then
                    If _entries(idx).Store.IsFromCentralCatalog Then
                        hasProtected = True
                    Else
                        toRemove.Add(_entries(idx))
                    End If
                End If
            Next

            If hasProtected AndAlso toRemove.Count = 0 Then
                ShowCustomMessageBox("The selected entries are from a central store and cannot be removed locally.", AN)
                Return
            End If
            If hasProtected Then
                ShowCustomMessageBox(
                    "Some selected entries are from a central store and will be skipped. " &
                    "Only local entries will be removed.", AN)
            End If

            If toRemove.Count = 0 Then Return

            Dim answer = ShowCustomYesNoBox(
                $"Remove {toRemove.Count} entry/entries from the Knowledge Store?",
                "Yes, remove", "Cancel")
            If answer <> 1 Then Return

            ' Group by store, remove from each manifest, save
            For Each grp In toRemove.GroupBy(Function(r) r.Store.Name, StringComparer.OrdinalIgnoreCase)
                Dim store = grp.First().Store
                Dim manifest = KnowledgeStoreManifest.Load(store)
                For Each item In grp
                    manifest.RemoveByPath(If(item.Entry.FilePath, ""))
                Next
                manifest.Save(store)
            Next

            LoadEntries()
        End Sub

#End Region

#Region "Edit Tags"

        Private Sub OnEditTags(sender As Object, e As EventArgs)
            If _grid.SelectedRows.Count = 0 Then
                ShowCustomMessageBox("Please select one or more entries to tag.", AN)
                Return
            End If

            Dim firstIdx = _grid.SelectedRows(0).Index
            If firstIdx < 0 OrElse firstIdx >= _entries.Count Then Return

            Dim currentTags As String = ""
            If _entries(firstIdx).Entry.Tags IsNot Nothing Then
                currentTags = String.Join(", ", _entries(firstIdx).Entry.Tags)
            End If

            Dim newTags = ShowCustomInputBox(
                "Enter tags separated by commas (e.g., contracts, employment, NDA):",
                $"{AN} — Edit Tags", True, currentTags)

            If newTags Is Nothing OrElse newTags.Equals("esc", StringComparison.OrdinalIgnoreCase) Then Return

            Dim tagList As String() = newTags.Split(","c).
                Select(Function(t) t.Trim()).
                Where(Function(t) Not String.IsNullOrWhiteSpace(t)).
                ToArray()

            ' Group modified entries by store for batch save
            Dim modified As New Dictionary(Of String, KnowledgeStoreCatalog.KnowledgeStoreDefinition)(StringComparer.OrdinalIgnoreCase)

            For Each gridRow As DataGridViewRow In _grid.SelectedRows
                Dim idx = gridRow.Index
                If idx >= 0 AndAlso idx < _entries.Count Then
                    Dim row = _entries(idx)
                    If row.Store.IsFromCentralCatalog Then Continue For
                    row.Entry.Tags = If(tagList.Length > 0, tagList, Nothing)
                    If Not modified.ContainsKey(row.Store.Name) Then
                        modified(row.Store.Name) = row.Store
                    End If
                End If
            Next

            ' Save each affected manifest
            For Each store In modified.Values
                Dim manifest = KnowledgeStoreManifest.Load(store)
                ' Re-apply tags from _entries for this store
                For Each row In _entries.Where(Function(r) r.Store.Name.Equals(store.Name, StringComparison.OrdinalIgnoreCase))
                    manifest.AddOrUpdate(row.Entry)
                Next
                manifest.Save(store)
            Next

            If modified.Count > 0 Then
                PopulateGrid()
            Else
                ShowCustomMessageBox("No local entries were modified. Central entries cannot be tagged locally.", AN)
            End If
        End Sub

#End Region

#Region "Refresh Documents"

        Private Async Sub OnRefreshSelected(sender As Object, e As EventArgs)
            If _grid.SelectedRows.Count = 0 Then
                ShowCustomMessageBox("Please select one or more entries to refresh.", AN)
                Return
            End If

            Dim refreshedStores As New Dictionary(Of String, KnowledgeStoreCatalog.KnowledgeStoreDefinition)(StringComparer.OrdinalIgnoreCase)
            Dim refreshed As Integer = 0

            For Each gridRow As DataGridViewRow In _grid.SelectedRows
                Dim idx = gridRow.Index
                If idx >= 0 AndAlso idx < _entries.Count Then
                    Dim row = _entries(idx)
                    If row.Store.IsFromCentralCatalog Then Continue For

                    Dim expandedPath = ExpandEnvironmentVariables(If(row.Entry.FilePath, ""))
                    If Not File.Exists(expandedPath) Then Continue For

                    Try
                        Dim updated = Await KnowledgeIndexer.IndexDocumentAsync(
                            expandedPath, _context, _context.INI_KnowledgeStoreUseLLMIndex)
                        If updated IsNot Nothing Then
                            updated.Tags = row.Entry.Tags ' Preserve tags
                            row.Entry = updated
                            If Not refreshedStores.ContainsKey(row.Store.Name) Then
                                refreshedStores(row.Store.Name) = row.Store
                            End If
                            refreshed += 1
                        End If
                    Catch ex As Exception
                        Debug.WriteLine($"Error refreshing '{expandedPath}': {ex.Message}")
                    End Try
                End If
            Next

            ' Save each affected manifest
            For Each store In refreshedStores.Values
                Dim manifest = KnowledgeStoreManifest.Load(store)
                For Each row In _entries.Where(Function(r) r.Store.Name.Equals(store.Name, StringComparison.OrdinalIgnoreCase))
                    manifest.AddOrUpdate(row.Entry)
                Next
                manifest.Save(store)
            Next

            If refreshed > 0 Then
                PopulateGrid()
                UpdateStatus()
            End If
        End Sub

#End Region

#Region "Save & Close"

        Private Sub OnFormClosing(sender As Object, e As FormClosingEventArgs)
            ' All mutations are saved immediately to manifests — no pending dirty state
        End Sub

        Private Sub OnKeyDown(sender As Object, e As KeyEventArgs)
            If e.KeyCode = Keys.Escape Then
                e.Handled = True
                Me.Close()
            End If
        End Sub

#End Region

    End Class

End Namespace