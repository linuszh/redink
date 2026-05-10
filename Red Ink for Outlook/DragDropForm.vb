Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

Public Enum DragDropMode
    FileOnly = 0
    DirectoryOnly = 1
    FileOrDirectory = 2
End Enum

Partial Public Class DragDropForm
    Inherits System.Windows.Forms.Form

    Private _selectedFilePath As String = String.Empty
    Private _selectionMode As DragDropMode = DragDropMode.FileOnly

    Private Const LabelToButtonSpacing As Integer = 20
    Private Const ButtonToFormBottomSpacing As Integer = 24

    Public ReadOnly Property SelectedFilePath As String
        Get
            Return _selectedFilePath
        End Get
    End Property

    Public ReadOnly Property IsDirectory As Boolean
        Get
            Return Directory.Exists(_selectedFilePath)
        End Get
    End Property

    Public ReadOnly Property SelectionMode As DragDropMode
        Get
            Return _selectionMode
        End Get
    End Property

    Public Sub New()
        Me.New(DragDropMode.FileOnly)
    End Sub

    Public Sub New(mode As DragDropMode)
        InitializeComponent()
        _selectionMode = mode

        Me.AllowDrop = True

        Select Case _selectionMode
            Case DragDropMode.FileOnly
                Me.Text = "Drag & Drop Your File or Click Browse"
                Me.Label2.Text = "Drop or browse for a file."
            Case DragDropMode.DirectoryOnly
                Me.Text = "Drag & Drop Your Folder or Click Browse"
                Me.Label2.Text = "Drop or browse for a folder."
            Case DragDropMode.FileOrDirectory
                Me.Text = "Drag & Drop Your File or Folder, or Click Browse"
                Me.Label2.Text = "Drop or browse for a file or folder."
        End Select

        AdjustFormLayout()
    End Sub

    Public Sub SetInstructionText(text As String)
        Me.Label2.Text = If(text, "")
        AdjustFormLayout()
    End Sub

    Private Sub AdjustFormLayout()
        Me.Label2.PerformLayout()
        Me.btnBrowse.Top = Me.Label2.Bottom + LabelToButtonSpacing
        Me.ClientSize = New Size(Me.ClientSize.Width, Me.btnBrowse.Bottom + ButtonToFormBottomSpacing)
    End Sub

    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
        Dim iconValue As Icon = Icon.FromHandle(bmp.GetHicon())
        Me.Icon = iconValue
        bmp.Dispose()

        Me.TopMost = True
        Me.BringToFront()
        Me.Activate()
    End Sub

    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Try
            Dim paths As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If paths Is Nothing OrElse paths.Length = 0 Then Return

            Dim droppedPath As String = paths(0)
            Dim isDir As Boolean = Directory.Exists(droppedPath)
            Dim isFile As Boolean = File.Exists(droppedPath)

            Select Case _selectionMode
                Case DragDropMode.FileOnly
                    If Not isFile Then
                        SharedMethods.ShowCustomMessageBox("Please drop a file, not a folder.")
                        Return
                    End If

                Case DragDropMode.DirectoryOnly
                    If Not isDir Then
                        SharedMethods.ShowCustomMessageBox("Please drop a folder, not a file.")
                        Return
                    End If

                Case DragDropMode.FileOrDirectory
            End Select

            _selectedFilePath = droppedPath
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            SharedMethods.ShowCustomMessageBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Select Case _selectionMode
            Case DragDropMode.FileOnly
                BrowseForFile()

            Case DragDropMode.DirectoryOnly
                BrowseForFolder()

            Case DragDropMode.FileOrDirectory
                Dim result As Integer = SharedMethods.ShowCustomYesNoBox(
                    "What do you want to browse for?",
                    "File",
                    "Folder")

                If result = 1 Then
                    BrowseForFile()
                ElseIf result = 2 Then
                    BrowseForFolder()
                End If
        End Select
    End Sub

    Private Sub BrowseForFile()
        Using ofd As New OpenFileDialog()
            If ThisAddIn.INI_AllowLegacyDocFiles Then
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|All Files (*.*)|*.*"
            Else
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|All Files (*.*)|*.*"
            End If

            ofd.Title = "Select a File"
            ofd.Multiselect = False

            If ofd.ShowDialog() = DialogResult.OK Then
                _selectedFilePath = ofd.FileName
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

    Private Sub BrowseForFolder()
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select a Folder"
            fbd.ShowNewFolderButton = True

            If fbd.ShowDialog() = DialogResult.OK Then
                _selectedFilePath = fbd.SelectedPath
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

End Class