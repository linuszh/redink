' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WordHelpers.Snapshotter.vb
' Purpose: Document/URL snapshot management for tracking changes over time.
'          Supports downloading content from URLs or local paths, optional
'          clutter removal via LLM, comparison between snapshots, and
'          archive management.
'
' Architecture:
' - Snapshot Library (.txt files):
'     - Central library at INI_SnapshotLibPath (shared/network location)
'     - Local library at INI_SnapshotLibPathLocal (user-specific)
'     - INI-style format with [Document Name] sections containing metadata
'
' - Snapshot Archive:
'     - Default subdirectory "SnapshotArchive" relative to library file
'     - Can be overridden globally or per-document
'     - Files named: YYMMDD_HHMMSS_shortname.txt
'
' - Workflow:
'     - SelectSnapshotDocument: Entry point, shows document list
'     - CreateNewSnapshotDocument: Adds new document to track
'     - TakeSnapshot: Downloads content, optionally removes clutter, stores
'     - CompareSnapshots: Compares two snapshots using Word comparison
'
' Dependencies:
' - SharedLibrary (LLM, dialogs, file content retrieval)
' - ThisAddIn.WordHelpers (CompareTextsAndShowHtml for comparison display)
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports Slib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "Data Classes"

    ''' <summary>
    ''' Represents a document configured for snapshotting.
    ''' </summary>
    Private Class SnapshotDocument
        Public Property FriendlyName As String = ""
        Public Property Shortname As String = ""
        Public Property Location As String = ""
        Public Property RemoveClutter As Boolean = False
        Public Property SnapshotArchive As String = ""
        Public Property Snapshots As New List(Of String)()
        Public Property SourceLibraryPath As String = ""  ' Which library file this came from
        Public Property IsLocal As Boolean = False

        ''' <summary>
        ''' Gets the latest snapshot filename, or Nothing if no snapshots exist.
        ''' </summary>
        Public ReadOnly Property LatestSnapshot As String
            Get
                If Snapshots.Count = 0 Then Return Nothing
                ' Snapshots are stored in order; last one is most recent
                Return Snapshots(Snapshots.Count - 1)
            End Get
        End Property

        ''' <summary>
        ''' Parses snapshot filename to extract date.
        ''' Format: YYMMDD_HHMMSS_shortname.txt
        ''' </summary>
        Public Shared Function ParseSnapshotDate(filename As String) As DateTime?
            Try
                If String.IsNullOrEmpty(filename) Then Return Nothing
                Dim name = Path.GetFileNameWithoutExtension(filename)
                If name.Length < 13 Then Return Nothing ' YYMMDD_HHMMSS minimum

                Dim datePart = name.Substring(0, 6)
                Dim timePart = name.Substring(7, 6)

                Dim year = 2000 + Integer.Parse(datePart.Substring(0, 2))
                Dim month = Integer.Parse(datePart.Substring(2, 2))
                Dim day = Integer.Parse(datePart.Substring(4, 2))
                Dim hour = Integer.Parse(timePart.Substring(0, 2))
                Dim minute = Integer.Parse(timePart.Substring(2, 2))
                Dim second = Integer.Parse(timePart.Substring(4, 2))

                Return New DateTime(year, month, day, hour, minute, second)
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Gets display name including latest snapshot date and (local) indicator.
        ''' </summary>
        Public Function GetDisplayName() As String
            Dim display = FriendlyName
            Dim latestDate = If(LatestSnapshot IsNot Nothing, ParseSnapshotDate(LatestSnapshot), Nothing)
            If latestDate.HasValue Then
                display &= $" [{latestDate.Value:yyyy-MM-dd}]"
            End If
            If IsLocal Then
                display &= " (local)"
            End If
            Return display
        End Function
    End Class

    ''' <summary>
    ''' Represents parsed snapshot library data.
    ''' </summary>
    Private Class SnapshotLibrary
        Public Property FilePath As String = ""
        Public Property GlobalSnapshotArchive As String = ""
        Public Property Documents As New List(Of SnapshotDocument)()
    End Class

#End Region

#Region "Entry Point"

    ''' <summary>
    ''' Main entry point for the snapshot feature.
    ''' Shows document selection and handles all snapshot operations.
    ''' </summary>
    Public Sub SelectSnapshotDocument()
        ' Expand environment variables in library paths
        Dim centralLibPath = If(Not String.IsNullOrWhiteSpace(INI_SnapshotLibPath),
                            ExpandEnvironmentVariables(INI_SnapshotLibPath),
                            "")
        Dim localLibPath = If(Not String.IsNullOrWhiteSpace(INI_SnapshotLibPathLocal),
                          ExpandEnvironmentVariables(INI_SnapshotLibPathLocal),
                          "")

        ' Ensure paths are absolute if they are not already
        If Not String.IsNullOrWhiteSpace(centralLibPath) AndAlso Not Path.IsPathRooted(centralLibPath) Then
            centralLibPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), centralLibPath)
        End If
        If Not String.IsNullOrWhiteSpace(localLibPath) AndAlso Not Path.IsPathRooted(localLibPath) Then
            localLibPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), localLibPath)
        End If

        ' Check if any library path is configured
        Dim hasCentral As Boolean = Not String.IsNullOrWhiteSpace(centralLibPath) AndAlso File.Exists(centralLibPath)
        Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(localLibPath)

        If Not hasCentral AndAlso Not hasLocal Then
            ShowCustomMessageBox("No snapshot library is configured. Please set 'SnapshotLibPath' or 'SnapshotLibPathLocal'.", AN)
            Return
        End If

        ' Load libraries
        Dim allDocuments As New List(Of SnapshotDocument)()
        Dim localLibrary As SnapshotLibrary = Nothing
        Dim centralLibrary As SnapshotLibrary = Nothing

        ' Load local library
        If hasLocal Then
            If File.Exists(localLibPath) Then
                localLibrary = ParseSnapshotLibrary(localLibPath, isLocal:=True)
                allDocuments.AddRange(localLibrary.Documents)
            Else
                localLibrary = New SnapshotLibrary() With {.FilePath = localLibPath}
            End If
        End If

        ' Load central library
        If hasCentral Then
            centralLibrary = ParseSnapshotLibrary(centralLibPath, isLocal:=False)
            For Each doc As SnapshotDocument In centralLibrary.Documents
                allDocuments.Add(doc)
            Next
        End If

        ' Build selection list
        ShowSnapshotDocumentSelection(allDocuments, localLibrary, centralLibrary)
    End Sub

    ''' <summary>
    ''' Displays the document selection dialog and handles user choice.
    ''' </summary>
    Private Async Sub ShowSnapshotDocumentSelection(
        allDocuments As List(Of SnapshotDocument),
        localLibrary As SnapshotLibrary,
        centralLibrary As SnapshotLibrary)

        Dim items As New List(Of SelectionItem)()
        Dim docMap As New Dictionary(Of Integer, SnapshotDocument)()

        ' Sort documents: local first, then by name
        Dim sortedDocs = allDocuments.OrderByDescending(Function(d As SnapshotDocument) d.IsLocal).ThenBy(Function(d As SnapshotDocument) d.FriendlyName).ToList()

        Dim idx As Integer = 1
        For Each doc As SnapshotDocument In sortedDocs
            items.Add(New SelectionItem(doc.GetDisplayName(), idx))
            docMap(idx) = doc
            idx += 1
        Next

        ' Add management options
        Const ViewSnapshotsValue As Integer = -4
        Const DeleteSnapshotsValue As Integer = -3
        Const CompareExistingValue As Integer = -5
        Const AddNewValue As Integer = -1
        Const EditLibraryValue As Integer = -2

        If allDocuments.Count > 0 Then
            items.Add(New SelectionItem("View snapshots...", ViewSnapshotsValue))
            items.Add(New SelectionItem("Compare two snapshots...", CompareExistingValue))
            items.Add(New SelectionItem("Delete snapshots...", DeleteSnapshotsValue))
        End If

        ' Add new document is available when any library exists
        If localLibrary IsNot Nothing OrElse centralLibrary IsNot Nothing Then
            items.Add(New SelectionItem("Add new document...", AddNewValue))
        End If

        ' Edit local library only available when local library is configured
        If localLibrary IsNot Nothing Then
            items.Add(New SelectionItem("Edit local snapshot library...", EditLibraryValue))
        End If

        If items.Count = 0 Then
            ShowCustomMessageBox("No snapshot documents configured and no library available.", AN)
            Return
        End If

        Dim result = SelectValue(items, If(sortedDocs.Count > 0, 1, AddNewValue),
                                 "Select a document to snapshot or compare:",
                                 $"{AN} - Document Snapshots")

        If result = 0 Then
            Return
        ElseIf result = ViewSnapshotsValue Then
            ViewSnapshots(allDocuments, localLibrary, centralLibrary)
        ElseIf result = CompareExistingValue Then
            CompareExistingSnapshots(allDocuments, localLibrary, centralLibrary)
        ElseIf result = DeleteSnapshotsValue Then
            ManageSnapshots(allDocuments, localLibrary, centralLibrary)
        ElseIf result = AddNewValue Then
            ' Determine target library
            Dim targetLibrary As SnapshotLibrary = Nothing
            If localLibrary IsNot Nothing AndAlso centralLibrary IsNot Nothing Then
                ' Both exist — ask the user
                Dim libChoice = ShowCustomYesNoBox(
                    "Which library should the new document be added to?",
                    "Local library",
                    "Central library")
                If libChoice = 0 Then Return ' User cancelled
                targetLibrary = If(libChoice = 1, localLibrary, centralLibrary)
            ElseIf localLibrary IsNot Nothing Then
                targetLibrary = localLibrary
            ElseIf centralLibrary IsNot Nothing Then
                targetLibrary = centralLibrary
            End If

            If targetLibrary Is Nothing Then
                ShowCustomMessageBox("No library available to add the document to.", AN)
                Return
            End If

            Await CreateNewSnapshotDocument(targetLibrary)
            SelectSnapshotDocument()
        ElseIf result = EditLibraryValue Then
            If localLibrary IsNot Nothing AndAlso Not String.IsNullOrEmpty(localLibrary.FilePath) Then
                If Not File.Exists(localLibrary.FilePath) Then
                    EnsureSnapshotLibraryExists(localLibrary.FilePath)
                End If
                ShowTextFileEditor(localLibrary.FilePath, $"{AN} - Edit Snapshot Library", False, _context)
            End If
            SelectSnapshotDocument()
        ElseIf docMap.ContainsKey(result) Then
            Dim selectedDoc As SnapshotDocument = docMap(result)
            ProcessSnapshotDocument(selectedDoc, localLibrary, centralLibrary)
        End If
    End Sub

#End Region

#Region "Snapshot Processing"

    ''' <summary>
    ''' Processes a selected document: downloads new content and compares with existing snapshot.
    ''' </summary>
    Private Async Sub ProcessSnapshotDocument(
        doc As SnapshotDocument,
        localLibrary As SnapshotLibrary,
        centralLibrary As SnapshotLibrary)

        Try
            ' Determine which library this document belongs to
            Dim docLibrary As SnapshotLibrary = GetLibraryForDocument(doc, localLibrary, centralLibrary)

            ' Determine archive path
            Dim archivePath = ResolveArchivePath(doc, docLibrary)

            ' Add debug info if we ended up in a weird place
            If String.IsNullOrEmpty(archivePath) Then
                ShowCustomMessageBox($"Could not determine snapshot archive path.{vbCrLf}IsLocal: {doc.IsLocal}", AN)
                Return
            End If

            ' Ensure archive directory exists
            If Not Directory.Exists(archivePath) Then
                Try
                    Directory.CreateDirectory(archivePath)
                Catch ex As Exception
                    ShowCustomMessageBox($"Failed to create archive directory:{vbCrLf}{archivePath}{vbCrLf}{vbCrLf}{ex.Message}", AN)
                    Return
                End Try
            End If

            Dim downloadTimestamp As DateTime = DateTime.Now

            Dim splash As New Slib.SplashScreen($"Downloading: {doc.FriendlyName}...")
            splash.Show()
            splash.Refresh()
            System.Windows.Forms.Application.DoEvents()

            Dim newContent As String = Nothing
            Dim rawContent As String = Nothing

            Try
                rawContent = Await RetrieveDocumentContentAsync(doc.Location)
                If String.IsNullOrEmpty(rawContent) Then
                    splash.Close()
                    ShowCustomMessageBox($"Failed to retrieve content from:{vbCrLf}{doc.Location}", AN)
                    Return
                End If
            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"Error retrieving document:{vbCrLf}{ex.Message}", AN)
                Return
            End Try

            splash.Close()

            If doc.RemoveClutter Then
                Try
                    Dim userPrompt = "<TEXTTOPROCESS>" & vbCrLf & rawContent & vbCrLf & "</TEXTTOPROCESS>"
                    newContent = Await LLM(SP_RemoveClutter, userPrompt, "", "", 0, False, False)

                    If String.IsNullOrWhiteSpace(newContent) OrElse newContent.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then
                        Dim useRaw = ShowCustomYesNoBox(
                            $"Clutter removal failed: {If(newContent, "Empty result")}{vbCrLf}{vbCrLf}Use raw content instead?",
                            "Yes", "No")
                        If useRaw = 1 Then newContent = rawContent Else Return
                    End If
                Catch ex As Exception
                    Dim useRaw = ShowCustomYesNoBox($"Clutter removal error: {ex.Message}{vbCrLf}{vbCrLf}Use raw content instead?", "Yes", "No")
                    If useRaw = 1 Then newContent = rawContent Else Return
                End Try
            Else
                newContent = rawContent
            End If

            If doc.Snapshots.Count = 0 Then
                Dim snapshotFullPath = SaveSnapshotWithTimestamp(doc, newContent, archivePath, downloadTimestamp)
                If Not String.IsNullOrEmpty(snapshotFullPath) Then
                    Dim snapshotFilename = Path.GetFileName(snapshotFullPath)
                    AddSnapshotToLibrary(doc, snapshotFilename, docLibrary)
                    ShowCustomMessageBox($"Initial snapshot saved:{vbCrLf}{snapshotFullPath}", AN)
                End If
                Return
            End If

            Dim previousSnapshot = SelectSnapshotForComparison(doc, archivePath)
            If String.IsNullOrEmpty(previousSnapshot) Then Return

            Dim previousPath = Path.Combine(archivePath, previousSnapshot)
            If Not File.Exists(previousPath) Then
                ' If previous file missing, offer to save new one
                Dim saveAnyway = ShowCustomYesNoBox(
                    $"Previous snapshot missing:{vbCrLf}{previousPath}{vbCrLf}Save new snapshot anyway?",
                    "Yes", "No")
                If saveAnyway = 1 Then
                    Dim snapshotFullPath = SaveSnapshotWithTimestamp(doc, newContent, archivePath, downloadTimestamp)
                    If Not String.IsNullOrEmpty(snapshotFullPath) Then
                        Dim fn = Path.GetFileName(snapshotFullPath)
                        AddSnapshotToLibrary(doc, fn, docLibrary)
                        ShowCustomMessageBox($"New snapshot saved:{vbCrLf}{snapshotFullPath}", AN)
                    End If
                End If
                Return
            End If

            Dim previousContent = File.ReadAllText(previousPath, Encoding.UTF8)
            Dim previousHash = ComputeContentHash(previousContent)
            Dim newHash = ComputeContentHash(newContent)

            If previousHash = newHash Then
                ShowCustomMessageBox("No changes detected.", AN)
                Return
            End If

            CompareAndShowSnapshots(doc, previousContent, newContent, archivePath, docLibrary, downloadTimestamp)

        Catch ex As Exception
            ShowCustomMessageBox($"Error processing snapshot:{vbCrLf}{ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Allows user to select which snapshot to compare against.
    ''' </summary>
    Private Function SelectSnapshotForComparison(doc As SnapshotDocument, archivePath As String) As String
        If doc.Snapshots.Count = 1 Then Return doc.Snapshots(0)

        Dim items As New List(Of SelectionItem)()
        Dim snapshotMap As New Dictionary(Of Integer, String)()
        Dim sortedSnapshots = doc.Snapshots.OrderByDescending(Function(s) s).ToList()

        Dim idx = 1
        For Each snapshot In sortedSnapshots
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshot)
            Dim displayName = If(snapshotDate.HasValue, $"{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}", snapshot)
            Dim fullPath = Path.Combine(archivePath, snapshot)
            If File.Exists(fullPath) Then
                Dim fileInfo = New FileInfo(fullPath)
                displayName &= $" ({FormatFileSize(fileInfo.Length)})"
            End If
            items.Add(New SelectionItem(displayName, idx))
            snapshotMap(idx) = snapshot
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select snapshot to compare against:", $"{AN} - Select Snapshot")
        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then Return snapshotMap(result)
        Return Nothing
    End Function

    ''' <summary>
    ''' Compares two snapshot contents and shows the result with save option.
    ''' </summary>
    Private Sub CompareAndShowSnapshots(doc As SnapshotDocument, previousContent As String, newContent As String, archivePath As String, docLibrary As SnapshotLibrary, downloadTimestamp As DateTime)
        Dim snapshotSaved As Boolean = False

        CompareTextsAndShowHtmlWithSnapshotOption(
            previousContent,
            newContent,
            $"{AN} Snapshot Compare - {doc.FriendlyName}",
            Sub()
                If snapshotSaved Then
                    ShowCustomMessageBox("Snapshot already saved.", AN)
                    Return
                End If
                Dim snapshotFullPath = SaveSnapshotWithTimestamp(doc, newContent, archivePath, downloadTimestamp)
                If Not String.IsNullOrEmpty(snapshotFullPath) Then
                    Dim snapshotFilename = Path.GetFileName(snapshotFullPath)
                    AddSnapshotToLibrary(doc, snapshotFilename, docLibrary)
                    ShowCustomMessageBox($"Snapshot saved:{vbCrLf}{snapshotFullPath}", AN)
                    snapshotSaved = True
                End If
            End Sub)
    End Sub

    ''' <summary>
    ''' Compares two text strings using Word's comparison and shows in HTML viewer with optional snapshot save.
    ''' Includes buttons for summarization, PDF export, clipboard copy, Word document export, and optional snapshot save.
    ''' </summary>
    Private Sub CompareTextsAndShowHtmlWithSnapshotOption(originalText As String, revisedText As String, title As String, Optional saveSnapshotCallback As System.Action = Nothing)
        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
        Try
            wordApp = Globals.ThisAddIn.Application
        Catch
            Try
                Dim wordAppObj = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")
                wordApp = TryCast(wordAppObj, Microsoft.Office.Interop.Word.Application)
            Catch
                ShowCustomMessageBox("Microsoft Word is not available.", AN)
                Return
            End Try
        End Try

        If wordApp Is Nothing Then
            ShowCustomMessageBox("Unable to access Word application.", AN)
            Return
        End If

        Dim tempDoc1, tempDoc2, compareDoc As Microsoft.Office.Interop.Word.Document
        Dim tempHtmlPath, tempFolder As String
        Dim prevScreenUpdating = wordApp.ScreenUpdating
        Dim prevAlerts = wordApp.DisplayAlerts

        Try
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

            tempDoc1 = wordApp.Documents.Add(Visible:=False)
            tempDoc1.Content.Text = originalText
            tempDoc2 = wordApp.Documents.Add(Visible:=False)
            tempDoc2.Content.Text = revisedText

            compareDoc = wordApp.CompareDocuments(tempDoc1, tempDoc2, Destination:=WdCompareDestination.wdCompareDestinationNew, Granularity:=WdGranularity.wdGranularityWordLevel, CompareFormatting:=False, CompareCaseChanges:=True, CompareWhitespace:=False, CompareTables:=True, CompareHeaders:=False, CompareFootnotes:=False, CompareTextboxes:=False, CompareFields:=False, CompareComments:=False, CompareMoves:=True, RevisedAuthor:=Environment.UserName, IgnoreAllComparisonWarnings:=True)

            Try : tempDoc1.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            Try : tempDoc2.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try

            If compareDoc Is Nothing Then
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating
                ShowCustomMessageBox("Comparison failed (no result).", AN)
                Return
            End If

            Try
                If compareDoc.Windows IsNot Nothing AndAlso compareDoc.Windows.Count > 0 Then compareDoc.Windows(1).Visible = False
            Catch
            End Try

            Dim extractedChangesText = ExtractChangesWithMarkupTags(compareDoc)
            tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_snapshot_compare_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempFolder)
            tempHtmlPath = Path.Combine(tempFolder, "comparison.htm")
            compareDoc.SaveAs2(FileName:=tempHtmlPath, FileFormat:=WdSaveFormat.wdFormatFilteredHTML)
            Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try

            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating

            ' Capture references for button closures
            Dim capturedWordApp = wordApp
            Dim capturedTempHtmlPath = tempHtmlPath
            Dim capturedTempFolder = tempFolder

            Dim htmlContent = ReadHtmlWithEncodingDetection(capturedTempHtmlPath)
            Dim baseHref = $"<base href=""file:///{capturedTempFolder.Replace("\", "/")}/"">"
            Dim motw = "<!-- saved from url=(0016)http://localhost -->"
            htmlContent = htmlContent.Replace("<head>", "<head>" & vbCrLf & motw & vbCrLf & baseHref)

            Dim additionalButtons As New List(Of System.Tuple(Of String, System.Action, Boolean))()

            ' 1. Summarize Changes (conditional on extracted changes)
            If Not String.IsNullOrWhiteSpace(extractedChangesText) Then
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)("Summarize Changes", Sub() SummarizeComparisonChangesAsync(extractedChangesText), False))
            End If

            ' 2. Send to PDF
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)("Send to PDF", Sub() ExportComparisonToPdfFromHtml(capturedTempHtmlPath, capturedWordApp, "Snapshot_Compare", Environment.GetFolderPath(Environment.SpecialFolder.Desktop)), False))

            ' 3. Copy to Clipboard
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Copy to Clipboard",
                Sub()
                    Try
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        Dim clipDoc As Microsoft.Office.Interop.Word.Document = Nothing
                        Dim prevSU = capturedWordApp.ScreenUpdating
                        Dim prevAL = capturedWordApp.DisplayAlerts
                        Try
                            capturedWordApp.ScreenUpdating = False
                            capturedWordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone
                            clipDoc = capturedWordApp.Documents.Open(FileName:=capturedTempHtmlPath, ReadOnly:=True, Visible:=False)
                            clipDoc.Content.Copy()
                            clipDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                            clipDoc = Nothing
                            ShowCustomMessageBox("Comparison copied to clipboard with formatting.", AN)
                        Finally
                            If clipDoc IsNot Nothing Then Try : clipDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                            capturedWordApp.DisplayAlerts = prevAL
                            capturedWordApp.ScreenUpdating = prevSU
                        End Try
                    Catch ex As Exception
                        ShowCustomMessageBox($"Failed to copy to clipboard: {ex.Message}", AN)
                    End Try
                End Sub,
                False))

            ' 4. Send to Document
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Send to Document",
                Sub()
                    Try
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        Dim newDoc = capturedWordApp.Documents.Open(FileName:=capturedTempHtmlPath, ReadOnly:=False, Visible:=True)
                        newDoc.Activate()
                    Catch ex As Exception
                        ShowCustomMessageBox($"Failed to open in Word: {ex.Message}", AN)
                    End Try
                End Sub,
                False))

            ' 5. Save Snapshot (only when callback provided)
            If saveSnapshotCallback IsNot Nothing Then
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)("Save Snapshot", Sub() saveSnapshotCallback.Invoke(), False))
            End If

            Dim cleanupAction As System.Action = Sub()
                                                     Try
                                                         If Directory.Exists(capturedTempFolder) Then Directory.Delete(capturedTempFolder, recursive:=True)
                                                     Catch
                                                     End Try
                                                 End Sub

            ShowHTMLCustomMessageBox(htmlContent, title, additionalButtons:=additionalButtons.ToArray(), onClose:=cleanupAction)

        Catch ex As Exception
            If compareDoc IsNot Nothing Then Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating
            ShowCustomMessageBox($"Comparison failed: {ex.Message}", AN)
        Finally
            If tempDoc1 IsNot Nothing Then Try : tempDoc1.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            If tempDoc2 IsNot Nothing Then Try : tempDoc2.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
        End Try
    End Sub

#End Region

#Region "Create New Document"

    ''' <summary>
    ''' Creates a new snapshot document entry in the specified library.
    ''' </summary>
    Private Async Function CreateNewSnapshotDocument(targetLibrary As SnapshotLibrary) As System.Threading.Tasks.Task
        If targetLibrary Is Nothing OrElse String.IsNullOrEmpty(targetLibrary.FilePath) Then
            ShowCustomMessageBox("No library is configured.", AN)
            Return
        End If

        ' Determine if this is the local library
        Dim isLocalLibrary As Boolean = False
        If Not String.IsNullOrWhiteSpace(INI_SnapshotLibPathLocal) Then
            Dim expandedLocal = Path.GetFullPath(ExpandEnvironmentVariables(INI_SnapshotLibPathLocal))
            Dim expandedTarget = Path.GetFullPath(targetLibrary.FilePath)
            isLocalLibrary = expandedTarget.Equals(expandedLocal, StringComparison.OrdinalIgnoreCase)
        End If

        Dim params(3) As Slib.InputParameter
        params(0) = New Slib.InputParameter("Friendly Name", "")
        params(1) = New Slib.InputParameter("Location (URL or path)", "")
        params(2) = New Slib.InputParameter("Remove Clutter", True)
        params(3) = New Slib.InputParameter("Archive Path (optional)", "")

        If Not ShowCustomVariableInputForm("Enter document details:", $"{AN} - Add Snapshot Document", params) Then Return

        Dim friendlyName = If(params(0).Value, "").ToString().Trim()
        Dim location = If(params(1).Value, "").ToString().Trim()
        Dim removeClutter = CBool(If(params(2).Value, True))
        Dim archivePath = If(params(3).Value, "").ToString().Trim()

        If String.IsNullOrEmpty(friendlyName) OrElse String.IsNullOrEmpty(location) Then
            ShowCustomMessageBox("Name and Location are required.", AN)
            Return
        End If

        Dim newDoc As New SnapshotDocument() With {
            .FriendlyName = friendlyName,
            .Shortname = GenerateShortname(friendlyName),
            .Location = location,
            .RemoveClutter = removeClutter,
            .SnapshotArchive = archivePath,
            .IsLocal = isLocalLibrary,
            .SourceLibraryPath = targetLibrary.FilePath
        }

        EnsureSnapshotLibraryExists(targetLibrary.FilePath)

        Dim confirmed = ShowCustomYesNoBox(
            $"Configured: {friendlyName}{vbCrLf}Location: {location}{vbCrLf}Library: {If(isLocalLibrary, "Local", "Central")}{vbCrLf}{vbCrLf}Download initial snapshot now?",
            "Yes, download", "No, just add to library")

        AddDocumentToLibrary(newDoc, targetLibrary)

        If confirmed = 1 Then
            Try
                ' Force resolution logic to run exactly as standard logic does
                Dim resolvedArchive = ResolveArchivePath(newDoc, targetLibrary)

                If String.IsNullOrEmpty(resolvedArchive) Then
                    ShowCustomMessageBox("Failed to resolve archive path.", AN)
                    Return
                End If

                If Not Directory.Exists(resolvedArchive) Then Directory.CreateDirectory(resolvedArchive)

                Dim timestamp = DateTime.Now
                Dim splash As New Slib.SplashScreen($"Downloading: {friendlyName}...")
                splash.Show()
                splash.Refresh()
                System.Windows.Forms.Application.DoEvents()

                Dim content = Await RetrieveDocumentContentAsync(location)
                splash.Close()

                If String.IsNullOrEmpty(content) Then
                    ShowCustomMessageBox("Failed to retrieve content.", AN)
                    Return
                End If

                If removeClutter Then
                    Try
                        Dim prompt = "<TEXTTOPROCESS>" & vbCrLf & content & vbCrLf & "</TEXTTOPROCESS>"
                        Dim cleaned = Await LLM(SP_RemoveClutter, prompt, "", "", 0, False, False)
                        If Not String.IsNullOrWhiteSpace(cleaned) AndAlso Not cleaned.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then
                            content = cleaned
                        End If
                    Catch
                    End Try
                End If

                Dim snapshotFullPath = SaveSnapshotWithTimestamp(newDoc, content, resolvedArchive, timestamp)
                If Not String.IsNullOrEmpty(snapshotFullPath) Then
                    Dim fn = Path.GetFileName(snapshotFullPath)
                    AddSnapshotToLibrary(newDoc, fn, targetLibrary)
                    ShowCustomMessageBox($"Snapshot saved:{vbCrLf}{snapshotFullPath}", AN)
                Else
                    ShowCustomMessageBox("Failed to save snapshot file.", AN)
                End If

            Catch ex As Exception
                ShowCustomMessageBox($"Error taking initial snapshot:{vbCrLf}{ex.Message}", AN)
            End Try
        End If
    End Function

#End Region

#Region "Snapshot Management"

    Private Sub ViewSnapshots(allDocuments As List(Of SnapshotDocument), localLibrary As SnapshotLibrary, centralLibrary As SnapshotLibrary)
        ' Implementation identical to original, abbreviated for brevity in this response
        ' Assuming original logic handles selection and calls ViewDocumentSnapshots

        ' Reusing existing ViewSnapshots implementation structure but ensuring connection to improved ResolveArchivePath
        ' First, select which document to view snapshots for
        Dim items As New List(Of SelectionItem)()
        Dim docMap As New Dictionary(Of Integer, SnapshotDocument)()

        Dim docsWithSnapshots = allDocuments.Where(Function(d As SnapshotDocument) d.Snapshots.Count > 0).ToList()

        If docsWithSnapshots.Count = 0 Then
            ShowCustomMessageBox("No documents have snapshots to view.", AN)
            Return
        End If

        Dim idx = 1
        For Each doc As SnapshotDocument In docsWithSnapshots
            items.Add(New SelectionItem($"{doc.GetDisplayName()} ({doc.Snapshots.Count} snapshots)", idx))
            docMap(idx) = doc
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select document to view snapshots:", $"{AN} - View Snapshots")

        If result <= 0 OrElse Not docMap.ContainsKey(result) Then
            Return
        End If

        Dim selectedDoc = docMap(result)
        Dim docLibrary = GetLibraryForDocument(selectedDoc, localLibrary, centralLibrary)
        ViewDocumentSnapshots(selectedDoc, docLibrary)
    End Sub

    ''' <summary>
    ''' Allows the user to pick a document, then select two existing snapshots to compare.
    ''' Only documents with at least 2 snapshots are shown.
    ''' </summary>
    Private Sub CompareExistingSnapshots(allDocuments As List(Of SnapshotDocument), localLibrary As SnapshotLibrary, centralLibrary As SnapshotLibrary)
        ' Filter to documents with at least 2 snapshots
        Dim docsWithMultiple = allDocuments.Where(Function(d As SnapshotDocument) d.Snapshots.Count >= 2).ToList()

        If docsWithMultiple.Count = 0 Then
            ShowCustomMessageBox("No documents have 2 or more snapshots to compare.", AN)
            Return
        End If

        ' Select document
        Dim items As New List(Of SelectionItem)()
        Dim docMap As New Dictionary(Of Integer, SnapshotDocument)()

        Dim idx = 1
        For Each doc As SnapshotDocument In docsWithMultiple
            items.Add(New SelectionItem($"{doc.GetDisplayName()} ({doc.Snapshots.Count} snapshots)", idx))
            docMap(idx) = doc
            idx += 1
        Next

        Dim docResult = SelectValue(items, 1, "Select document to compare snapshots:", $"{AN} - Compare Snapshots")
        If docResult <= 0 OrElse Not docMap.ContainsKey(docResult) Then Return

        Dim selectedDoc = docMap(docResult)
        Dim docLibrary = GetLibraryForDocument(selectedDoc, localLibrary, centralLibrary)
        Dim archivePath = ResolveArchivePath(selectedDoc, docLibrary)

        If String.IsNullOrEmpty(archivePath) Then
            ShowCustomMessageBox("Could not determine snapshot archive path.", AN)
            Return
        End If

        ' Build snapshot list (sorted newest first)
        Dim sortedSnapshots = selectedDoc.Snapshots.OrderByDescending(Function(s) s).ToList()

        ' Select first (older/base) snapshot - default to second-most-recent
        Dim firstSnapshot = SelectSnapshotFromList(sortedSnapshots, archivePath, "Select OLDER snapshot (base):",
                                                    $"{AN} - Compare Snapshots (1/2)",
                                                    If(sortedSnapshots.Count >= 2, 2, 1))
        If String.IsNullOrEmpty(firstSnapshot) Then Return

        ' Select second (newer/revised) snapshot - default to most-recent, exclude the first selection
        Dim remainingSnapshots = sortedSnapshots.Where(Function(s) Not s.Equals(firstSnapshot, StringComparison.OrdinalIgnoreCase)).ToList()

        If remainingSnapshots.Count = 0 Then
            ShowCustomMessageBox("No other snapshot available to compare against.", AN)
            Return
        End If

        Dim secondSnapshot = SelectSnapshotFromList(remainingSnapshots, archivePath, "Select NEWER snapshot (revised):",
                                                     $"{AN} - Compare Snapshots (2/2)", 1)
        If String.IsNullOrEmpty(secondSnapshot) Then Return

        ' Read both files
        Dim firstPath = Path.Combine(archivePath, firstSnapshot)
        Dim secondPath = Path.Combine(archivePath, secondSnapshot)

        If Not File.Exists(firstPath) Then
            ShowCustomMessageBox($"Snapshot file not found:{vbCrLf}{firstPath}", AN)
            Return
        End If
        If Not File.Exists(secondPath) Then
            ShowCustomMessageBox($"Snapshot file not found:{vbCrLf}{secondPath}", AN)
            Return
        End If

        Dim firstContent = File.ReadAllText(firstPath, Encoding.UTF8)
        Dim secondContent = File.ReadAllText(secondPath, Encoding.UTF8)

        ' Check for identical content
        If ComputeContentHash(firstContent) = ComputeContentHash(secondContent) Then
            ShowCustomMessageBox("The two snapshots are identical. No differences to show.", AN)
            Return
        End If

        ' Compare without save callback (no Save Snapshot button)
        CompareTextsAndShowHtmlWithSnapshotOption(firstContent, secondContent,
            $"{AN} Snapshot Compare - {selectedDoc.FriendlyName}")
    End Sub

    ''' <summary>
    ''' Shows a snapshot selection dialog from a given list with customizable prompt, title, and default selection.
    ''' </summary>
    Private Function SelectSnapshotFromList(snapshots As List(Of String), archivePath As String, prompt As String, title As String, defaultIndex As Integer) As String
        Dim items As New List(Of SelectionItem)()
        Dim snapshotMap As New Dictionary(Of Integer, String)()

        Dim idx = 1
        For Each snapshot In snapshots
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshot)
            Dim displayName = If(snapshotDate.HasValue, $"{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}", snapshot)
            Dim fullPath = Path.Combine(archivePath, snapshot)
            If File.Exists(fullPath) Then
                Dim fileInfo = New FileInfo(fullPath)
                displayName &= $" ({FormatFileSize(fileInfo.Length)})"
            Else
                displayName &= " [FILE MISSING]"
            End If
            items.Add(New SelectionItem(displayName, idx))
            snapshotMap(idx) = snapshot
            idx += 1
        Next

        Dim result = SelectValue(items, Math.Min(defaultIndex, items.Count), prompt, title)
        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then Return snapshotMap(result)
        Return Nothing
    End Function

    Private Sub ViewDocumentSnapshots(doc As SnapshotDocument, library As SnapshotLibrary)
        Dim archivePath = ResolveArchivePath(doc, library)
        If String.IsNullOrEmpty(archivePath) Then
            ShowCustomMessageBox("Could not determine snapshot archive path.", AN)
            Return
        End If

        Dim items As New List(Of SelectionItem)()
        Dim snapshotMap As New Dictionary(Of Integer, String)()
        Dim sortedSnapshots = doc.Snapshots.OrderByDescending(Function(s) s).ToList()

        Dim idx = 1
        For Each snapshot In sortedSnapshots
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshot)
            Dim displayName = If(snapshotDate.HasValue, $"{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}", snapshot)
            Dim fullPath = Path.Combine(archivePath, snapshot)
            If File.Exists(fullPath) Then
                Dim fileInfo = New FileInfo(fullPath)
                displayName &= $" ({FormatFileSize(fileInfo.Length)})"
            Else
                displayName &= " [FILE MISSING]"
            End If
            items.Add(New SelectionItem(displayName, idx))
            snapshotMap(idx) = snapshot
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select snapshot to view:", $"{AN} - View Snapshots")
        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then
            Dim snapshotToView = snapshotMap(result)
            Dim fullPath = Path.Combine(archivePath, snapshotToView)
            If File.Exists(fullPath) Then
                ShowTextFileEditor(fullPath, $"{AN} - {doc.FriendlyName}", True, _context)
                ViewDocumentSnapshots(doc, library)
            Else
                ShowCustomMessageBox("File not found.", AN)
            End If
        End If
    End Sub

    Private Sub ManageSnapshots(allDocuments As List(Of SnapshotDocument), localLibrary As SnapshotLibrary, centralLibrary As SnapshotLibrary)
        Dim items As New List(Of SelectionItem)()
        Dim docMap As New Dictionary(Of Integer, SnapshotDocument)()
        Dim docsWithSnapshots = allDocuments.Where(Function(d As SnapshotDocument) d.Snapshots.Count > 0).ToList()

        If docsWithSnapshots.Count = 0 Then
            ShowCustomMessageBox("No documents have snapshots to manage.", AN)
            Return
        End If

        Dim idx = 1
        For Each doc As SnapshotDocument In docsWithSnapshots
            items.Add(New SelectionItem($"{doc.GetDisplayName()} ({doc.Snapshots.Count} snapshots)", idx))
            docMap(idx) = doc
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select document to manage:", $"{AN} - Manage Snapshots")
        If result > 0 AndAlso docMap.ContainsKey(result) Then
            Dim selectedDoc = docMap(result)
            Dim docLibrary = GetLibraryForDocument(selectedDoc, localLibrary, centralLibrary)
            ManageDocumentSnapshots(selectedDoc, docLibrary)
        End If
    End Sub

    Private Sub ManageDocumentSnapshots(doc As SnapshotDocument, library As SnapshotLibrary)
        Dim archivePath = ResolveArchivePath(doc, library)
        If String.IsNullOrEmpty(archivePath) Then
            ShowCustomMessageBox("Could not determine snapshot archive path.", AN)
            Return
        End If

        Dim items As New List(Of SelectionItem)()
        Dim snapshotMap As New Dictionary(Of Integer, String)()
        Dim sortedSnapshots = doc.Snapshots.OrderByDescending(Function(s) s).ToList()

        Dim idx = 1
        For Each snapshot In sortedSnapshots
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshot)
            Dim displayName = If(snapshotDate.HasValue, $"{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}", snapshot)
            Dim fullPath = Path.Combine(archivePath, snapshot)
            If File.Exists(fullPath) Then
                Dim fileInfo = New FileInfo(fullPath)
                displayName &= $" ({FormatFileSize(fileInfo.Length)})"
            Else
                displayName &= " [FILE MISSING]"
            End If
            items.Add(New SelectionItem(displayName, idx))
            snapshotMap(idx) = snapshot
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select snapshot to delete:", $"{AN} - Manage Snapshots")
        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then
            Dim snapshotToDelete = snapshotMap(result)
            Dim fullPath = Path.Combine(archivePath, snapshotToDelete)
            If ShowCustomYesNoBox($"Delete this snapshot?{vbCrLf}{fullPath}", "Yes", "No") = 1 Then
                Try
                    If File.Exists(fullPath) Then File.Delete(fullPath)
                    RemoveSnapshotFromLibrary(doc, snapshotToDelete, library)
                    doc.Snapshots.Remove(snapshotToDelete)
                    If doc.Snapshots.Count > 0 Then ManageDocumentSnapshots(doc, library)
                Catch ex As Exception
                    ShowCustomMessageBox($"Failed to delete: {ex.Message}", AN)
                End Try
            End If
        End If
    End Sub

#End Region

#Region "Library Parsing and Writing"

    ' Implementation identical to original, abbreviated for brevity but keeping all helper methods
    Private Function ParseSnapshotLibrary(filePath As String, isLocal As Boolean) As SnapshotLibrary
        Dim library As New SnapshotLibrary() With {.FilePath = filePath}
        If Not File.Exists(filePath) Then Return library
        Try
            Dim lines = File.ReadAllLines(filePath, Encoding.UTF8)
            Dim currentDoc As SnapshotDocument = Nothing
            For Each line As String In lines
                Dim trimmedLine = line.Trim()
                If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then Continue For
                If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                    If currentDoc IsNot Nothing Then library.Documents.Add(currentDoc)
                    Dim docName = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim()
                    currentDoc = New SnapshotDocument() With {.FriendlyName = docName, .IsLocal = isLocal, .SourceLibraryPath = filePath}
                    Continue For
                End If
                Dim eqIndex = trimmedLine.IndexOf("="c)
                If eqIndex > 0 Then
                    Dim key = trimmedLine.Substring(0, eqIndex).Trim().ToLowerInvariant()
                    Dim value = trimmedLine.Substring(eqIndex + 1).Trim()
                    If currentDoc Is Nothing Then
                        If key = "snapshotarchive" Then library.GlobalSnapshotArchive = value
                    Else
                        Select Case key
                            Case "shortname" : currentDoc.Shortname = value
                            Case "location" : currentDoc.Location = value
                            Case "removeclutter" : currentDoc.RemoveClutter = (value.ToLower() = "true" OrElse value = "1" OrElse value.ToLower() = "yes")
                            Case "snapshotarchive" : currentDoc.SnapshotArchive = value
                            Case "snapshot" : currentDoc.Snapshots.Add(value)
                        End Select
                    End If
                End If
            Next
            If currentDoc IsNot Nothing Then library.Documents.Add(currentDoc)
            For Each doc As SnapshotDocument In library.Documents
                If String.IsNullOrEmpty(doc.Shortname) Then doc.Shortname = GenerateShortname(doc.FriendlyName)
            Next
        Catch ex As Exception
            Debug.WriteLine($"Failed to parse snapshot library: {ex.Message}")
        End Try
        Return library
    End Function

    Private Sub EnsureSnapshotLibraryExists(filePath As String)
        If File.Exists(filePath) Then Return
        Try
            Dim dir = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
            Dim content = "; Snapshot Library" & vbCrLf & $"[{DateTime.Now}]" & vbCrLf & "SnapshotArchive = SnapshotArchive" & vbCrLf
            File.WriteAllText(filePath, content, Encoding.UTF8)
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to create library file: {ex.Message}", AN)
        End Try
    End Sub

    Private Sub AddDocumentToLibrary(doc As SnapshotDocument, library As SnapshotLibrary)
        Try
            EnsureSnapshotLibraryExists(library.FilePath)
            Dim sb As New StringBuilder()
            sb.AppendLine()
            sb.AppendLine($"[{doc.FriendlyName}]")
            sb.AppendLine($"Shortname = {doc.Shortname}")
            sb.AppendLine($"Location = {doc.Location}")
            sb.AppendLine($"RemoveClutter = {If(doc.RemoveClutter, "True", "False")}")
            If Not String.IsNullOrEmpty(doc.SnapshotArchive) Then sb.AppendLine($"SnapshotArchive = {doc.SnapshotArchive}")
            File.AppendAllText(library.FilePath, sb.ToString(), Encoding.UTF8)
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to add document: {ex.Message}", AN)
        End Try
    End Sub

    Private Sub AddSnapshotToLibrary(doc As SnapshotDocument, snapshotFilename As String, library As SnapshotLibrary)
        Try
            If library Is Nothing OrElse String.IsNullOrEmpty(library.FilePath) OrElse Not File.Exists(library.FilePath) Then Return
            Dim lines = File.ReadAllLines(library.FilePath, Encoding.UTF8).ToList()
            Dim inTargetSection = False
            Dim insertIndex = -1
            For i = 0 To lines.Count - 1
                Dim line = lines(i).Trim()
                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    If inTargetSection Then
                        insertIndex = i
                        Exit For
                    End If
                    Dim sectionName = line.Substring(1, line.Length - 2).Trim()
                    If sectionName.Equals(doc.FriendlyName, StringComparison.OrdinalIgnoreCase) Then inTargetSection = True
                End If
            Next
            If inTargetSection AndAlso insertIndex = -1 Then insertIndex = lines.Count
            If insertIndex >= 0 Then
                lines.Insert(insertIndex, $"Snapshot = {snapshotFilename}")
                File.WriteAllLines(library.FilePath, lines, Encoding.UTF8)
            End If
        Catch ex As Exception
            Debug.WriteLine($"Failed to add snapshot: {ex.Message}")
        End Try
    End Sub

    Private Sub RemoveSnapshotFromLibrary(doc As SnapshotDocument, snapshotFilename As String, library As SnapshotLibrary)
        Try
            If library Is Nothing OrElse String.IsNullOrEmpty(library.FilePath) OrElse Not File.Exists(library.FilePath) Then Return
            Dim lines = File.ReadAllLines(library.FilePath, Encoding.UTF8).ToList()
            Dim lineToRemove = $"Snapshot = {snapshotFilename}"
            Dim inTargetSection = False
            Dim removed = False
            For i = 0 To lines.Count - 1
                Dim line = lines(i).Trim()
                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    Dim sectionName = line.Substring(1, line.Length - 2).Trim()
                    inTargetSection = sectionName.Equals(doc.FriendlyName, StringComparison.OrdinalIgnoreCase)
                ElseIf inTargetSection AndAlso line.Equals(lineToRemove, StringComparison.OrdinalIgnoreCase) Then
                    lines.RemoveAt(i)
                    removed = True
                    Exit For
                End If
            Next
            If removed Then File.WriteAllLines(library.FilePath, lines, Encoding.UTF8)
        Catch ex As Exception
            Debug.WriteLine($"Failed to remove snapshot: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "Helper Methods"

    Private Function GetLibraryForDocument(doc As SnapshotDocument, localLibrary As SnapshotLibrary, centralLibrary As SnapshotLibrary) As SnapshotLibrary
        If Not String.IsNullOrEmpty(doc.SourceLibraryPath) Then
            If localLibrary IsNot Nothing AndAlso localLibrary.FilePath.Equals(doc.SourceLibraryPath, StringComparison.OrdinalIgnoreCase) Then Return localLibrary
            If centralLibrary IsNot Nothing AndAlso centralLibrary.FilePath.Equals(doc.SourceLibraryPath, StringComparison.OrdinalIgnoreCase) Then Return centralLibrary
        End If
        Return If(doc.IsLocal, localLibrary, centralLibrary)
    End Function

    Private Async Function RetrieveDocumentContentAsync(location As String) As Task(Of String)
        If String.IsNullOrEmpty(location) Then Return Nothing
        Try
            If location.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
                Return Await RetrieveWebsiteContent_WebView2(location)
            Else
                If File.Exists(location) Then
                    Return Await GetFileContent(location, True, SharedMethods.IsOcrAvailable(_context), False)
                End If
            End If
        Catch
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Resolves the archive path for a document.
    ''' STRICTLY uses INI global paths to ensure correct root.
    ''' </summary>
    Private Function ResolveArchivePath(doc As SnapshotDocument, library As SnapshotLibrary) As String
        ' 1. Trust the IsLocal flag to pick the authoritative source key
        Dim baseVar As String
        If doc.IsLocal Then
            baseVar = INI_SnapshotLibPathLocal
        Else
            baseVar = INI_SnapshotLibPath
        End If

        ' Only fallback to doc source if global var is completely missing (unlikely)
        If String.IsNullOrWhiteSpace(baseVar) Then
            baseVar = doc.SourceLibraryPath
        End If

        If String.IsNullOrWhiteSpace(baseVar) Then
            Return Nothing
        End If

        ' 2. Start from the library file path
        Dim libraryFilePath As String = ExpandEnvironmentVariables(baseVar)

        libraryFilePath = Path.GetFullPath(libraryFilePath)

        ' 3. Get the directory of that library file
        Dim baseDir = Path.GetDirectoryName(libraryFilePath)
        If String.IsNullOrWhiteSpace(baseDir) Then Return Nothing

        ' 4. Determine archive name
        Dim archivePath = doc.SnapshotArchive
        If String.IsNullOrWhiteSpace(archivePath) AndAlso library IsNot Nothing Then
            archivePath = library.GlobalSnapshotArchive
        End If
        If String.IsNullOrWhiteSpace(archivePath) Then
            archivePath = "SnapshotArchive"
        End If

        archivePath = ExpandEnvironmentVariables(archivePath)

        ' 5. Combine archive path (if relative) with the base directory
        If Not Path.IsPathRooted(archivePath) Then
            archivePath = Path.GetFullPath(Path.Combine(baseDir, archivePath))
        End If

        Return archivePath
    End Function

    Private Function SaveSnapshotWithTimestamp(doc As SnapshotDocument, content As String, archivePath As String, timestamp As DateTime) As String
        Try
            If Not Directory.Exists(archivePath) Then Directory.CreateDirectory(archivePath)
            Dim timestampStr = timestamp.ToString("yyMMdd_HHmmss")
            Dim filename = $"{timestampStr}_{doc.Shortname}.txt"
            Dim fullPath = Path.Combine(archivePath, filename)
            If File.Exists(fullPath) Then
                ShowCustomMessageBox($"Snapshot exists: {fullPath}", AN)
                Return Nothing
            End If
            Dim normalizedContent = content
            If Not String.IsNullOrEmpty(normalizedContent) Then
                normalizedContent = normalizedContent.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            End If
            File.WriteAllText(fullPath, normalizedContent, Encoding.UTF8)
            Return fullPath
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to save: {ex.Message}", AN)
            Return Nothing
        End Try
    End Function

    Private Function SaveSnapshot(doc As SnapshotDocument, content As String, archivePath As String) As String
        Return SaveSnapshotWithTimestamp(doc, content, archivePath, DateTime.Now)
    End Function

    Private Function GenerateShortname(friendlyName As String) As String
        If String.IsNullOrEmpty(friendlyName) Then Return "document"
        Dim invalid = Path.GetInvalidFileNameChars()
        Dim result = friendlyName
        For Each c In invalid : result = result.Replace(c, "_"c) : Next
        result = result.Replace(" "c, "_"c)
        While result.Contains("__") : result = result.Replace("__", "_") : End While
        result = result.Trim("_"c)
        If result.Length > 40 Then result = result.Substring(0, 40)
        If String.IsNullOrEmpty(result) Then result = "document"
        Return result.ToLowerInvariant()
    End Function

    Private Function ComputeContentHash(content As String) As String
        If String.IsNullOrEmpty(content) Then Return ""
        Using hasher As SHA256 = SHA256.Create()
            Dim bytes = Encoding.UTF8.GetBytes(content)
            Dim hash = hasher.ComputeHash(bytes)
            Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    Private Function FormatFileSize(bytes As Long) As String
        If bytes < 1024 Then Return $"{bytes} B"
        If bytes < 1024 * 1024 Then Return $"{bytes / 1024.0:F1} KB"
        Return $"{bytes / (1024.0 * 1024.0):F1} MB"
    End Function

#End Region

End Class