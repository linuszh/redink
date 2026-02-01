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

        ' Check if any library path is configured
        Dim hascentral As Boolean = Not String.IsNullOrWhiteSpace(centralLibPath) AndAlso File.Exists(centralLibPath)
        Dim hasLocal As Boolean = Not String.IsNullOrWhiteSpace(localLibPath)

        If Not hascentral AndAlso Not hasLocal Then
            ShowCustomMessageBox("No snapshot library is configured. Please set the parameters 'SnapshotLibPath' or 'SnapshotLibPathLocal' in your configuration file.", AN)
            Return
        End If

        ' Load libraries
        Dim allDocuments As New List(Of SnapshotDocument)()
        Dim localLibrary As SnapshotLibrary = Nothing
        Dim centralLibrary As SnapshotLibrary = Nothing

        ' Load local library first (takes precedence in display order)
        If hasLocal Then
            If File.Exists(localLibPath) Then
                localLibrary = ParseSnapshotLibrary(localLibPath, isLocal:=True)
                allDocuments.AddRange(localLibrary.Documents)
            Else
                ' Local path configured but file doesn't exist yet - that's OK, we can create it
                localLibrary = New SnapshotLibrary() With {.FilePath = localLibPath}
            End If
        End If

        ' Load central library
        If hascentral Then
            centralLibrary = ParseSnapshotLibrary(centralLibPath, isLocal:=False)
            ' Add central documents, but mark duplicates
            For Each doc As SnapshotDocument In centralLibrary.Documents
                ' Check if already exists in local
                Dim existsInLocal = allDocuments.Any(Function(d As SnapshotDocument) d.FriendlyName.Equals(doc.FriendlyName, StringComparison.OrdinalIgnoreCase))
                If Not existsInLocal Then
                    allDocuments.Add(doc)
                End If
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

        ' Add management options (no separator - direct options)
        Const ViewSnapshotsValue As Integer = -4
        Const DeleteSnapshotsValue As Integer = -3
        Const AddNewValue As Integer = -1
        Const EditLibraryValue As Integer = -2

        If allDocuments.Count > 0 Then
            items.Add(New SelectionItem("View snapshots...", ViewSnapshotsValue))
            items.Add(New SelectionItem("Delete snapshots...", DeleteSnapshotsValue))
        End If

        If localLibrary IsNot Nothing Then
            items.Add(New SelectionItem("Add new document...", AddNewValue))
            items.Add(New SelectionItem("Edit local snapshot library...", EditLibraryValue))
        End If

        If items.Count = 0 Then
            ShowCustomMessageBox("No snapshot documents configured and no local library available to add new ones.", AN)
            Return
        End If

        Dim result = SelectValue(items, If(sortedDocs.Count > 0, 1, AddNewValue),
                                 "Select a document to snapshot or compare:",
                                 $"{AN} - Document Snapshots")

        If result = 0 Then
            ' Cancelled
            Return
        ElseIf result = ViewSnapshotsValue Then
            ' View snapshots
            ViewSnapshots(allDocuments, localLibrary, centralLibrary)
        ElseIf result = DeleteSnapshotsValue Then
            ' Delete snapshots (formerly Manage snapshots)
            ManageSnapshots(allDocuments, localLibrary, centralLibrary)
        ElseIf result = AddNewValue Then
            ' Add new document
            Await CreateNewSnapshotDocument(localLibrary)
            ' Refresh the list
            SelectSnapshotDocument()
        ElseIf result = EditLibraryValue Then
            ' Edit local library
            If localLibrary IsNot Nothing AndAlso Not String.IsNullOrEmpty(localLibrary.FilePath) Then
                ' Ensure file exists
                If Not File.Exists(localLibrary.FilePath) Then
                    EnsureSnapshotLibraryExists(localLibrary.FilePath)
                End If
                ShowTextFileEditor(localLibrary.FilePath, $"{AN} - Edit Snapshot Library", False, _context)
            End If
            ' Refresh the list
            SelectSnapshotDocument()
        ElseIf docMap.ContainsKey(result) Then
            ' Document selected - process snapshot
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
            ' Determine archive path
            Dim archivePath = ResolveArchivePath(doc, If(doc.IsLocal, localLibrary, centralLibrary))
            If String.IsNullOrEmpty(archivePath) Then
                ShowCustomMessageBox("Could not determine snapshot archive path.", AN)
                Return
            End If

            ' Ensure archive directory exists
            If Not Directory.Exists(archivePath) Then
                Try
                    Directory.CreateDirectory(archivePath)
                Catch ex As Exception
                    ShowCustomMessageBox($"Failed to create archive directory: {ex.Message}", AN)
                    Return
                End Try
            End If

            ' Capture download timestamp BEFORE downloading (used for snapshot filename)
            Dim downloadTimestamp As DateTime = DateTime.Now

            ' Show splash screen for download
            Dim splash As New Slib.SplashScreen($"Downloading: {doc.FriendlyName}...")
            splash.Show()
            splash.Refresh()
            System.Windows.Forms.Application.DoEvents()

            ' Download new content
            Dim newContent As String = Nothing
            Dim rawContent As String = Nothing

            Try
                rawContent = Await RetrieveDocumentContentAsync(doc.Location)
                If String.IsNullOrEmpty(rawContent) Then
                    splash.Close()
                    ShowCustomMessageBox($"Failed to retrieve content from: {doc.Location}", AN)
                    Return
                End If
            Catch ex As Exception
                splash.Close()
                ShowCustomMessageBox($"Error retrieving document: {ex.Message}", AN)
                Return
            End Try

            ' Close download splash before LLM call
            splash.Close()

            ' Optionally remove clutter
            If doc.RemoveClutter Then
                Try
                    Dim userPrompt = "<TEXTTOPROCESS>" & vbCrLf & rawContent & vbCrLf & "</TEXTTOPROCESS>"
                    ' Let LLM show its own splash (HideSplash:=False)
                    newContent = Await LLM(
                        SP_RemoveClutter,
                        userPrompt,
                        "",
                        "",
                        0,
                        False,
                        False)

                    If String.IsNullOrWhiteSpace(newContent) OrElse newContent.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then
                        ' LLM failed - offer to use raw content
                        Dim useRaw = ShowCustomYesNoBox(
                            $"Clutter removal failed: {If(newContent, "Empty result")}{vbCrLf}{vbCrLf}Do you want to continue with the raw content instead?",
                            "Yes, use raw content",
                            "No, cancel")
                        If useRaw = 1 Then
                            newContent = rawContent
                        Else
                            Return
                        End If
                    End If
                Catch ex As Exception
                    Dim useRaw = ShowCustomYesNoBox(
                        $"Clutter removal error: {ex.Message}{vbCrLf}{vbCrLf}Do you want to continue with the raw content instead?",
                        "Yes, use raw content",
                        "No, cancel")
                    If useRaw = 1 Then
                        newContent = rawContent
                    Else
                        Return
                    End If
                End Try
            Else
                newContent = rawContent
            End If

            ' Check if we have existing snapshots to compare
            If doc.Snapshots.Count = 0 Then
                ' No existing snapshots - just save the new one
                Dim snapshotFilename = SaveSnapshotWithTimestamp(doc, newContent, archivePath, downloadTimestamp)
                If Not String.IsNullOrEmpty(snapshotFilename) Then
                    ' Update library file
                    AddSnapshotToLibrary(doc, snapshotFilename, If(doc.IsLocal, localLibrary, centralLibrary))
                    ShowCustomMessageBox($"Initial snapshot saved: {snapshotFilename}", AN)
                End If
                Return
            End If

            ' We have existing snapshots - let user select which to compare against
            Dim previousSnapshot = SelectSnapshotForComparison(doc, archivePath)
            If String.IsNullOrEmpty(previousSnapshot) Then
                Return ' User cancelled
            End If

            ' Load previous snapshot content
            Dim previousPath = Path.Combine(archivePath, previousSnapshot)
            If Not File.Exists(previousPath) Then
                ShowCustomMessageBox($"Previous snapshot file not found: {previousSnapshot}", AN)
                Return
            End If

            Dim previousContent = File.ReadAllText(previousPath, Encoding.UTF8)

            ' Check if content is identical using hash
            Dim previousHash = ComputeContentHash(previousContent)
            Dim newHash = ComputeContentHash(newContent)

            If previousHash = newHash Then
                ShowCustomMessageBox("No changes detected. The current content is identical to the selected snapshot.", AN)
                Return
            End If

            ' Show comparison with save option - pass the download timestamp for saving
            CompareAndShowSnapshots(doc, previousContent, newContent, archivePath, localLibrary, centralLibrary, downloadTimestamp)

        Catch ex As Exception
            ShowCustomMessageBox($"Error processing snapshot: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Allows user to select which snapshot to compare against.
    ''' </summary>
    Private Function SelectSnapshotForComparison(doc As SnapshotDocument, archivePath As String) As String
        If doc.Snapshots.Count = 1 Then
            Return doc.Snapshots(0)
        End If

        ' Build selection items - newest first
        Dim items As New List(Of SelectionItem)()
        Dim snapshotMap As New Dictionary(Of Integer, String)()

        Dim sortedSnapshots = doc.Snapshots.OrderByDescending(Function(s) s).ToList()

        Dim idx = 1
        For Each snapshot In sortedSnapshots
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshot)
            Dim displayName = If(snapshotDate.HasValue,
                                 $"{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}",
                                 snapshot)

            ' Check if file exists and get size
            Dim fullPath = Path.Combine(archivePath, snapshot)
            If File.Exists(fullPath) Then
                Dim fileInfo = New FileInfo(fullPath)
                displayName &= $" ({FormatFileSize(fileInfo.Length)})"
            End If

            items.Add(New SelectionItem(displayName, idx))
            snapshotMap(idx) = snapshot
            idx += 1
        Next

        Dim result = SelectValue(items, 1,
                                 "Select the snapshot to compare against (newest shown first):",
                                 $"{AN} - Select Snapshot")

        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then
            Return snapshotMap(result)
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Compares two snapshot contents and shows the result with save option.
    ''' </summary>
    Private Sub CompareAndShowSnapshots(
        doc As SnapshotDocument,
        previousContent As String,
        newContent As String,
        archivePath As String,
        localLibrary As SnapshotLibrary,
        centralLibrary As SnapshotLibrary,
        downloadTimestamp As DateTime)

        ' Track whether snapshot has been saved to prevent duplicate saves
        Dim snapshotSaved As Boolean = False

        ' Use Word comparison via the existing helper
        CompareTextsAndShowHtmlWithSnapshotOption(
            previousContent,
            newContent,
            $"{AN} Snapshot Compare - {doc.FriendlyName}",
            Sub()
                ' Save snapshot callback - use the download timestamp
                If snapshotSaved Then
                    ShowCustomMessageBox("This snapshot has already been saved.", AN)
                    Return
                End If

                Dim snapshotFilename = SaveSnapshotWithTimestamp(doc, newContent, archivePath, downloadTimestamp)
                If Not String.IsNullOrEmpty(snapshotFilename) Then
                    AddSnapshotToLibrary(doc, snapshotFilename, If(doc.IsLocal, localLibrary, centralLibrary))
                    ShowCustomMessageBox($"Snapshot saved: {snapshotFilename}", AN)
                    snapshotSaved = True
                End If
            End Sub)
    End Sub

    ''' <summary>
    ''' Compares two text strings using Word's comparison and shows in HTML viewer with snapshot save option.
    ''' </summary>
    Private Sub CompareTextsAndShowHtmlWithSnapshotOption(
        originalText As String,
        revisedText As String,
        title As String,
        saveSnapshotCallback As System.Action)

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

        Dim tempDoc1 As Microsoft.Office.Interop.Word.Document = Nothing
        Dim tempDoc2 As Microsoft.Office.Interop.Word.Document = Nothing
        Dim compareDoc As Microsoft.Office.Interop.Word.Document = Nothing
        Dim tempHtmlPath As String = Nothing
        Dim tempFolder As String = Nothing

        Dim prevScreenUpdating As Boolean = wordApp.ScreenUpdating
        Dim prevAlerts As Microsoft.Office.Interop.Word.WdAlertLevel = wordApp.DisplayAlerts

        Try
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone

            ' Create temporary documents
            tempDoc1 = wordApp.Documents.Add(Visible:=False)
            tempDoc1.Content.Text = originalText

            tempDoc2 = wordApp.Documents.Add(Visible:=False)
            tempDoc2.Content.Text = revisedText

            ' Compare
            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=tempDoc1,
                RevisedDocument:=tempDoc2,
                Destination:=WdCompareDestination.wdCompareDestinationNew,
                Granularity:=WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=False,
                CompareCaseChanges:=True,
                CompareWhitespace:=False,
                CompareTables:=True,
                CompareHeaders:=False,
                CompareFootnotes:=False,
                CompareTextboxes:=False,
                CompareFields:=False,
                CompareComments:=False,
                CompareMoves:=True,
                RevisedAuthor:=Environment.UserName,
                IgnoreAllComparisonWarnings:=True)

            ' Close temp documents
            Try : tempDoc1.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            tempDoc1 = Nothing
            Try : tempDoc2.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            tempDoc2 = Nothing

            If compareDoc Is Nothing Then
                wordApp.DisplayAlerts = prevAlerts
                wordApp.ScreenUpdating = prevScreenUpdating
                ShowCustomMessageBox("Word did not produce a comparison result.", AN)
                Return
            End If

            ' Hide comparison window
            Try
                If compareDoc.Windows IsNot Nothing AndAlso compareDoc.Windows.Count > 0 Then
                    compareDoc.Windows(1).Visible = False
                End If
            Catch
            End Try

            ' Extract changes for summarization
            Dim extractedChangesText = ExtractChangesWithMarkupTags(compareDoc)

            ' Export to HTML
            tempFolder = Path.Combine(Path.GetTempPath(), $"{AN2}_snapshot_compare_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempFolder)
            tempHtmlPath = Path.Combine(tempFolder, "comparison.htm")

            compareDoc.SaveAs2(FileName:=tempHtmlPath, FileFormat:=WdSaveFormat.wdFormatFilteredHTML)

            ' Close comparison doc
            Try : compareDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
            compareDoc = Nothing

            ' Restore UI BEFORE showing the HTML dialog
            wordApp.DisplayAlerts = prevAlerts
            wordApp.ScreenUpdating = prevScreenUpdating

            ' Read HTML
            Dim htmlContent = ReadHtmlWithEncodingDetection(tempHtmlPath)

            ' Inject base href and MOTW
            Dim baseHref = $"<base href=""file:///{tempFolder.Replace("\", "/")}/"">"
            Dim motw = "<!-- saved from url=(0016)http://localhost -->"
            htmlContent = htmlContent.Replace("<head>", "<head>" & vbCrLf & motw & vbCrLf & baseHref)

            ' Capture for closures
            Dim capturedWordApp = wordApp
            Dim capturedTempFolder = tempFolder
            Dim capturedTempHtmlPath = tempHtmlPath
            Dim capturedSaveCallback = saveSnapshotCallback

            ' Build buttons
            Dim additionalButtons As New List(Of System.Tuple(Of String, System.Action, Boolean))()

            ' Summarize Changes
            If Not String.IsNullOrWhiteSpace(extractedChangesText) Then
                additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                    "Summarize Changes",
                    Sub() SummarizeComparisonChangesAsync(extractedChangesText),
                    False))
            End If

            ' Send to PDF
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Send to PDF",
                Sub()
                    If Not File.Exists(capturedTempHtmlPath) Then
                        ShowCustomMessageBox("The comparison file is no longer available.", AN)
                        Return
                    End If
                    ExportComparisonToPdfFromHtml(capturedTempHtmlPath, capturedWordApp, "Snapshot_Compare", Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
                End Sub,
                False))

            ' Copy to Clipboard
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Copy to Clipboard",
                Sub()
                    Try
                        If Not File.Exists(capturedTempHtmlPath) Then
                            ShowCustomMessageBox("The comparison file is no longer available.", AN)
                            Return
                        End If
                        Dim tempDoc As Microsoft.Office.Interop.Word.Document = Nothing
                        Dim prevScreenUpdating2 = capturedWordApp.ScreenUpdating
                        Dim prevAlerts2 = capturedWordApp.DisplayAlerts
                        Try
                            capturedWordApp.ScreenUpdating = False
                            capturedWordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone
                            tempDoc = capturedWordApp.Documents.Open(FileName:=capturedTempHtmlPath, ReadOnly:=True, Visible:=False)
                            tempDoc.Content.Copy()
                            tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                            tempDoc = Nothing
                            capturedWordApp.DisplayAlerts = prevAlerts2
                            capturedWordApp.ScreenUpdating = prevScreenUpdating2
                            ShowCustomMessageBox("Comparison copied to clipboard with formatting.", AN)
                        Finally
                            If tempDoc IsNot Nothing Then Try : tempDoc.Close(WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
                            capturedWordApp.DisplayAlerts = prevAlerts2
                            capturedWordApp.ScreenUpdating = prevScreenUpdating2
                        End Try
                    Catch ex As Exception
                        ShowCustomMessageBox($"Failed to copy to clipboard: {ex.Message}", AN)
                    End Try
                End Sub,
                False))

            ' Send to Document
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

            ' SAVE SNAPSHOT button
            additionalButtons.Add(New System.Tuple(Of String, System.Action, Boolean)(
                "Save Snapshot",
                Sub()
                    capturedSaveCallback?.Invoke()
                End Sub,
                False))

            ' Cleanup action
            Dim cleanupAction As System.Action =
                Sub()
                    Try
                        If Directory.Exists(capturedTempFolder) Then
                            Directory.Delete(capturedTempFolder, recursive:=True)
                        End If
                    Catch
                    End Try
                End Sub

            ' Show the HTML dialog - use modal (default) which runs on its own STA thread
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
    ''' Creates a new snapshot document entry using a multi-field input form.
    ''' </summary>
    Private Async Function CreateNewSnapshotDocument(localLibrary As SnapshotLibrary) As System.Threading.Tasks.Task
        If localLibrary Is Nothing OrElse String.IsNullOrEmpty(localLibrary.FilePath) Then
            ShowCustomMessageBox("Local library is not configured.", AN)
            Return
        End If

        ' Prepare form fields using InputParameter - array size should match number of fields (4 fields = index 0-3)
        Dim params(3) As Slib.InputParameter
        params(0) = New Slib.InputParameter("Friendly Name", "")
        params(1) = New Slib.InputParameter("Location (URL or path)", "")
        params(2) = New Slib.InputParameter("Remove Clutter", "Yes")
        params(2).Options = New List(Of String) From {"Yes", "No"}
        params(3) = New Slib.InputParameter("Archive Path (optional)", "")

        Dim ok As Boolean = ShowCustomVariableInputForm("Enter document details:", $"{AN} - Add Snapshot Document", params)
        If Not ok Then
            Return ' Cancelled
        End If

        Dim friendlyName = If(params(0).Value, "").ToString().Trim()
        Dim location = If(params(1).Value, "").ToString().Trim()
        Dim removeClutterStr = If(params(2).Value, "Yes").ToString().Trim()
        Dim archivePath = If(params(3).Value, "").ToString().Trim()

        If String.IsNullOrEmpty(friendlyName) Then
            ShowCustomMessageBox("Friendly name is required.", AN)
            Return
        End If

        If String.IsNullOrEmpty(location) Then
            ShowCustomMessageBox("Location (URL or path) is required.", AN)
            Return
        End If

        ' Generate shortname from friendly name
        Dim shortname = GenerateShortname(friendlyName)

        ' Determine if clutter removal is enabled
        Dim removeClutter = removeClutterStr.StartsWith("Y", StringComparison.OrdinalIgnoreCase) OrElse
                           removeClutterStr.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse
                           removeClutterStr = "1"

        ' Create document object
        Dim newDoc As New SnapshotDocument() With {
            .FriendlyName = friendlyName,
            .Shortname = shortname,
            .Location = location,
            .RemoveClutter = removeClutter,
            .SnapshotArchive = archivePath,
            .IsLocal = True,
            .SourceLibraryPath = localLibrary.FilePath
        }

        ' Ensure library file exists
        EnsureSnapshotLibraryExists(localLibrary.FilePath)

        ' Test retrieval and save initial snapshot
        Dim confirmed = ShowCustomYesNoBox(
            $"Document configured:{vbCrLf}{vbCrLf}" &
            $"Name: {friendlyName}{vbCrLf}" &
            $"Location: {location}{vbCrLf}" &
            $"Remove Clutter: {If(removeClutter, "Yes", "No")}{vbCrLf}{vbCrLf}" &
            "Do you want to download and save the initial snapshot now?",
            "Yes, download now",
            "No, just add to library")

        ' Add to library file
        AddDocumentToLibrary(newDoc, localLibrary)

        If confirmed = 1 Then
            ' Attempt to download initial snapshot
            Try
                Dim resolvedArchive = ResolveArchivePath(newDoc, localLibrary)
                If Not Directory.Exists(resolvedArchive) Then
                    Directory.CreateDirectory(resolvedArchive)
                End If

                ' Capture download timestamp
                Dim downloadTimestamp As DateTime = DateTime.Now

                ' Show splash for download
                Dim splash As New Slib.SplashScreen($"Downloading: {friendlyName}...")
                splash.Show()
                splash.Refresh()
                System.Windows.Forms.Application.DoEvents()

                Dim content = Await RetrieveDocumentContentAsync(location)
                If String.IsNullOrEmpty(content) Then
                    splash.Close()
                    ShowCustomMessageBox("Failed to retrieve document content. The document has been added to the library, but no snapshot was saved.", AN)
                    Return
                End If

                ' Close splash before LLM call
                splash.Close()

                ' Remove clutter if requested
                If removeClutter Then
                    Try
                        Dim userPrompt = "<TEXTTOPROCESS>" & vbCrLf & content & vbCrLf & "</TEXTTOPROCESS>"
                        ' Let LLM show its own splash (HideSplash:=False)
                        Dim cleanedContent = Await LLM(
                            SP_RemoveClutter,
                            userPrompt,
                            "",
                            "",
                            0,
                            False,
                            False)

                        If Not String.IsNullOrWhiteSpace(cleanedContent) AndAlso Not cleanedContent.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then
                            content = cleanedContent
                        End If
                    Catch
                        ' Keep raw content on LLM failure
                    End Try
                End If

                ' Save snapshot using download timestamp
                Dim snapshotFilename = SaveSnapshotWithTimestamp(newDoc, content, resolvedArchive, downloadTimestamp)
                If Not String.IsNullOrEmpty(snapshotFilename) Then
                    AddSnapshotToLibrary(newDoc, snapshotFilename, localLibrary)
                    ShowCustomMessageBox($"Document added and initial snapshot saved: {snapshotFilename}", AN)
                Else
                    ShowCustomMessageBox("Document added to library, but failed to save snapshot.", AN)
                End If

            Catch ex As Exception
                ShowCustomMessageBox($"Document added to library, but snapshot failed: {ex.Message}", AN)
            End Try
        Else
            ShowCustomMessageBox("Document added to library. You can take a snapshot later.", AN)
        End If
    End Function

#End Region

#Region "Snapshot Management"

    ''' <summary>
    ''' Views snapshots - allows viewing content of existing snapshots.
    ''' </summary>
    Private Sub ViewSnapshots(
        allDocuments As List(Of SnapshotDocument),
        localLibrary As SnapshotLibrary,
        centralLibrary As SnapshotLibrary)

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
            items.Add(New SelectionItem($"{doc.FriendlyName} ({doc.Snapshots.Count} snapshots)", idx))
            docMap(idx) = doc
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select document to view snapshots:", $"{AN} - View Snapshots")

        If result <= 0 OrElse Not docMap.ContainsKey(result) Then
            Return
        End If

        Dim selectedDoc = docMap(result)
        ViewDocumentSnapshots(selectedDoc, If(selectedDoc.IsLocal, localLibrary, centralLibrary))
    End Sub

    ''' <summary>
    ''' Views snapshots for a specific document.
    ''' </summary>
    Private Sub ViewDocumentSnapshots(doc As SnapshotDocument, library As SnapshotLibrary)
        Dim archivePath = ResolveArchivePath(doc, library)

        ' Build list of snapshots with details
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

        Dim result = SelectValue(items, 1, $"Snapshots for: {doc.FriendlyName}{vbCrLf}Select a snapshot to view:", $"{AN} - View Snapshots")

        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then
            Dim snapshotToView = snapshotMap(result)
            Dim fullPath = Path.Combine(archivePath, snapshotToView)

            If Not File.Exists(fullPath) Then
                ShowCustomMessageBox($"Snapshot file not found: {snapshotToView}", AN)
                Return
            End If

            ' Open snapshot in the built-in text editor (read-only)
            Dim snapshotDate = SnapshotDocument.ParseSnapshotDate(snapshotToView)
            Dim title = $"{AN} - {doc.FriendlyName}"
            If snapshotDate.HasValue Then
                title &= $" [{snapshotDate.Value:yyyy-MM-dd HH:mm:ss}]"
            End If

            ShowTextFileEditor(fullPath, title, True, _context)

            ' Allow viewing another snapshot
            ViewDocumentSnapshots(doc, library)
        End If
    End Sub

    ''' <summary>
    ''' Manages snapshots - allows deletion of old snapshots.
    ''' </summary>
    Private Sub ManageSnapshots(
        allDocuments As List(Of SnapshotDocument),
        localLibrary As SnapshotLibrary,
        centralLibrary As SnapshotLibrary)

        ' First, select which document to manage
        Dim items As New List(Of SelectionItem)()
        Dim docMap As New Dictionary(Of Integer, SnapshotDocument)()

        Dim docsWithSnapshots = allDocuments.Where(Function(d As SnapshotDocument) d.Snapshots.Count > 0).ToList()

        If docsWithSnapshots.Count = 0 Then
            ShowCustomMessageBox("No documents have snapshots to manage.", AN)
            Return
        End If

        Dim idx = 1
        For Each doc As SnapshotDocument In docsWithSnapshots
            items.Add(New SelectionItem($"{doc.FriendlyName} ({doc.Snapshots.Count} snapshots)", idx))
            docMap(idx) = doc
            idx += 1
        Next

        Dim result = SelectValue(items, 1, "Select document to manage snapshots:", $"{AN} - Manage Snapshots")

        If result <= 0 OrElse Not docMap.ContainsKey(result) Then
            Return
        End If

        Dim selectedDoc = docMap(result)
        ManageDocumentSnapshots(selectedDoc, If(selectedDoc.IsLocal, localLibrary, centralLibrary))
    End Sub

    ''' <summary>
    ''' Manages snapshots for a specific document.
    ''' </summary>
    Private Sub ManageDocumentSnapshots(doc As SnapshotDocument, library As SnapshotLibrary)
        Dim archivePath = ResolveArchivePath(doc, library)

        ' Build list of snapshots with details
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

        ' For simplicity, use single selection but with delete confirmation
        ' A full multi-select would require a custom form
        Dim result = SelectValue(items, 1, $"Snapshots for: {doc.FriendlyName}{vbCrLf}Select a snapshot to delete:", $"{AN} - Manage Snapshots")

        If result > 0 AndAlso snapshotMap.ContainsKey(result) Then
            Dim snapshotToDelete = snapshotMap(result)

            Dim confirm = ShowCustomYesNoBox(
                $"Are you sure you want to delete this snapshot?{vbCrLf}{vbCrLf}{snapshotToDelete}{vbCrLf}{vbCrLf}This cannot be undone.",
                "Yes, delete",
                "No, cancel")

            If confirm = 1 Then
                Try
                    ' Delete file
                    Dim fullPath = Path.Combine(archivePath, snapshotToDelete)
                    If File.Exists(fullPath) Then
                        File.Delete(fullPath)
                    End If

                    ' Remove from library
                    RemoveSnapshotFromLibrary(doc, snapshotToDelete, library)

                    ShowCustomMessageBox($"Snapshot deleted: {snapshotToDelete}", AN)

                    ' Refresh if more snapshots remain
                    doc.Snapshots.Remove(snapshotToDelete)
                    If doc.Snapshots.Count > 0 Then
                        ManageDocumentSnapshots(doc, library)
                    End If

                Catch ex As Exception
                    ShowCustomMessageBox($"Failed to delete snapshot: {ex.Message}", AN)
                End Try
            End If
        End If
    End Sub




#End Region

#Region "Library Parsing and Writing"

    ''' <summary>
    ''' Parses a snapshot library file.
    ''' </summary>
    Private Function ParseSnapshotLibrary(filePath As String, isLocal As Boolean) As SnapshotLibrary
        Dim library As New SnapshotLibrary() With {.FilePath = filePath}

        If Not File.Exists(filePath) Then
            Return library
        End If

        Try
            Dim lines = File.ReadAllLines(filePath, Encoding.UTF8)
            Dim currentDoc As SnapshotDocument = Nothing
            Dim libraryDir = Path.GetDirectoryName(filePath)

            For Each line As String In lines
                Dim trimmedLine = line.Trim()

                ' Skip empty lines and comments
                If String.IsNullOrEmpty(trimmedLine) OrElse trimmedLine.StartsWith(";") Then
                    Continue For
                End If

                ' Check for section header [Document Name]
                If trimmedLine.StartsWith("[") AndAlso trimmedLine.EndsWith("]") Then
                    ' Save previous document if exists
                    If currentDoc IsNot Nothing Then
                        library.Documents.Add(currentDoc)
                    End If

                    ' Start new document
                    Dim docName = trimmedLine.Substring(1, trimmedLine.Length - 2).Trim()
                    currentDoc = New SnapshotDocument() With {
                        .FriendlyName = docName,
                        .IsLocal = isLocal,
                        .SourceLibraryPath = filePath
                    }
                    Continue For
                End If

                ' Parse key=value
                Dim eqIndex = trimmedLine.IndexOf("="c)
                If eqIndex > 0 Then
                    Dim key = trimmedLine.Substring(0, eqIndex).Trim().ToLowerInvariant()
                    Dim value = trimmedLine.Substring(eqIndex + 1).Trim()

                    If currentDoc Is Nothing Then
                        ' Global settings
                        If key = "snapshotarchive" Then
                            library.GlobalSnapshotArchive = value
                        End If
                    Else
                        ' Document settings
                        Select Case key
                            Case "shortname"
                                currentDoc.Shortname = value
                            Case "location"
                                currentDoc.Location = value
                            Case "removeclutter"
                                currentDoc.RemoveClutter = value.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse
                                                          value.Equals("Yes", StringComparison.OrdinalIgnoreCase) OrElse
                                                          value = "1"
                            Case "snapshotarchive"
                                currentDoc.SnapshotArchive = value
                            Case "snapshot"
                                currentDoc.Snapshots.Add(value)
                        End Select
                    End If
                End If
            Next

            ' Don't forget the last document
            If currentDoc IsNot Nothing Then
                library.Documents.Add(currentDoc)
            End If

            ' Ensure all documents have shortnames
            For Each doc As SnapshotDocument In library.Documents
                If String.IsNullOrEmpty(doc.Shortname) Then
                    doc.Shortname = GenerateShortname(doc.FriendlyName)
                End If
            Next

        Catch ex As Exception
            ' Return empty library on parse error
        End Try

        Return library
    End Function

    ''' <summary>
    ''' Ensures the snapshot library file exists with basic structure.
    ''' </summary>
    Private Sub EnsureSnapshotLibraryExists(filePath As String)
        If File.Exists(filePath) Then Return

        Try
            Dim dir = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If

            Dim content = New StringBuilder()
            content.AppendLine("; Snapshot Library")
            content.AppendLine($"; Created: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            content.AppendLine()
            content.AppendLine("; Global settings")
            content.AppendLine("SnapshotArchive = SnapshotArchive")
            content.AppendLine()
            content.AppendLine("; Add documents below using this format:")
            content.AppendLine("; [Friendly Document Name]")
            content.AppendLine("; Location = https://example.com/page or C:\path\to\file.pdf")
            content.AppendLine("; RemoveClutter = True")
            content.AppendLine("; Snapshot = YYMMDD_HHMMSS_shortname.txt")
            content.AppendLine()

            File.WriteAllText(filePath, content.ToString(), Encoding.UTF8)

        Catch ex As Exception
            ShowCustomMessageBox($"Failed to create library file: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Adds a new document to the library file.
    ''' </summary>
    Private Sub AddDocumentToLibrary(doc As SnapshotDocument, library As SnapshotLibrary)
        Try
            EnsureSnapshotLibraryExists(library.FilePath)

            Dim sb As New StringBuilder()
            sb.AppendLine()
            sb.AppendLine($"[{doc.FriendlyName}]")
            sb.AppendLine($"Shortname = {doc.Shortname}")
            sb.AppendLine($"Location = {doc.Location}")
            sb.AppendLine($"RemoveClutter = {If(doc.RemoveClutter, "True", "False")}")
            If Not String.IsNullOrEmpty(doc.SnapshotArchive) Then
                sb.AppendLine($"SnapshotArchive = {doc.SnapshotArchive}")
            End If

            File.AppendAllText(library.FilePath, sb.ToString(), Encoding.UTF8)

        Catch ex As Exception
            ShowCustomMessageBox($"Failed to add document to library: {ex.Message}", AN)
        End Try
    End Sub

    ''' <summary>
    ''' Adds a snapshot entry to an existing document in the library.
    ''' </summary>
    Private Sub AddSnapshotToLibrary(doc As SnapshotDocument, snapshotFilename As String, library As SnapshotLibrary)
        Try
            If Not File.Exists(library.FilePath) Then Return

            Dim lines = File.ReadAllLines(library.FilePath, Encoding.UTF8).ToList()
            Dim inTargetSection = False
            Dim insertIndex = -1

            For i = 0 To lines.Count - 1
                Dim line = lines(i).Trim()

                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    If inTargetSection Then
                        ' We've reached the next section, insert before it
                        insertIndex = i
                        Exit For
                    End If

                    Dim sectionName = line.Substring(1, line.Length - 2).Trim()
                    If sectionName.Equals(doc.FriendlyName, StringComparison.OrdinalIgnoreCase) Then
                        inTargetSection = True
                    End If
                End If
            Next

            If inTargetSection AndAlso insertIndex = -1 Then
                ' Document section is the last one, append at end
                insertIndex = lines.Count
            End If

            If insertIndex >= 0 Then
                lines.Insert(insertIndex, $"Snapshot = {snapshotFilename}")
                File.WriteAllLines(library.FilePath, lines, Encoding.UTF8)
            End If

        Catch ex As Exception
            ' Silently fail - snapshot is saved, just not recorded
        End Try
    End Sub

    ''' <summary>
    ''' Removes a snapshot entry from the library file.
    ''' </summary>
    Private Sub RemoveSnapshotFromLibrary(doc As SnapshotDocument, snapshotFilename As String, library As SnapshotLibrary)
        Try
            If library Is Nothing OrElse String.IsNullOrEmpty(library.FilePath) Then Return
            If Not File.Exists(library.FilePath) Then Return

            Dim lines = File.ReadAllLines(library.FilePath, Encoding.UTF8).ToList()
            Dim lineToRemove = $"Snapshot = {snapshotFilename}"

            ' Find the document section and remove the snapshot line
            Dim inTargetSection = False
            Dim removed = False

            For i = 0 To lines.Count - 1
                Dim line = lines(i).Trim()

                ' Check for section headers
                If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                    Dim sectionName = line.Substring(1, line.Length - 2).Trim()
                    inTargetSection = sectionName.Equals(doc.FriendlyName, StringComparison.OrdinalIgnoreCase)
                ElseIf inTargetSection AndAlso line.Equals(lineToRemove, StringComparison.OrdinalIgnoreCase) Then
                    lines.RemoveAt(i)
                    removed = True
                    Exit For
                End If
            Next

            If removed Then
                File.WriteAllLines(library.FilePath, lines, Encoding.UTF8)
            End If

        Catch ex As Exception
            ' Log but don't show error to user - snapshot file is already deleted
            Debug.WriteLine($"Failed to remove snapshot from library: {ex.Message}")
        End Try
    End Sub


#End Region

#Region "Helper Methods"

    ''' <summary>
    ''' Retrieves document content from URL or local path (async version).
    ''' </summary>
    Private Async Function RetrieveDocumentContentAsync(location As String) As Task(Of String)
        If String.IsNullOrEmpty(location) Then Return Nothing

        Try
            ' Check if it's a URL
            If location.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
               location.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                ' Use WebView2 for web content
                Return Await RetrieveWebsiteContent_WebView2(location)
            Else
                ' Local file or UNC path
                If File.Exists(location) Then
                    ' Use GetFileContent for proper handling of various formats
                    Return Await GetFileContent(location, True, SharedMethods.IsOcrAvailable(_context), False)
                Else
                    Return Nothing
                End If
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Retrieves document content from URL or local path (sync version for backward compatibility).
    ''' </summary>
    Private Function RetrieveDocumentContent(location As String) As String
        Return RetrieveDocumentContentAsync(location).GetAwaiter().GetResult()
    End Function

    ''' <summary>
    ''' Resolves the archive path for a document.
    ''' </summary>
    Private Function ResolveArchivePath(doc As SnapshotDocument, library As SnapshotLibrary) As String
        Dim libraryDir = Path.GetDirectoryName(library.FilePath)

        ' Priority: Document-specific > Library global > Default
        Dim archivePath = doc.SnapshotArchive
        If String.IsNullOrEmpty(archivePath) Then
            archivePath = library.GlobalSnapshotArchive
        End If
        If String.IsNullOrEmpty(archivePath) Then
            archivePath = "SnapshotArchive"
        End If

        archivePath = ExpandEnvironmentVariables(archivePath)

        ' Resolve relative paths
        If Not Path.IsPathRooted(archivePath) Then
            archivePath = Path.Combine(libraryDir, archivePath)
        End If

        Return archivePath
    End Function

    ''' <summary>
    ''' Saves content as a snapshot file using the provided timestamp (download time).
    ''' </summary>
    Private Function SaveSnapshotWithTimestamp(doc As SnapshotDocument, content As String, archivePath As String, timestamp As DateTime) As String
        Try
            Dim timestampStr = timestamp.ToString("yyMMdd_HHmmss")
            Dim filename = $"{timestampStr}_{doc.Shortname}.txt"
            Dim fullPath = Path.Combine(archivePath, filename)

            ' Check if file already exists (duplicate save attempt)
            If File.Exists(fullPath) Then
                ShowCustomMessageBox($"A snapshot with this timestamp already exists: {filename}", AN)
                Return Nothing
            End If

            ' Normalize line endings to CRLF for Windows compatibility
            Dim normalizedContent = content
            If Not String.IsNullOrEmpty(normalizedContent) Then
                ' First convert any CRLF to LF, then convert all LF to CRLF
                normalizedContent = normalizedContent.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, vbCrLf)
            End If

            File.WriteAllText(fullPath, normalizedContent, Encoding.UTF8)

            Return filename
        Catch ex As Exception
            ShowCustomMessageBox($"Failed to save snapshot: {ex.Message}", AN)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Saves content as a snapshot file (uses current time - for backward compatibility).
    ''' </summary>
    Private Function SaveSnapshot(doc As SnapshotDocument, content As String, archivePath As String) As String
        Return SaveSnapshotWithTimestamp(doc, content, archivePath, DateTime.Now)
    End Function

    ''' <summary>
    ''' Generates a valid filename shortname from a friendly name.
    ''' </summary>
    Private Function GenerateShortname(friendlyName As String) As String
        If String.IsNullOrEmpty(friendlyName) Then Return "document"

        ' Remove invalid filename characters
        Dim invalid = Path.GetInvalidFileNameChars()
        Dim result = friendlyName

        For Each c In invalid
            result = result.Replace(c, "_"c)
        Next

        ' Replace spaces with underscores
        result = result.Replace(" "c, "_"c)

        ' Remove consecutive underscores
        While result.Contains("__")
            result = result.Replace("__", "_")
        End While

        ' Trim underscores from ends
        result = result.Trim("_"c)

        ' Limit length
        If result.Length > 30 Then
            result = result.Substring(0, 30)
        End If

        ' Ensure not empty
        If String.IsNullOrEmpty(result) Then
            result = "document"
        End If

        Return result.ToLowerInvariant()
    End Function

    ''' <summary>
    ''' Computes SHA256 hash of content for comparison.
    ''' </summary>
    Private Function ComputeContentHash(content As String) As String
        If String.IsNullOrEmpty(content) Then Return ""

        Using hasher As SHA256 = SHA256.Create()
            Dim bytes = Encoding.UTF8.GetBytes(content)
            Dim hash = hasher.ComputeHash(bytes)
            Return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant()
        End Using
    End Function

    ''' <summary>
    ''' Formats file size for display.
    ''' </summary>
    Private Function FormatFileSize(bytes As Long) As String
        If bytes < 1024 Then Return $"{bytes} B"
        If bytes < 1024 * 1024 Then Return $"{bytes / 1024.0:F1} KB"
        Return $"{bytes / (1024.0 * 1024.0):F1} MB"
    End Function

#End Region

End Class