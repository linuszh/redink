' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.Agent.vb
' Purpose: Chat Agent mode — provides the same document-processing tooling
'          available in AutoPilot to the browser-based local chat. Users can
'          load multiple files, and output files are collected into a Desktop
'          folder after the LLM tooling loop completes.
'
' Architecture:
'   - Shares the AutoPilot internal tool infrastructure via the same shared
'     fields (_apCurrentTempDir, _apCurrentAttachments, _apCurrentMailInfo).
'   - Blocked when AutoPilot is actively running to avoid state collision.
'   - Creates a per-job output folder on Desktop\Inky\yymmdd_hh-mm and copies
'     result files there after the tooling loop finishes.
'   - The summarize_thread tool is excluded (e-mail-specific).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Threading
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CHAT AGENT STATE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>True while a chat agent tooling job is executing. Enables internal tool routing.</summary>
    Friend _chatAgentActive As Boolean = False

    ''' <summary>Files loaded by the user for the current chat agent session.</summary>
    Private _chatAgentFiles As New List(Of AutoPilotAttachmentInfo)()

    ''' <summary>Temp directory used by the current chat agent session.</summary>
    Private _chatAgentTempDir As String = Nothing

    ''' <summary>Prefix for chat agent temp directories.</summary>
    Private Const CA_TempPrefix As String = AN2 & "_chatagent_"

    ''' <summary>Maximum file size for chat agent uploads (same as AutoPilot default).</summary>
    Private Const CA_DefaultMaxAttachmentBytes As Long = 10 * 1024 * 1024

    ' ═══════════════════════════════════════════════════════════════════════════
    '  FILE MANAGEMENT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Adds a file from disk to the chat agent file list. The file is copied to a
    ''' per-session temp directory so tools can operate on it safely.
    ''' The source file is deleted after a successful copy to avoid orphaned uploads.
    ''' </summary>
    Private Function ChatAgentAddFile(sourcePath As String) As AutoPilotAttachmentInfo
        If String.IsNullOrWhiteSpace(sourcePath) OrElse Not File.Exists(sourcePath) Then Return Nothing

        ' Ensure per-session temp dir exists
        If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
            _chatAgentTempDir = Path.Combine(Path.GetTempPath(), CA_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(_chatAgentTempDir)
        End If

        ' The upload handler (inky_upload) saves files as "{32-hex-guid}_{originalName}".
        ' Strip that GUID prefix so tools see the user's original filename.
        Dim fileName = Path.GetFileName(sourcePath)
        If fileName.Length > 33 AndAlso fileName(32) = "_"c Then
            Dim possibleGuid = fileName.Substring(0, 32)
            Dim g As Guid
            If Guid.TryParse(possibleGuid, g) Then
                fileName = fileName.Substring(33)
            End If
        End If

        Dim destPath = Path.Combine(_chatAgentTempDir, fileName)

        ' Prevent collision
        Dim counter = 1
        While File.Exists(destPath)
            Dim baseName = Path.GetFileNameWithoutExtension(fileName)
            Dim ext = Path.GetExtension(fileName)
            destPath = Path.Combine(_chatAgentTempDir, baseName & $"_{counter}{ext}")
            counter += 1
        End While

        File.Copy(sourcePath, destPath, overwrite:=True)

        ' Delete the source upload file — the copy in the agent temp dir is the working copy
        Try
            File.Delete(sourcePath)
        Catch
        End Try

        Dim fi = New FileInfo(destPath)
        Dim att As New AutoPilotAttachmentInfo() With {
            .OriginalFileName = Path.GetFileName(destPath),
            .TempFilePath = destPath,
            .Extension = fi.Extension.ToLowerInvariant(),
            .SizeBytes = fi.Length,
            .IsOverSizeLimit = (fi.Length > CA_DefaultMaxAttachmentBytes),
            .StatusMessage = If(fi.Length > CA_DefaultMaxAttachmentBytes, "Over size limit", "OK"),
            .CreatedTime = fi.CreationTimeUtc,
            .LastModifiedTime = fi.LastWriteTimeUtc,
            .OutputFiles = New List(Of String)(),
            .IsToolOutput = False
        }

        _chatAgentFiles.Add(att)
        Return att
    End Function

    ''' <summary>
    ''' Returns a browser-friendly list of loaded agent files.
    ''' </summary>
    Private Function GetAgentFileListForBrowser() As List(Of Object)
        Dim result As New List(Of Object)()
        For Each att In _chatAgentFiles
            result.Add(New With {
                .name = att.OriginalFileName,
                .size = att.SizeBytes,
                .ext = att.Extension,
                .overLimit = att.IsOverSizeLimit
            })
        Next
        Return result
    End Function

    ''' <summary>
    ''' Clears all chat agent files and removes the temp directory.
    ''' </summary>
    Private Sub ChatAgentClearFiles()
        CleanupChatAgentTempDir()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  AGENT CONTEXT SETUP / TEARDOWN
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Sets up the shared AutoPilot fields so that internal tools can operate
    ''' on the chat agent's files. Must be called before ExecuteToolingLoop.
    ''' Returns the combined tool list (internal agent tools + selected external tools).
    ''' </summary>
    Private Function ChatAgentSetupToolContext() As List(Of ModelConfig)
        ' Ensure per-session temp dir exists (tools may create files even without user uploads)
        If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
            _chatAgentTempDir = Path.Combine(Path.GetTempPath(), CA_TempPrefix & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(_chatAgentTempDir)
        End If

        ' Set shared fields that tools read
        _apCurrentTempDir = _chatAgentTempDir
        _apCurrentAttachments = _chatAgentFiles
        _apCurrentToolCallLog = New List(Of AutoPilotToolCallEntry)()

        ' Synthesize a minimal mail info (some tools guard on it)
        _apCurrentMailInfo = New AutoPilotMailInfo() With {
            .EntryID = "",
            .Subject = "Chat Agent Session",
            .SenderName = Environment.UserName,
            .SenderEmail = "",
            .Body = "",
            .ReceivedTime = DateTime.Now,
            .HasAutoReplyHeader = False,
            .ThreadAIReplyCount = 0,
            .AttachmentCount = _chatAgentFiles.Count,
            .AttachmentNames = _chatAgentFiles.Select(Function(a) a.OriginalFileName).ToList(),
            .FolderPath = "",
            .MessageClass = "",
            .InternetHeaders = ""
        }

        _chatAgentActive = True

        ' Build tool list: internal tools (excluding summarize_thread) + selected external chat tools
        Dim tools As New List(Of ModelConfig)()
        Dim internalTools = GetAutoPilotInternalTools()

        ' Exclude e-mail-specific tools
        For Each t In internalTools
            If t.ToolName = AP_Tool_SummarizeThread Then Continue For
            tools.Add(t)
        Next

        ' Add selected external/web tools from the chat session
        If _selectedToolsForChat IsNot Nothing Then
            tools.AddRange(_selectedToolsForChat)
        End If

        Return tools
    End Function

    ''' <summary>
    ''' Clears the shared AutoPilot fields after an agent job completes.
    ''' Deletes the temp directory so no temp files remain on disk.
    ''' Clears attachment caches for security.
    ''' </summary>
    Private Sub ChatAgentTeardownToolContext()
        _chatAgentActive = False
        ClearAttachmentCaches()
        _apCurrentTempDir = Nothing
        _apCurrentAttachments = Nothing
        _apCurrentMailInfo = Nothing
        _apCurrentToolCallLog = Nothing
        ' Ensure temp dir is cleaned up (covers cancellation / exception paths
        ' where ChatAgentCollectAndCopyOutputs was never reached)
        CleanupChatAgentTempDir()
    End Sub

    ' ═══════════════════════════════════════════════════════════════════════════
    '  OUTPUT FOLDER
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Collects output files from the agent session, copies them to
    ''' Desktop\Inky\yymmdd_hh-mm\, and opens the folder in Explorer.
    ''' Returns the list of copied file paths (for the assistant message).
    ''' After collection, deletes the entire temp directory so no temp files
    ''' remain on disk — only the Desktop output folder survives.
    ''' </summary>
    Private Function ChatAgentCollectAndCopyOutputs() As List(Of String)
        Dim copiedFiles As New List(Of String)()

        If String.IsNullOrWhiteSpace(_chatAgentTempDir) OrElse Not Directory.Exists(_chatAgentTempDir) Then
            Return copiedFiles
        End If

        ' Collect result files using the same logic as AutoPilot
        Dim resultFiles = CollectResultAttachments(_chatAgentTempDir, _chatAgentFiles)
        If resultFiles Is Nothing OrElse resultFiles.Count = 0 Then
            ' Even with no outputs, clean up the temp dir
            CleanupChatAgentTempDir()
            Return copiedFiles
        End If

        ' Create output folder: Desktop\Inky\yymmdd_hh-mm
        Dim desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        Dim timestamp = DateTime.Now.ToString("yyMMdd_HH-mm")
        Dim outputDir = Path.Combine(desktopPath, "Inky", timestamp)

        ' Handle collision (rapid successive runs)
        Dim counter = 1
        While Directory.Exists(outputDir)
            outputDir = Path.Combine(desktopPath, "Inky", timestamp & $"_{counter}")
            counter += 1
        End While

        Directory.CreateDirectory(outputDir)

        For Each srcPath In resultFiles
            Try
                Dim destName = Path.GetFileName(srcPath)
                Dim destPath = Path.Combine(outputDir, destName)
                ' Handle name collision in output folder
                Dim fileCounter = 1
                While File.Exists(destPath)
                    Dim baseName = Path.GetFileNameWithoutExtension(destName)
                    Dim ext = Path.GetExtension(destName)
                    destPath = Path.Combine(outputDir, baseName & $"_{fileCounter}{ext}")
                    fileCounter += 1
                End While
                File.Copy(srcPath, destPath, overwrite:=False)
                copiedFiles.Add(destPath)
            Catch
            End Try
        Next

        ' Open the folder in Explorer
        If copiedFiles.Count > 0 Then
            Try
                Process.Start("explorer.exe", outputDir)
            Catch
            End Try
        End If

        ' Clean up: delete the entire temp directory and reset tracking.
        ' The only surviving files are the copies in the Desktop output folder.
        CleanupChatAgentTempDir()

        Return copiedFiles
    End Function

    ''' <summary>
    ''' Deletes the chat agent temp directory (recursively, including subdirectories)
    ''' and resets the file tracking list. Safe to call multiple times.
    ''' </summary>
    Private Sub CleanupChatAgentTempDir()
        _chatAgentFiles.Clear()
        Try
            If Not String.IsNullOrWhiteSpace(_chatAgentTempDir) AndAlso Directory.Exists(_chatAgentTempDir) Then
                Directory.Delete(_chatAgentTempDir, recursive:=True)
            End If
        Catch
        End Try
        _chatAgentTempDir = Nothing
    End Sub

    ''' <summary>
    ''' Returns True if the chat agent mode is blocked because AutoPilot is running.
    ''' </summary>
    Private Function IsChatAgentBlocked() As Boolean
        Return _apActive
    End Function

End Class