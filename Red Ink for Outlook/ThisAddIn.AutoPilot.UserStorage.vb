' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.UserStorage.vb
' Purpose:
'   Manages per-user persistent storage for AutoPilot — both InkyMemory
'   (natural-language preferences) and Home Directory (persistent user files).
'
' Architecture:
'  - Storage layout (alongside redink.ini / scheduler files):
'      {INI dir}/autopilot_users/{sanitized_email}/
'          memory.txt           — user's InkyMemory file
'          home/                — user's persistent file storage
'              template.docx
'              ...
'
'  - E-mail addresses are sanitised into filesystem-safe folder names
'    using SanitizeEmailToFolderName(). Sanitisation is deterministic and
'    case-insensitive (lowercased first).
'
'  - Security model:
'      * Every path operation validates that resolved paths stay within
'        the user's own subdirectory (path-prefix containment check).
'      * No API accepts a raw path from the user — only filenames that
'        are resolved against the user's home directory.
'      * Cross-user access is impossible because the sender e-mail from
'        the incoming mail determines the directory, not user input.
'
'  - Operator (admin) functions are provided for the dashboard:
'      * ListAllUserStorageDirs() — enumerate all user directories
'      * ListUserHomeFiles() — list files in a user's home directory
'      * DeleteUserFile() / DeleteUserAllFiles() / DeleteUserMemory()
'      * ReadUserMemoryContent() / WriteUserMemoryContent()
'
' Threading:
'  - File I/O is synchronous with retry-based locking (same as InkyMemory).
'  - No concurrent mutation is expected per-user (one mail processed at a time).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ═══════════════════════════════════════════════════════════════════════════
    '  CONSTANTS
    ' ═══════════════════════════════════════════════════════════════════════════

    Private Const AP_UserStorageDir As String = "autopilot_users"
    Private Const AP_UserMemoryFileName As String = "memory.txt"
    Private Const AP_UserHomeSubdir As String = "home"
    Private Const AP_UserHomeMaxBytes As Long = 100 * 1024 * 1024  ' 100 MB per user

    ' ═══════════════════════════════════════════════════════════════════════════
    '  PATH HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Returns the root directory for all per-user AutoPilot storage:
    ''' {INI dir}/autopilot_users/
    ''' </summary>
    Private Function GetUserStorageRootDir() As String
        Dim iniPath = GetDefaultINIPath(_context.RDV)
        Dim iniDir = Path.GetDirectoryName(iniPath)
        Return Path.Combine(iniDir, AP_UserStorageDir)
    End Function

    ''' <summary>
    ''' Returns the per-user directory: {root}/autopilot_users/{sanitized_email}/
    ''' Creates the directory if it does not exist.
    ''' </summary>
    Private Function GetUserDir(senderEmail As String) As String
        Dim sanitized = SanitizeEmailToFolderName(senderEmail)
        Dim userDir = Path.Combine(GetUserStorageRootDir(), sanitized)
        If Not Directory.Exists(userDir) Then Directory.CreateDirectory(userDir)
        Return userDir
    End Function

    ''' <summary>
    ''' Returns the full path to a user's InkyMemory file:
    ''' {root}/autopilot_users/{sanitized_email}/memory.txt
    ''' </summary>
    Private Function GetUserMemoryFilePath(senderEmail As String) As String
        Return Path.Combine(GetUserDir(senderEmail), AP_UserMemoryFileName)
    End Function

    ''' <summary>
    ''' Returns the full path to a user's home directory:
    ''' {root}/autopilot_users/{sanitized_email}/home/
    ''' Creates the directory if it does not exist.
    ''' </summary>
    Private Function GetUserHomeDir(senderEmail As String) As String
        Dim homeDir = Path.Combine(GetUserDir(senderEmail), AP_UserHomeSubdir)
        If Not Directory.Exists(homeDir) Then Directory.CreateDirectory(homeDir)
        Return homeDir
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  E-MAIL SANITISATION
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Converts an e-mail address into a filesystem-safe, deterministic folder name.
    ''' Case-insensitive (lowercased). Only alphanumeric, hyphen, and underscore survive;
    ''' '@' becomes '_at_', '.' becomes '_', everything else becomes '_'.
    ''' </summary>
    ''' <example>"John.Doe@Example.COM" → "john_doe_at_example_com"</example>
    Private Shared Function SanitizeEmailToFolderName(email As String) As String
        If String.IsNullOrWhiteSpace(email) Then Return "_unknown_"
        Dim lower = email.Trim().ToLowerInvariant()
        Dim sb As New StringBuilder(lower.Length)
        For Each c In lower
            If Char.IsLetterOrDigit(c) OrElse c = "-"c Then
                sb.Append(c)
            ElseIf c = "@"c Then
                sb.Append("_at_")
            Else
                sb.Append("_"c)
            End If
        Next
        Dim result = sb.ToString().Trim("_"c)
        If result.Length = 0 Then result = "_unknown_"
        Return result
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  SECURITY: PATH CONTAINMENT
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Validates that <paramref name="candidatePath"/> resolves to a location
    ''' strictly inside <paramref name="requiredParent"/>. Prevents path traversal attacks.
    ''' </summary>
    ''' <returns>True if the candidate is safely contained; False otherwise.</returns>
    Private Shared Function IsPathContained(candidatePath As String, requiredParent As String) As Boolean
        Dim fullCandidate = Path.GetFullPath(candidatePath)
        Dim fullParent = Path.GetFullPath(requiredParent)
        If Not fullParent.EndsWith(Path.DirectorySeparatorChar.ToString()) Then
            fullParent &= Path.DirectorySeparatorChar
        End If
        Return fullCandidate.StartsWith(fullParent, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Resolves a user-provided filename (no path separators allowed) against a user's
    ''' home directory and validates path containment. Returns the safe full path, or
    ''' Nothing if the filename is unsafe.
    ''' </summary>
    Private Function ResolveUserHomeFileSafe(senderEmail As String, fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then Return Nothing
        ' Strip any path component — only the bare filename is allowed
        Dim bareName = Path.GetFileName(fileName)
        If String.IsNullOrWhiteSpace(bareName) Then Return Nothing
        If bareName <> fileName.Trim() Then Return Nothing ' reject if path separators were present

        Dim homeDir = GetUserHomeDir(senderEmail)
        Dim fullPath = Path.Combine(homeDir, bareName)
        If Not IsPathContained(fullPath, homeDir) Then Return Nothing
        Return fullPath
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  USER MEMORY — READ / WRITE / ENABLE / DISABLE
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Returns True if the user has an active memory file.</summary>
    Private Function IsUserMemoryEnabled(senderEmail As String) As Boolean
        Return File.Exists(GetUserMemoryFilePath(senderEmail))
    End Function

    ''' <summary>
    ''' Reads a user's memory file and returns its content as a string suitable for
    ''' inclusion in a system prompt. Returns empty string if not found or empty.
    ''' </summary>
    Private Function ReadUserMemory(senderEmail As String, maxItems As Integer) As String
        Return ReadInkyMemoryFromFile(GetUserMemoryFilePath(senderEmail), maxItems)
    End Function

    ''' <summary>Creates a user's memory file with the default header (opt-in).</summary>
    Private Sub EnableUserMemory(senderEmail As String)
        Dim filePath = GetUserMemoryFilePath(senderEmail)
        If Not File.Exists(filePath) Then
            Dim folder = Path.GetDirectoryName(filePath)
            If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
            WriteFileWithRetry(filePath, GetDefaultMemoryFileContent())
        End If
    End Sub

    ''' <summary>Deletes a user's memory file (opt-out). Does NOT delete home files.</summary>
    Private Sub DisableUserMemory(senderEmail As String)
        Dim filePath = GetUserMemoryFilePath(senderEmail)
        Try
            If File.Exists(filePath) Then File.Delete(filePath)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Processes an LLM response for automatic memory learning — parses and applies
    ''' &lt;INKY_MEMORY&gt; blocks to the user's per-user memory file.
    ''' Returns the cleaned response.
    ''' </summary>
    Private Function ProcessUserMemoryResponse(senderEmail As String, llmResponse As String, maxItems As Integer) As String
        Return ProcessInkyMemoryResponseForFile(GetUserMemoryFilePath(senderEmail), llmResponse, maxItems)
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  USER HOME DIRECTORY — FILE OPERATIONS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>Returns True if the user has any files in their home directory.</summary>
    Private Function HasUserHomeFiles(senderEmail As String) As Boolean
        Dim homeDir = GetUserHomeDir(senderEmail)
        Return Directory.Exists(homeDir) AndAlso Directory.GetFiles(homeDir).Length > 0
    End Function

    ''' <summary>Lists all files in the user's home directory with name and size.</summary>
    Private Function ListUserHomeFiles(senderEmail As String) As List(Of (Name As String, SizeBytes As Long))
        Dim result As New List(Of (Name As String, SizeBytes As Long))()
        Dim homeDir = GetUserHomeDir(senderEmail)
        If Not Directory.Exists(homeDir) Then Return result
        For Each filePath In Directory.GetFiles(homeDir)
            Dim fi As New FileInfo(filePath)
            result.Add((fi.Name, fi.Length))
        Next
        Return result
    End Function

    ''' <summary>
    ''' Returns the total size in bytes of all files in the user's home directory.
    ''' </summary>
    Private Function GetUserHomeSize(senderEmail As String) As Long
        Dim homeDir = GetUserHomeDir(senderEmail)
        If Not Directory.Exists(homeDir) Then Return 0
        Dim total As Long = 0
        For Each filePath In Directory.GetFiles(homeDir)
            total += New FileInfo(filePath).Length
        Next
        Return total
    End Function

    ''' <summary>
    ''' Copies a file from the current mail's temp directory into the user's home directory.
    ''' Validates path containment and enforces the per-user size cap.
    ''' </summary>
    ''' <returns>Success message or error description.</returns>
    Private Function StoreFileToUserHome(senderEmail As String, sourceFilePath As String, targetFileName As String) As String
        Try
            If Not File.Exists(sourceFilePath) Then Return "Error: Source file not found."

            Dim safePath = ResolveUserHomeFileSafe(senderEmail, targetFileName)
            If safePath Is Nothing Then Return "Error: Invalid filename (path separators or traversal not allowed)."

            ' Size check
            Dim sourceSize = New FileInfo(sourceFilePath).Length
            Dim currentTotal = GetUserHomeSize(senderEmail)
            If currentTotal + sourceSize > AP_UserHomeMaxBytes Then
                Return $"Error: Storage limit exceeded. Current usage: {currentTotal / 1024 / 1024:F1} MB, " &
                       $"file size: {sourceSize / 1024 / 1024:F1} MB, limit: {AP_UserHomeMaxBytes / 1024 / 1024:F0} MB."
            End If

            File.Copy(sourceFilePath, safePath, overwrite:=True)
            Return $"File '{targetFileName}' stored successfully. " &
                   "Note: This file is now available in all future AutoPilot sessions and can be referenced " &
                   "by tools such as process_word_document. Use manage_user_files to list, remove, or retrieve it."
        Catch ex As Exception
            Return $"Error storing file: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' Removes a file from the user's home directory.
    ''' </summary>
    Private Function RemoveFileFromUserHome(senderEmail As String, fileName As String) As String
        Try
            Dim safePath = ResolveUserHomeFileSafe(senderEmail, fileName)
            If safePath Is Nothing Then Return "Error: Invalid filename."
            If Not File.Exists(safePath) Then Return $"Error: File '{fileName}' not found."
            File.Delete(safePath)
            Return $"File '{fileName}' removed successfully."
        Catch ex As Exception
            Return $"Error removing file: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' Copies a file from the user's home directory into the current mail's temp directory
    ''' and registers it as an available attachment for the current processing session.
    ''' Used for "checkout" (return to user) and "use" (load into session for tool access).
    ''' </summary>
    ''' <returns>The AutoPilotAttachmentInfo for the loaded file, or Nothing on failure.</returns>
    Private Function LoadFileFromUserHome(senderEmail As String, fileName As String) As AutoPilotAttachmentInfo
        Try
            Dim safePath = ResolveUserHomeFileSafe(senderEmail, fileName)
            If safePath Is Nothing OrElse Not File.Exists(safePath) Then Return Nothing
            If _apCurrentTempDir Is Nothing OrElse Not Directory.Exists(_apCurrentTempDir) Then Return Nothing

            Dim destPath = Path.Combine(_apCurrentTempDir, Path.GetFileName(safePath))
            Dim counter = 1
            While File.Exists(destPath)
                destPath = Path.Combine(_apCurrentTempDir,
                    Path.GetFileNameWithoutExtension(fileName) & $"_{counter}" & Path.GetExtension(fileName))
                counter += 1
            End While

            File.Copy(safePath, destPath)

            Dim info As New AutoPilotAttachmentInfo() With {
                .OriginalFileName = Path.GetFileName(destPath),
                .TempFilePath = destPath,
                .Extension = Path.GetExtension(destPath).ToLowerInvariant(),
                .SizeBytes = New FileInfo(destPath).Length,
                .IsOverSizeLimit = False,
                .StatusMessage = $"Loaded from user home directory",
                .IsToolOutput = False,
                .OutputFiles = New List(Of String)()
            }

            ' Register in current attachments so other tools can find it
            If _apCurrentAttachments IsNot Nothing Then
                _apCurrentAttachments.Add(info)
            End If

            Return info
        Catch
            Return Nothing
        End Try
    End Function

    ' ═══════════════════════════════════════════════════════════════════════════
    '  ADMIN / DASHBOARD HELPERS
    ' ═══════════════════════════════════════════════════════════════════════════

    ''' <summary>
    ''' Lists all user storage directories (for the admin dashboard).
    ''' Returns tuples of (sanitizedFolderName, hasMemory, homeFileCount, homeSizeBytes).
    ''' </summary>
    Friend Function ListAllUserStorageDirs() As List(Of (FolderName As String, HasMemory As Boolean, HomeFileCount As Integer, HomeSizeBytes As Long))
        Dim result As New List(Of (String, Boolean, Integer, Long))()
        Dim rootDir = GetUserStorageRootDir()
        If Not Directory.Exists(rootDir) Then Return result
        For Each userDir In Directory.GetDirectories(rootDir)
            Dim folderName = Path.GetFileName(userDir)
            Dim hasMemory = File.Exists(Path.Combine(userDir, AP_UserMemoryFileName))
            Dim homeDir = Path.Combine(userDir, AP_UserHomeSubdir)
            Dim fileCount = 0
            Dim totalSize As Long = 0
            If Directory.Exists(homeDir) Then
                Dim files = Directory.GetFiles(homeDir)
                fileCount = files.Length
                For Each f In files
                    totalSize += New FileInfo(f).Length
                Next
            End If
            result.Add((folderName, hasMemory, fileCount, totalSize))
        Next
        Return result
    End Function

    ''' <summary>Reads the raw content of a user's memory file (for admin editing).</summary>
    Friend Function ReadUserMemoryContent(senderEmail As String) As String
        Dim filePath = GetUserMemoryFilePath(senderEmail)
        If Not File.Exists(filePath) Then Return ""
        Return ReadFileWithRetry(filePath)
    End Function

    ''' <summary>Writes raw content to a user's memory file (for admin editing).</summary>
    Friend Sub WriteUserMemoryContent(senderEmail As String, content As String)
        Dim filePath = GetUserMemoryFilePath(senderEmail)
        Dim folder = Path.GetDirectoryName(filePath)
        If Not Directory.Exists(folder) Then Directory.CreateDirectory(folder)
        WriteFileWithRetry(filePath, content)
    End Sub

    ''' <summary>Deletes a specific file from a user's home directory (admin).</summary>
    Friend Function AdminDeleteUserFile(senderEmail As String, fileName As String) As Boolean
        Try
            Dim safePath = ResolveUserHomeFileSafe(senderEmail, fileName)
            If safePath IsNot Nothing AndAlso File.Exists(safePath) Then
                File.Delete(safePath)
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    ''' <summary>Deletes all files in a user's home directory (admin).</summary>
    Friend Sub AdminDeleteAllUserFiles(senderEmail As String)
        Try
            Dim homeDir = GetUserHomeDir(senderEmail)
            If Directory.Exists(homeDir) Then
                For Each filePath In Directory.GetFiles(homeDir)
                    Try : File.Delete(filePath) : Catch : End Try
                Next
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Deletes a user's memory file (admin).</summary>
    Friend Sub AdminDeleteUserMemory(senderEmail As String)
        DisableUserMemory(senderEmail)
    End Sub

    ''' <summary>Deletes the entire user directory including memory and all home files (admin).</summary>
    Friend Sub AdminDeleteUserStorage(senderEmail As String)
        Try
            Dim sanitized = SanitizeEmailToFolderName(senderEmail)
            Dim userDir = Path.Combine(GetUserStorageRootDir(), sanitized)
            If Directory.Exists(userDir) Then
                Directory.Delete(userDir, recursive:=True)
            End If
        Catch
        End Try
    End Sub

End Class