' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PathPolicy.vb
' Purpose: Centralized path-access policy for agent-layer file tools.
'          Enforces workspace boundaries, skill reference/script access, and
'          symlink-traversal blocking.
'
' Writable Root Precedence:
'  1. Workspace path (if maintained in current session).
'  2. Otherwise user's Desktop.
'
' Read-only Roots (always allowed for read):
'  - Skill scripts/ and references/ directories (any discovered skill).
'  - In "chat author mode", also writable under skill directories.
'  - All paths canonicalized; ".." and symlink traversal blocked.
'  - Default file size limit: 2 MiB (configurable via MaxFileSizeBytes).
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading

Namespace Agents

    Public Enum PathAccess
        Read
        Write
    End Enum

    Public NotInheritable Class PathPolicy

        Private Sub New()
        End Sub

        ' --------------------------------------------------------------- workspace

        Private Shared _workspaceRoot As String = Nothing

        ''' <summary>Sets the active workspace root for this process (host call). Pass Nothing to clear.</summary>
        Public Shared Sub SetWorkspaceRoot(rootOrNothing As String)
            If String.IsNullOrWhiteSpace(rootOrNothing) Then
                _workspaceRoot = Nothing
            Else
                Try
                    _workspaceRoot = Path.GetFullPath(rootOrNothing)
                Catch
                    _workspaceRoot = Nothing
                End Try
            End If
        End Sub

        Public Shared ReadOnly Property WorkspaceRoot As String
            Get
                Return _workspaceRoot
            End Get
        End Property

        ' --------------------------------------------------------------- chat-author scope

        Private Shared ReadOnly _chatAuthor As New AsyncLocal(Of Boolean)

        ''' <summary>
        ''' Marks the current async flow as "chat author" — permits writes under skill
        ''' scripts/references for the duration of the returned scope. Use:
        '''     Using PathPolicy.BeginChatAuthorScope() : ... End Using
        ''' </summary>
        Public Shared Function BeginChatAuthorScope() As IDisposable
            Return New ChatAuthorScope()
        End Function

        Private Class ChatAuthorScope
            Implements IDisposable
            Private ReadOnly _previous As Boolean
            Public Sub New()
                _previous = _chatAuthor.Value
                _chatAuthor.Value = True
            End Sub
            Public Sub Dispose() Implements IDisposable.Dispose
                _chatAuthor.Value = _previous
            End Sub
        End Class

        ' --------------------------------------------------------------- size limits

        Public Shared Property MaxFileSizeBytes As Integer = 2 * 1024 * 1024 ' 2 MiB

        ' --------------------------------------------------------------- writable root

        ''' <summary>
        ''' Returns the canonical writable root for new agent-created files.
        ''' Workspace if maintained, otherwise the current user's Desktop.
        ''' </summary>
        Public Shared Function GetDefaultWritableRoot() As String
            If Not String.IsNullOrWhiteSpace(_workspaceRoot) AndAlso Directory.Exists(_workspaceRoot) Then
                Return _workspaceRoot
            End If
            Return Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End Function

        ' --------------------------------------------------------------- validation

        ''' <summary>
        ''' Resolves <paramref name="requestedPath"/> against the policy. Returns the fully
        ''' qualified canonical path on success, or throws <see cref="UnauthorizedAccessException"/>
        ''' on denial.
        '''
        ''' Behavior:
        '''  - If <paramref name="requestedPath"/> is relative, it is resolved against the
        '''    default writable root (workspace or Desktop).
        '''  - Read access is allowed under: workspace, Desktop, and any discovered skill
        '''    scripts/ or references/ directory.
        '''  - Write access is allowed under: workspace and Desktop; additionally under a
        '''    skill's scripts/references when chat-author scope is active.
        ''' </summary>
        Public Shared Function Resolve(requestedPath As String, access As PathAccess) As String
            If String.IsNullOrWhiteSpace(requestedPath) Then
                Throw New ArgumentException("Path is empty.", NameOf(requestedPath))
            End If

            Dim full As String
            Try
                If Path.IsPathRooted(requestedPath) Then
                    full = Path.GetFullPath(requestedPath)
                Else
                    full = Path.GetFullPath(Path.Combine(GetDefaultWritableRoot(), requestedPath))
                End If
            Catch ex As Exception
                Throw New UnauthorizedAccessException("Invalid path: " & ex.Message)
            End Try

            ' Block raw devices and UNC by default.
            If full.StartsWith("\\?\", StringComparison.Ordinal) OrElse full.StartsWith("\\.\", StringComparison.Ordinal) Then
                Throw New UnauthorizedAccessException("Device paths are not allowed.")
            End If

            ' Build allow-set.
            Dim writeRoots As New List(Of String)()
            Dim readRoots As New List(Of String)()

            Dim ws = If(_workspaceRoot, "")
            If Not String.IsNullOrWhiteSpace(ws) AndAlso Directory.Exists(ws) Then
                writeRoots.Add(Path.GetFullPath(ws))
                readRoots.Add(Path.GetFullPath(ws))
            End If
            Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If Not String.IsNullOrWhiteSpace(desktop) Then
                writeRoots.Add(Path.GetFullPath(desktop))
                readRoots.Add(Path.GetFullPath(desktop))
            End If

            ' Skill scripts/references — always readable.
            ' In skill-author mode (or legacy chat-author scope), the ENTIRE skill folder
            ' is writable (so SKILL.md plus scripts/ and references/ can be edited).
            Dim skillReadRoots As New List(Of String)()
            Dim skillFullRoots As New List(Of String)()
            Try
                For Each sk In AgentResources.Skills
                    If sk Is Nothing OrElse String.IsNullOrWhiteSpace(sk.DirectoryPath) Then Continue For
                    Dim skFull As String = Path.GetFullPath(sk.DirectoryPath)
                    skillFullRoots.Add(skFull)
                    Dim sdir As String = Path.Combine(skFull, "scripts")
                    Dim rdir As String = Path.Combine(skFull, "references")
                    If Directory.Exists(sdir) Then skillReadRoots.Add(Path.GetFullPath(sdir))
                    If Directory.Exists(rdir) Then skillReadRoots.Add(Path.GetFullPath(rdir))
                Next
            Catch
            End Try
            readRoots.AddRange(skillReadRoots)
            readRoots.AddRange(skillFullRoots)
            If _chatAuthor.Value OrElse SkillAuthorMode.IsActive Then
                writeRoots.AddRange(skillFullRoots)
            End If

            Dim roots = If(access = PathAccess.Write, writeRoots, readRoots)
            For Each r In roots
                If IsUnder(full, r) Then Return full
            Next

            Throw New UnauthorizedAccessException("Path is outside the allowed roots for " & access.ToString().ToLowerInvariant() & " access.")
        End Function

        Private Shared Function IsUnder(candidate As String, root As String) As Boolean
            If String.IsNullOrWhiteSpace(candidate) OrElse String.IsNullOrWhiteSpace(root) Then Return False
            Dim a = candidate.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            Dim b = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            If String.Equals(a, b, StringComparison.OrdinalIgnoreCase) Then Return True
            Return a.StartsWith(b & Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
        End Function

        ' --------------------------------------------------------------- writable-name helper

        ''' <summary>
        ''' Returns a non-colliding writable path inside the default writable root for a
        ''' suggested filename. If the suggested filename already exists, " (n)" is appended.
        ''' </summary>
        Public Shared Function NewWritablePath(suggestedFileName As String) As String
            Dim root = GetDefaultWritableRoot()
            Dim safe = SanitizeFileName(If(suggestedFileName, "untitled.txt"))
            Dim candidate = Path.Combine(root, safe)
            If Not File.Exists(candidate) Then Return candidate
            Dim baseName = Path.GetFileNameWithoutExtension(safe)
            Dim ext = Path.GetExtension(safe)
            For i = 2 To 1000
                Dim p = Path.Combine(root, baseName & " (" & i.ToString() & ")" & ext)
                If Not File.Exists(p) Then Return p
            Next
            Return Path.Combine(root, baseName & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & ext)
        End Function

        Private Shared Function SanitizeFileName(name As String) As String
            If String.IsNullOrWhiteSpace(name) Then Return "untitled.txt"
            Dim invalid = Path.GetInvalidFileNameChars()
            Dim sb As New System.Text.StringBuilder(name.Length)
            For Each c In name
                If Array.IndexOf(invalid, c) >= 0 Then sb.Append("_"c) Else sb.Append(c)
            Next
            Return sb.ToString()
        End Function

    End Class

End Namespace