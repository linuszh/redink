' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.InkyMemory.vb
' Purpose: Manages InkyMemory files — persistent, natural-language memory stores
'          for user preferences, recurring corrections, and working style.
'
'          Supports two modes:
'          (1) Local single-user memory at %AppData%\redink\RI_memory.txt
'              shared by Word, Excel, and Outlook chatbots on the same machine.
'          (2) Per-user memory at an arbitrary file path (used by AutoPilot
'              for multi-user server-side memory).
'
' Architecture / How it works:
'  - The memory file is a plain text file containing one memory item per line,
'    prefixed with "- ".
'  - ReadInkyMemory() / ReadInkyMemoryFromFile() read and return content
'    (capped to INI_InkyMemoryCap items).
'  - ProcessInkyMemoryResponse() / ProcessInkyMemoryResponseForFile() parse
'    <INKY_MEMORY> blocks from LLM responses, apply ADD/REMOVE/AMEND operations,
'    and return the cleaned response.
'  - StripInkyMemoryBlock() removes the <INKY_MEMORY> block from displayed output.
'  - File access uses retry-based locking to handle concurrent access.
'
' Memory file format (natural language, user-editable):
'    # Inky Memory
'    # Last updated: 2026-03-30 14:22:15
'    # This file is automatically maintained. You can edit it manually.
'    
'    - The user prefers British English spelling in formal documents.
'    - Never suggest contractions in legal text.
'
' LLM output format (appended at end of response):
'    <INKY_MEMORY>
'    ADD: The user prefers British English spelling.
'    REMOVE: The user prefers American English spelling.
'    AMEND: Use formal salutations → Use formal salutations, but only for external emails
'    </INKY_MEMORY>
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' ═══════════════════════════════════════════════════════════════════════════
        '  LOCAL SINGLE-USER API (existing — delegates to generalised methods)
        ' ═══════════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Returns the full path to the local InkyMemory file: %AppData%\redink\RI_memory.txt
        ''' </summary>
        Public Shared Function GetInkyMemoryFilePath() As String
            Dim folder As String = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), InkyMemoryFolder)
            Return Path.Combine(folder, InkyMemoryFileName)
        End Function

        ''' <summary>
        ''' Reads the local InkyMemory file and returns its content as a string suitable for
        ''' inclusion in a system prompt. Returns empty string if the file does not exist,
        ''' is empty, or an error occurs. Items are capped to <paramref name="maxItems"/>.
        ''' </summary>
        ''' <param name="maxItems">Maximum number of memory items to include (from INI_InkyMemoryCap).</param>
        ''' <returns>Memory content string, or empty string if unavailable.</returns>
        Public Shared Function ReadInkyMemory(maxItems As Integer) As String
            Return ReadInkyMemoryFromFile(GetInkyMemoryFilePath(), maxItems)
        End Function

        ''' <summary>
        ''' Processes an LLM response that may contain an &lt;INKY_MEMORY&gt; block.
        ''' Parses ADD/REMOVE/AMEND operations, applies them to the local memory file,
        ''' and returns the response with the memory block stripped.
        ''' </summary>
        ''' <param name="llmResponse">The raw LLM response text.</param>
        ''' <param name="maxItems">Maximum number of memory items to retain (from INI_InkyMemoryCap).</param>
        ''' <returns>The cleaned response with the INKY_MEMORY block removed.</returns>
        Public Shared Function ProcessInkyMemoryResponse(llmResponse As String, maxItems As Integer) As String
            Return ProcessInkyMemoryResponseForFile(GetInkyMemoryFilePath(), llmResponse, maxItems)
        End Function

        ''' <summary>
        ''' Strips the &lt;INKY_MEMORY&gt; block from a response without applying any operations.
        ''' Use this when you only need to clean the display text.
        ''' </summary>
        ''' <param name="llmResponse">The raw LLM response text.</param>
        ''' <returns>Response with the INKY_MEMORY block removed.</returns>
        Public Shared Function StripInkyMemoryBlock(llmResponse As String) As String
            If String.IsNullOrWhiteSpace(llmResponse) Then Return If(llmResponse, "")
            Dim dummy As String = Nothing
            Return ExtractAndStripMemoryBlock(llmResponse, dummy)
        End Function

        ''' <summary>
        ''' Opens the local InkyMemory file in the shared text editor, creating it with
        ''' a default header if it does not yet exist.
        ''' </summary>
        ''' <param name="_context">Optional shared context (kept for API consistency).</param>
        Public Shared Sub EditInkyMemoryFile(Optional _context As ISharedContext = Nothing,
                                     Optional ownerHandle As IntPtr = Nothing)
            Try
                Dim filePath As String = GetInkyMemoryFilePath()
                Dim folder As String = Path.GetDirectoryName(filePath)

                ' Ensure directory exists
                If Not Directory.Exists(folder) Then
                    Directory.CreateDirectory(folder)
                End If

                ' Create file with default header if it doesn't exist
                If Not File.Exists(filePath) Then
                    WriteFileWithRetry(filePath, GetDefaultMemoryFileContent())
                End If

                ShowTextFileEditor(
            filePath,
            "Inky Memory — Edit your persistent preferences and learning items below.",
            ownerHandle:=ownerHandle
        )
            Catch ex As Exception
                ShowCustomMessageBox($"Could not open memory file: {ex.Message}")
            End Try
        End Sub

        ' ═══════════════════════════════════════════════════════════════════════════
        '  GENERALISED FILE-PATH API (used by AutoPilot per-user memory)
        ' ═══════════════════════════════════════════════════════════════════════════

        ''' <summary>
        ''' Reads an InkyMemory file at an arbitrary path and returns its content as a string
        ''' suitable for inclusion in a system prompt. Returns empty string if the file does
        ''' not exist, is empty, or an error occurs. Items are capped to <paramref name="maxItems"/>.
        ''' </summary>
        ''' <param name="filePath">Full path to the memory file.</param>
        ''' <param name="maxItems">Maximum number of memory items to include.</param>
        ''' <returns>Memory content string, or empty string if unavailable.</returns>
        Public Shared Function ReadInkyMemoryFromFile(filePath As String, maxItems As Integer) As String
            Try
                If Not File.Exists(filePath) Then Return ""

                Dim content As String = ReadFileWithRetry(filePath)
                If String.IsNullOrWhiteSpace(content) Then Return ""

                ' Parse items (lines starting with "- "), cap to maxItems
                Dim lines As String() = content.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None)
                Dim items As New List(Of String)()
                For Each line In lines
                    Dim trimmed = line.Trim()
                    If trimmed.StartsWith("- ") AndAlso trimmed.Length > 2 Then
                        items.Add(trimmed)
                        If items.Count >= maxItems Then Exit For
                    End If
                Next

                If items.Count = 0 Then Return ""

                Return String.Join(Environment.NewLine, items)

            Catch
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Processes an LLM response that may contain an &lt;INKY_MEMORY&gt; block.
        ''' Parses ADD/REMOVE/AMEND operations, applies them to the memory file at the
        ''' specified path, and returns the response with the memory block stripped.
        ''' </summary>
        ''' <param name="filePath">Full path to the memory file to update.</param>
        ''' <param name="llmResponse">The raw LLM response text.</param>
        ''' <param name="maxItems">Maximum number of memory items to retain.</param>
        ''' <returns>The cleaned response with the INKY_MEMORY block removed.</returns>
        Public Shared Function ProcessInkyMemoryResponseForFile(filePath As String, llmResponse As String, maxItems As Integer) As String
            If String.IsNullOrWhiteSpace(llmResponse) Then Return If(llmResponse, "")

            ' Extract the <INKY_MEMORY> block
            Dim memoryBlock As String = Nothing
            Dim cleanedResponse As String = ExtractAndStripMemoryBlock(llmResponse, memoryBlock)

            ' If no memory block found, return original
            If String.IsNullOrWhiteSpace(memoryBlock) Then Return llmResponse

            ' Parse operations from the block
            Dim operations As List(Of MemoryOperation) = ParseMemoryOperations(memoryBlock)
            If operations.Count = 0 Then Return cleanedResponse

            ' Apply operations to the memory file
            Try
                ApplyMemoryOperationsToFile(filePath, operations, maxItems)
            Catch
                ' Silently fail — memory update is best-effort
            End Try

            Return cleanedResponse
        End Function

        ''' <summary>
        ''' Applies parsed memory operations to a memory file at the specified path.
        ''' Reads the current file, applies ADD/REMOVE/AMEND, caps to maxItems, and writes back.
        ''' Creates the file and parent directories if they do not exist.
        ''' </summary>
        ''' <param name="filePath">Full path to the memory file.</param>
        ''' <param name="operations">List of parsed memory operations to apply.</param>
        ''' <param name="maxItems">Maximum number of memory items to retain.</param>
        Public Shared Sub ApplyMemoryOperationsToFile(filePath As String, operations As List(Of MemoryOperation), maxItems As Integer)
            Dim folder As String = Path.GetDirectoryName(filePath)

            ' Ensure directory exists
            If Not Directory.Exists(folder) Then
                Directory.CreateDirectory(folder)
            End If

            ' Read existing items
            Dim items As New List(Of String)()
            If File.Exists(filePath) Then
                Dim content = ReadFileWithRetry(filePath)
                If Not String.IsNullOrWhiteSpace(content) Then
                    Dim lines = content.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None)
                    For Each line In lines
                        Dim trimmed = line.Trim()
                        If trimmed.StartsWith("- ") AndAlso trimmed.Length > 2 Then
                            items.Add(trimmed.Substring(2).Trim())
                        End If
                    Next
                End If
            End If

            ' Apply operations in order
            For Each op In operations
                Select Case op.Type
                    Case MemoryOperation.OpType.Add
                        ' Only add if not already present (fuzzy check)
                        If Not items.Any(Function(i) i.Equals(op.Value, StringComparison.OrdinalIgnoreCase)) Then
                            items.Add(op.Value)
                        End If

                    Case MemoryOperation.OpType.Remove
                        ' Remove by fuzzy match (case-insensitive contains)
                        Dim toRemove = items.Where(Function(i)
                                                       Return i.Equals(op.Value, StringComparison.OrdinalIgnoreCase) OrElse
                                                              i.IndexOf(op.Value, StringComparison.OrdinalIgnoreCase) >= 0
                                                   End Function).ToList()
                        For Each item In toRemove
                            items.Remove(item)
                        Next

                    Case MemoryOperation.OpType.Amend
                        ' Find best match and replace
                        Dim matchIdx = items.FindIndex(Function(i)
                                                           Return i.Equals(op.Value, StringComparison.OrdinalIgnoreCase) OrElse
                                                                  i.IndexOf(op.Value, StringComparison.OrdinalIgnoreCase) >= 0
                                                       End Function)
                        If matchIdx >= 0 Then
                            items(matchIdx) = op.NewValue
                        Else
                            ' If old item not found, treat as ADD
                            items.Add(op.NewValue)
                        End If
                End Select
            Next

            ' Cap to maxItems (keep most recent = last items)
            If items.Count > maxItems Then
                items = items.Skip(items.Count - maxItems).ToList()
            End If

            ' Build file content
            Dim sb As New StringBuilder()
            sb.AppendLine("# Inky Memory")
            sb.AppendLine($"# Last updated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine("# This file is automatically maintained. You can edit it manually.")
            sb.AppendLine()
            For Each item In items
                sb.AppendLine($"- {item}")
            Next

            WriteFileWithRetry(filePath, sb.ToString())
        End Sub

#Region "Private Helpers"

        ''' <summary>
        ''' Regex pattern matching the INKY_MEMORY block (case-insensitive, singleline).
        ''' </summary>
        Private Shared ReadOnly MemoryBlockPattern As New Regex(
            "<INKY_MEMORY>\s*(.*?)\s*</INKY_MEMORY>",
            RegexOptions.Singleline Or RegexOptions.IgnoreCase)

        ''' <summary>
        ''' Represents a single parsed memory operation.
        ''' </summary>
        Public Class MemoryOperation
            Public Enum OpType
                Add
                Remove
                Amend
            End Enum

            Public Property [Type] As OpType
            Public Property Value As String
            ''' <summary>For AMEND: the new replacement text.</summary>
            Public Property NewValue As String
        End Class

        ''' <summary>
        ''' Extracts the INKY_MEMORY block content and returns the response with the block removed.
        ''' </summary>
        Private Shared Function ExtractAndStripMemoryBlock(response As String, ByRef blockContent As String) As String
            blockContent = Nothing

            Dim m As Match = MemoryBlockPattern.Match(response)
            If Not m.Success Then Return response

            blockContent = m.Groups(1).Value.Trim()

            ' Remove the block and any surrounding whitespace/newlines
            Dim cleaned As String = response.Remove(m.Index, m.Length).TrimEnd()
            Return cleaned
        End Function

        ''' <summary>
        ''' Parses ADD, REMOVE, and AMEND operations from the raw memory block text.
        ''' </summary>
        Public Shared Function ParseMemoryOperations(blockText As String) As List(Of MemoryOperation)
            Dim ops As New List(Of MemoryOperation)()
            If String.IsNullOrWhiteSpace(blockText) Then Return ops

            Dim lines As String() = blockText.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                Dim trimmed = line.Trim()

                If trimmed.StartsWith("ADD:", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = trimmed.Substring(4).Trim()
                    If value.Length > 0 Then
                        ops.Add(New MemoryOperation With {.Type = MemoryOperation.OpType.Add, .Value = value})
                    End If

                ElseIf trimmed.StartsWith("REMOVE:", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = trimmed.Substring(7).Trim()
                    If value.Length > 0 Then
                        ops.Add(New MemoryOperation With {.Type = MemoryOperation.OpType.Remove, .Value = value})
                    End If

                ElseIf trimmed.StartsWith("AMEND:", StringComparison.OrdinalIgnoreCase) Then
                    Dim value = trimmed.Substring(6).Trim()
                    ' Parse "old text → new text" (supports → and ->)
                    Dim separatorIdx = value.IndexOf("→", StringComparison.Ordinal)
                    If separatorIdx < 0 Then separatorIdx = value.IndexOf("->", StringComparison.Ordinal)

                    If separatorIdx > 0 Then
                        Dim sepLen = If(value.Substring(separatorIdx).StartsWith("→"), 1, 2)
                        Dim oldText = value.Substring(0, separatorIdx).Trim()
                        Dim newText = value.Substring(separatorIdx + sepLen).Trim()
                        If oldText.Length > 0 AndAlso newText.Length > 0 Then
                            ops.Add(New MemoryOperation With {
                                .Type = MemoryOperation.OpType.Amend,
                                .Value = oldText,
                                .NewValue = newText
                            })
                        End If
                    End If
                End If
            Next

            Return ops
        End Function

        ''' <summary>
        ''' Applies parsed memory operations to the local memory file with file-level locking.
        ''' Reads the current file, applies ADD/REMOVE/AMEND, caps to maxItems, and writes back.
        ''' </summary>
        Private Shared Sub ApplyMemoryOperations(operations As List(Of MemoryOperation), maxItems As Integer)
            ApplyMemoryOperationsToFile(GetInkyMemoryFilePath(), operations, maxItems)
        End Sub

        ''' <summary>
        ''' Returns the default content for a newly created memory file.
        ''' </summary>
        Public Shared Function GetDefaultMemoryFileContent() As String
            Dim sb As New StringBuilder()
            sb.AppendLine("# Inky Memory")
            sb.AppendLine($"# Last updated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine("# This file is automatically maintained. You can edit it manually.")
            sb.AppendLine("# Add one memory item per line, prefixed with ""- "".")
            sb.AppendLine()
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Reads a file with retry logic to handle concurrent access from multiple Office add-ins.
        ''' </summary>
        Public Shared Function ReadFileWithRetry(filePath As String, Optional maxRetries As Integer = 3) As String
            For attempt = 1 To maxRetries
                Try
                    Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                        Using sr As New StreamReader(fs, Encoding.UTF8)
                            Return sr.ReadToEnd()
                        End Using
                    End Using
                Catch ex As IOException When attempt < maxRetries
                    System.Threading.Thread.Sleep(50 * attempt)
                End Try
            Next
            Return ""
        End Function

        ''' <summary>
        ''' Writes a file with retry logic to handle concurrent access from multiple Office add-ins.
        ''' </summary>
        Public Shared Sub WriteFileWithRetry(filePath As String, content As String, Optional maxRetries As Integer = 3)
            For attempt = 1 To maxRetries
                Try
                    Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.Read)
                        Using sw As New StreamWriter(fs, Encoding.UTF8)
                            sw.Write(content)
                        End Using
                    End Using
                    Return
                Catch ex As IOException When attempt < maxRetries
                    System.Threading.Thread.Sleep(50 * attempt)
                End Try
            Next
        End Sub

#End Region

    End Class
End Namespace