' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SessionMemory.vb
' Purpose: Process-wide session-scoped key/value memory that the agent can write
'          to and read from. Persists to disk under the user-local agent resources
'          path so it survives Outlook/Word restarts until the user clears it.
'
' Storage layout:
'   <INI_AgentResourcesPathLocal>\.session\memory.json
'   (Falls back to %LOCALAPPDATA%\RedInk\.session\memory.json if the local path is unset.)
'
' Each entry has:
'   key       - opaque, stable identifier (assistant-chosen or auto-generated)
'   summary   - short textual stub the model sees in subsequent turns
'   value     - the actual payload (any JSON value or string)
'   createdAt - UTC ISO-8601
'   tags      - optional list of tags
'
' Thread-safety: a coarse SyncLock protects all mutations and the on-disk write.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Agents

    Public Class SessionMemoryEntry
        Public Property Key As String
        Public Property Summary As String
        Public Property Value As JToken
        Public Property CreatedAt As DateTime
        Public Property Tags As List(Of String)
    End Class

    Public NotInheritable Class SessionMemory

        Private Sub New()
        End Sub

        Private Shared ReadOnly _sync As New Object()
        Private Shared _loaded As Boolean
        Private Shared _entries As New Dictionary(Of String, SessionMemoryEntry)(StringComparer.Ordinal)

        ' ------------------------------------------------------------------ persistence

        Private Shared Function GetStorePath() As String
            Dim root As String = TryGetSharedPath("INI_AgentResourcesPathLocal")
            If String.IsNullOrWhiteSpace(root) Then
                root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "RedInk")
            End If
            Dim dir = Path.Combine(root, ".session")
            Try
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Catch
            End Try
            Return Path.Combine(dir, "memory.json")
        End Function

        Private Shared Sub EnsureLoaded()
            If _loaded Then Return
            SyncLock _sync
                If _loaded Then Return
                _loaded = True
                Try
                    Dim p = GetStorePath()
                    If File.Exists(p) Then
                        Dim raw = File.ReadAllText(p, Encoding.UTF8)
                        Dim list = JsonConvert.DeserializeObject(Of List(Of SessionMemoryEntry))(raw)
                        If list IsNot Nothing Then
                            For Each e In list
                                If e IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Key) Then
                                    _entries(e.Key) = e
                                End If
                            Next
                        End If
                    End If
                Catch
                    ' ignore corrupt memory; start clean
                    _entries = New Dictionary(Of String, SessionMemoryEntry)(StringComparer.Ordinal)
                End Try
            End SyncLock
        End Sub

        Private Shared Sub SaveUnlocked()
            Try
                Dim p = GetStorePath()
                Dim list = _entries.Values.OrderBy(Function(e) e.CreatedAt).ToList()
                Dim json = JsonConvert.SerializeObject(list, Formatting.Indented)
                File.WriteAllText(p, json, Encoding.UTF8)
            Catch
                ' best-effort
            End Try
        End Sub

        ' ------------------------------------------------------------------ API

        Public Shared Function Put(key As String,
                                   summary As String,
                                   value As JToken,
                                   Optional tags As IEnumerable(Of String) = Nothing) As SessionMemoryEntry
            EnsureLoaded()
            SyncLock _sync
                If String.IsNullOrWhiteSpace(key) Then key = "mem_" & Guid.NewGuid().ToString("N").Substring(0, 12)
                Dim e As New SessionMemoryEntry With {
                    .Key = key,
                    .Summary = If(summary, ""),
                    .Value = If(value, JValue.CreateNull()),
                    .CreatedAt = DateTime.UtcNow,
                    .Tags = If(tags Is Nothing, New List(Of String)(), tags.Where(Function(t) Not String.IsNullOrWhiteSpace(t)).ToList())
                }
                _entries(key) = e
                SaveUnlocked()
                Return e
            End SyncLock
        End Function

        Public Shared Function [Get](key As String) As SessionMemoryEntry
            EnsureLoaded()
            SyncLock _sync
                Dim e As SessionMemoryEntry = Nothing
                _entries.TryGetValue(If(key, ""), e)
                Return e
            End SyncLock
        End Function

        Public Shared Function List() As List(Of SessionMemoryEntry)
            EnsureLoaded()
            SyncLock _sync
                Return _entries.Values.OrderBy(Function(e) e.CreatedAt).ToList()
            End SyncLock
        End Function

        Public Shared Function Delete(key As String) As Boolean
            EnsureLoaded()
            SyncLock _sync
                If String.IsNullOrWhiteSpace(key) Then Return False
                Dim removed = _entries.Remove(key)
                If removed Then SaveUnlocked()
                Return removed
            End SyncLock
        End Function

        Public Shared Sub Clear()
            EnsureLoaded()
            SyncLock _sync
                _entries.Clear()
                SaveUnlocked()
            End SyncLock
        End Sub

        ''' <summary>
        ''' Produces a one-line stub the model can see in lieu of a full value.
        ''' Used when tool results are large and we want context to stay lean.
        ''' </summary>
        Public Shared Function BuildStub(entry As SessionMemoryEntry) As String
            If entry Is Nothing Then Return ""
            Dim summary = If(entry.Summary, "").Trim()
            If summary.Length = 0 Then summary = "(stored in session memory)"
            Return "[memory:" & entry.Key & "] " & summary
        End Function

        ' ------------------------------------------------------------------ shared-properties bridge

        Private Shared Function TryGetSharedPath(propertyName As String) As String
            Try
                Dim asm = GetType(SharedLibrary.SharedContext).Assembly
                For Each typeFullName In {"SharedLibrary.SharedProperties", "SharedLibrary.SharedContext"}
                    Dim t = asm.GetType(typeFullName, throwOnError:=False, ignoreCase:=False)
                    If t Is Nothing Then Continue For
                    Dim pi = t.GetProperty(propertyName,
                        Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static Or Reflection.BindingFlags.Instance)
                    If pi Is Nothing Then Continue For
                    Dim val As Object = Nothing
                    Dim getter = pi.GetGetMethod()
                    If getter IsNot Nothing AndAlso getter.IsStatic Then
                        val = pi.GetValue(Nothing, Nothing)
                    End If
                    If TypeOf val Is String AndAlso Not String.IsNullOrWhiteSpace(CStr(val)) Then
                        Return CStr(val)
                    End If
                Next
            Catch
            End Try
            Return Nothing
        End Function

    End Class

End Namespace