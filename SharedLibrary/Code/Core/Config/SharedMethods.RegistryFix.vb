' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.RegistryFix.vb
' Purpose: Applies administrator-configured registry entries from the INI file
'          to harden the Office add-in against being disabled (e.g., slow-load
'          resiliency) or to enforce specific load behaviors.
'
' Architecture:
'  - INI keys prefixed with "RegFix_" define registry operations.
'  - Each entry has the format: Host|RegPath|ValueName|ValueKind|Data
'  - Supported tokens in RegPath/ValueName/Data:
'      {ProgID}      → VSTO ProgID of the current add-in (e.g., "Red Ink for Outlook")
'      {Host}        → Office host application name (e.g., "Outlook")
'      {HostVersion} → Office major version as "nn.0" (e.g., "16.0")
'  - Only entries whose Host field matches the current host (or "All") are applied.
'  - Only writes to HKEY_CURRENT_USER for safety.
'  - Idempotent: skips writes when the existing value already matches.
'
' Sample INI entries:
' ; Prevent Office from disabling the add-in when it loads slowly (resiliency protection)
' RegFix_1 = Outlook|HKEY_CURRENT_USER\Software\Microsoft\Office\{HostVersion}\Outlook\Resiliency\DoNotDisableAddinList|{ProgID}|DWORD|1
' RegFix_2 = Word|HKEY_CURRENT_USER\Software\Microsoft\Office\{HostVersion}\Word\Resiliency\DoNotDisableAddinList|{ProgID}|DWORD|1
' RegFix_3 = Excel|HKEY_CURRENT_USER\Software\Microsoft\Office\{HostVersion}\Excel\Resiliency\DoNotDisableAddinList|{ProgID}|DWORD|1
'
' ; Ensure add-ins load at startup (LoadBehavior=3)
' RegFix_4 = Outlook|HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\{ProgID}|LoadBehavior|DWORD|3
' RegFix_5 = Word|HKEY_CURRENT_USER\Software\Microsoft\Office\Word\Addins\{ProgID}|LoadBehavior|DWORD|3
' RegFix_6 = Excel|HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\{ProgID}|LoadBehavior|DWORD|3
' =============================================================================

Option Strict On
Option Explicit On

Imports Microsoft.Win32

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ' ProgID as registered in the Windows registry by the VSTO installer.
        ' Derived from the AssemblyName in each .vbproj (not from RootNamespace, which uses underscores).
        ' Registry location: HKCU\Software\Microsoft\Office\<Host>\Addins\<ProgID>
        Private Const ProgID_Word As String = "Red Ink for Word"
        Private Const ProgID_Outlook As String = "Red Ink for Outlook"
        Private Const ProgID_Excel As String = "Red Ink for Excel"

        ''' <summary>
        ''' Processes all <c>RegFix_*</c> entries from <paramref name="configDict"/> and applies matching
        ''' registry values for the current host add-in.
        ''' </summary>
        ''' <param name="configDict">INI configuration dictionary.</param>
        ''' <param name="context">Shared context (provides <c>RDV</c> for host detection).</param>
        Public Shared Sub ApplyRegistryFixes(configDict As Dictionary(Of String, String), context As SharedContext.ISharedContext)
            Try
                Dim currentHost As String = GetHostName(context.RDV)
                If String.IsNullOrWhiteSpace(currentHost) Then Return

                Dim progId As String = GetProgID(currentHost)
                If String.IsNullOrWhiteSpace(progId) Then Return

                Dim hostVersion As String = GetOfficeVersion()

                ' Collect all RegFix_ entries, sorted by key for deterministic order.
                Dim regFixEntries As New SortedDictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each kvp In configDict
                    If kvp.Key.StartsWith("RegFix_", StringComparison.OrdinalIgnoreCase) Then
                        regFixEntries(kvp.Key) = kvp.Value
                    End If
                Next

                For Each kvp In regFixEntries
                    Try
                        ApplySingleRegistryFix(kvp.Key, kvp.Value, currentHost, progId, hostVersion)
                    Catch ex As Exception
                        ' Log but don't block startup for a single failed registry fix.
                        System.Diagnostics.Debug.WriteLine($"RegFix warning ({kvp.Key}): {ex.Message}")
                    End Try
                Next

            Catch ex As Exception
                ' Never block config loading due to registry fix errors.
                System.Diagnostics.Debug.WriteLine($"ApplyRegistryFixes error: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Parses and applies a single RegFix entry.
        ''' Format: <c>Host|RegPath|ValueName|ValueKind|Data</c>
        ''' </summary>
        Private Shared Sub ApplySingleRegistryFix(entryKey As String, entryValue As String,
                                                   currentHost As String, progId As String, hostVersion As String)

            ' Parse the pipe-delimited entry.
            Dim parts As String() = entryValue.Split("|"c)
            If parts.Length <> 5 Then
                System.Diagnostics.Debug.WriteLine($"RegFix warning ({entryKey}): Expected 5 pipe-delimited parts, got {parts.Length}. Skipping.")
                Return
            End If

            Dim targetHost As String = parts(0).Trim()
            Dim regPath As String = parts(1).Trim()
            Dim valueName As String = parts(2).Trim()
            Dim valueKindStr As String = parts(3).Trim()
            Dim data As String = parts(4).Trim()

            ' Only apply entries for the current host (or "All").
            If Not targetHost.Equals(currentHost, StringComparison.OrdinalIgnoreCase) AndAlso
               Not targetHost.Equals("All", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If

            ' Replace tokens.
            regPath = ReplaceRegFixTokens(regPath, progId, currentHost, hostVersion)
            valueName = ReplaceRegFixTokens(valueName, progId, currentHost, hostVersion)
            data = ReplaceRegFixTokens(data, progId, currentHost, hostVersion)

            ' Security: Only allow HKEY_CURRENT_USER.
            If Not regPath.StartsWith("HKEY_CURRENT_USER\", StringComparison.OrdinalIgnoreCase) Then
                System.Diagnostics.Debug.WriteLine($"RegFix blocked ({entryKey}): Only HKEY_CURRENT_USER is permitted. Path: {regPath}")
                Return
            End If

            ' Parse the value kind.
            Dim valueKind As RegistryValueKind
            Select Case valueKindStr.ToUpperInvariant()
                Case "DWORD"
                    valueKind = RegistryValueKind.DWord
                Case "STRING", "SZ"
                    valueKind = RegistryValueKind.String
                Case "QWORD"
                    valueKind = RegistryValueKind.QWord
                Case Else
                    System.Diagnostics.Debug.WriteLine($"RegFix warning ({entryKey}): Unsupported ValueKind '{valueKindStr}'. Skipping.")
                    Return
            End Select

            ' Parse the data value according to kind.
            Dim typedData As Object
            Select Case valueKind
                Case RegistryValueKind.DWord
                    Dim dw As Integer
                    If Not Integer.TryParse(data, dw) Then
                        System.Diagnostics.Debug.WriteLine($"RegFix warning ({entryKey}): Cannot parse '{data}' as DWORD. Skipping.")
                        Return
                    End If
                    typedData = dw
                Case RegistryValueKind.QWord
                    Dim qw As Long
                    If Not Long.TryParse(data, qw) Then
                        System.Diagnostics.Debug.WriteLine($"RegFix warning ({entryKey}): Cannot parse '{data}' as QWORD. Skipping.")
                        Return
                    End If
                    typedData = qw
                Case Else
                    typedData = data
            End Select

            ' Extract subkey path (strip "HKEY_CURRENT_USER\").
            Dim subKeyPath As String = regPath.Substring("HKEY_CURRENT_USER\".Length)

            ' Check if the value already matches — skip write if so (idempotent).
            Try
                Using subKey As RegistryKey = Registry.CurrentUser.OpenSubKey(subKeyPath, False)
                    If subKey IsNot Nothing Then
                        Dim existing As Object = subKey.GetValue(valueName, Nothing)
                        If existing IsNot Nothing AndAlso existing.ToString() = typedData.ToString() Then
                            Return ' Already set to desired value.
                        End If
                    End If
                End Using
            Catch
                ' If we can't read, proceed to write.
            End Try

            ' Write the value.
            Using subKey As RegistryKey = Registry.CurrentUser.CreateSubKey(subKeyPath, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue(valueName, typedData, valueKind)
                End If
            End Using

        End Sub

        ''' <summary>
        ''' Replaces <c>{ProgID}</c>, <c>{Host}</c>, and <c>{HostVersion}</c> tokens in <paramref name="input"/>.
        ''' </summary>
        Private Shared Function ReplaceRegFixTokens(input As String, progId As String, host As String, hostVersion As String) As String
            Dim result As String = input
            result = result.Replace("{ProgID}", progId)
            result = result.Replace("{Host}", host)
            result = result.Replace("{HostVersion}", hostVersion)
            Return result
        End Function

        ''' <summary>
        ''' Extracts the host application name (Word, Outlook, Excel) from <c>context.RDV</c>.
        ''' RDV format is e.g. "Outlook (V.070326)" — the host is the first word before the space.
        ''' </summary>
        Private Shared Function GetHostName(rdv As String) As String
            If String.IsNullOrWhiteSpace(rdv) Then Return ""
            Dim spaceIdx As Integer = rdv.IndexOf(" "c)
            If spaceIdx > 0 Then
                Return rdv.Substring(0, spaceIdx)
            End If
            Return rdv
        End Function

        ''' <summary>
        ''' Returns the VSTO ProgID for the given host.
        ''' The ProgID matches the AssemblyName from the .vbproj and is what Office registers
        ''' under <c>HKCU\Software\Microsoft\Office\{Host}\Addins\</c>.
        ''' </summary>
        Private Shared Function GetProgID(host As String) As String
            Select Case host.ToLower()
                Case "word"
                    Return ProgID_Word
                Case "outlook"
                    Return ProgID_Outlook
                Case "excel"
                    Return ProgID_Excel
                Case Else
                    Return ""
            End Select
        End Function

        ''' <summary>
        ''' Detects the installed Office major version from the registry (e.g., "16.0" for Office 2016/2019/2021/365).
        ''' Falls back to "16.0" if detection fails.
        ''' </summary>
        Private Shared Function GetOfficeVersion() As String
            ' Office 2016+ (including 365) registers under 16.0.
            ' Check in descending order for future-proofing.
            Dim candidates As String() = {"17.0", "16.0", "15.0", "14.0"}

            For Each ver In candidates
                Try
                    Using key As RegistryKey = Registry.CurrentUser.OpenSubKey($"Software\Microsoft\Office\{ver}\Common\General", False)
                        If key IsNot Nothing Then Return ver
                    End Using
                Catch
                End Try
            Next

            ' Fallback: virtually all current installations are 16.0.
            Return "16.0"
        End Function

    End Class
End Namespace