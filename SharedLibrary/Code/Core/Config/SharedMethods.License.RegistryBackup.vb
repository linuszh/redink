' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.License.RegistryBackup.vb
' Purpose: Registry-based backup/restore of Pro license credentials to survive
'          user.config (My.Settings) deletions caused by VSTO updates, profile
'          resets, or roaming profile sync issues.
'
' Architecture / How it works:
'  - On every successful Pro license save to My.Settings (via SaveProLicenseToSettings),
'    the same credentials are written as an XOR-encoded JSON blob to the registry
'    under RegPath_Base & RegPath_License (default value).
'  - On startup, if My.Settings has no stored Pro license, the registry backup
'    is read and restored silently before prompting the user.
'  - On license clear/deactivation, the registry backup is also cleared.
'  - All operations are fully silent (no UI) and fail-safe (exceptions are swallowed
'    with logging only).
'
' Notes:
'  - Private licenses are NOT backed up (by design — they are trivial to recreate).
'  - The existing WriteToRegistry helper is NOT used because it shows a MessageBox
'    on success. Direct Microsoft.Win32.Registry access is used instead.
'  - Encoding uses the existing CodeString/DecodeString (XOR + Base64) with SK as key.
' =============================================================================

Option Strict On
Option Explicit On

Imports Microsoft.Win32
Imports Newtonsoft.Json.Linq

Namespace SharedLibrary

    Partial Public Class SharedMethods

#Region "Registry License Backup"

        ''' <summary>
        ''' Full registry path for the license backup value.
        ''' Resolves to: HKEY_CURRENT_USER\Software\Red Ink\License
        ''' </summary>
        Private Shared ReadOnly RegistryLicenseFullPath As String = RegPath_Base & RegPath_License

        ''' <summary>
        ''' Saves a Pro license backup to the registry as an encoded JSON blob.
        ''' Called from <see cref="SaveProLicenseToSettings"/> after writing to My.Settings.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub BackupProLicenseToRegistry(productId As String,
                                                      licenseKey As String,
                                                      userId As String,
                                                      productName As String,
                                                      apiConfirmed As Boolean)
            Try
                Dim json As New JObject From {
                    {"T", "Pro"},
                    {"P", If(productId, "")},
                    {"K", If(licenseKey, "")},
                    {"U", If(userId, "")},
                    {"N", If(productName, "")},
                    {"A", apiConfirmed},
                    {"D", Date.UtcNow.ToString("o")}
                }

                Dim encoded As String = CodeString(json.ToString(Newtonsoft.Json.Formatting.None), SK)
                WriteLicenseRegistryValue(encoded)

                LogLicenseEvent("REGISTRY_BACKUP", "Pro license backup saved to registry.")

            Catch ex As Exception
                LogLicenseEvent("REGISTRY_BACKUP_ERROR",
                                $"Failed to save Pro license backup to registry: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Attempts to restore Pro license credentials from the registry backup.
        ''' Called from <see cref="LicenseOK"/> when no stored Pro license is found in My.Settings.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        ''' <returns><see langword="True"/> if a backup was found and successfully restored to My.Settings;
        ''' otherwise <see langword="False"/>.</returns>
        Friend Shared Function TryRestoreProLicenseFromRegistry() As Boolean
            Try
                Dim encoded As String = ReadLicenseRegistryValue()
                If String.IsNullOrWhiteSpace(encoded) Then Return False

                Dim decoded As String = DecodeString(encoded, SK)
                If String.IsNullOrWhiteSpace(decoded) OrElse
                   decoded.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim json As JObject = JObject.Parse(decoded)

                ' Must be a Pro license backup
                Dim licType As String = json.Value(Of String)("T")
                If Not "Pro".Equals(licType, StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim productId As String = json.Value(Of String)("P")
                Dim licenseKey As String = json.Value(Of String)("K")
                Dim userId As String = json.Value(Of String)("U")
                Dim productName As String = json.Value(Of String)("N")
                Dim apiConfirmed As Boolean = json.Value(Of Boolean)("A")

                ' Validate minimum required fields
                If String.IsNullOrWhiteSpace(productId) OrElse
                   String.IsNullOrWhiteSpace(licenseKey) OrElse
                   String.IsNullOrWhiteSpace(userId) Then
                    LogLicenseEvent("REGISTRY_RESTORE", "Registry backup found but incomplete. Skipping.")
                    Return False
                End If

                ' Restore to My.Settings via the existing save method
                SaveProLicenseToSettings(productId, licenseKey, userId,
                                         If(productName, ""), apiConfirmed)

                LogLicenseEvent("REGISTRY_RESTORE",
                                $"Pro license restored from registry backup (Product: {If(productName, "unknown")}).",
                                alwaysLog:=True)
                Return True

            Catch ex As Exception
                ' Corrupted, tampered, or missing registry data — silently ignore
                LogLicenseEvent("REGISTRY_RESTORE_ERROR",
                                $"Failed to restore license from registry: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Clears the license backup from the registry.
        ''' Called from <see cref="ClearStoredLicense"/>.
        ''' Fully silent — never shows UI, never throws.
        ''' </summary>
        Friend Shared Sub ClearLicenseRegistryBackup()
            Try
                WriteLicenseRegistryValue("")
                LogLicenseEvent("REGISTRY_BACKUP", "License registry backup cleared.")
            Catch ex As Exception
                LogLicenseEvent("REGISTRY_BACKUP_ERROR",
                                $"Failed to clear license registry backup: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "Registry Low-Level Helpers"

        ''' <summary>
        ''' Writes a string to the registry license backup location (default value).
        ''' Uses direct registry access to avoid the MessageBox in <see cref="WriteToRegistry"/>.
        ''' </summary>
        Private Shared Sub WriteLicenseRegistryValue(value As String)
            ' RegPath_Base includes "HKEY_CURRENT_USER\" prefix — strip the hive name to get the subkey path
            Dim fullPath As String = RegistryLicenseFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.CreateSubKey(subKeyPath, True)
                If subKey IsNot Nothing Then
                    subKey.SetValue("", If(value, ""), RegistryValueKind.String)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' Reads the string from the registry license backup location (default value).
        ''' Returns an empty string if the key or value does not exist.
        ''' </summary>
        Private Shared Function ReadLicenseRegistryValue() As String
            Dim fullPath As String = RegistryLicenseFullPath
            Dim hiveName As String = fullPath.Split("\"c)(0)
            Dim subKeyPath As String = fullPath.Substring(hiveName.Length + 1)

            Using subKey As RegistryKey = Registry.CurrentUser.OpenSubKey(subKeyPath)
                If subKey Is Nothing Then Return ""
                Dim val As Object = subKey.GetValue("", Nothing)
                Return If(val?.ToString(), "")
            End Using
        End Function

#End Region

    End Class

End Namespace