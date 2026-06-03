' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Cleanup.vb
' Purpose:
'   Inky AutoPilot retention cleanup for Outlook.
'   Automatically deletes AutoPilot-processed mail threads after a configured
'   number of hours, including items already moved to Deleted Items.
'
' Architecture:
'  - Uses hidden MAPI properties to group an incoming mail with all AutoPilot
'    replies and follow-up notices.
'  - Starts the retention clock only after a substantive reply is sent.
'  - Runs a periodic cleanup timer alongside the AutoPilot session.
'  - Scans mailbox folders and Deleted Items to remove expired tagged items.
'  - Keeps Outlook COM access on the UI thread via `SwitchToUi`.
'
' Retention Model:
'  - `AutoDeleteAfterHours = 0` disables cleanup.
'  - When enabled, the original mail and all AutoPilot replies in the same group
'    are eligible for deletion after the configured retention window.
'  - Cleanup includes normal folders and Deleted Items.
'
' Security & Reliability:
'  - Deletion is based on hidden MAPI metadata, not EntryID or subject text.
'  - The cleanup process is best-effort and fail-safe.
'  - COM objects are released after use to avoid Outlook resource leaks.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop.Outlook

Partial Public Class ThisAddIn

    Private Const AP_CleanupGroupIdProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteGroupId"
    Private Const AP_CleanupEligibleProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteEligible"
    Private Const AP_CleanupAnsweredUtcProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteAnsweredUtc"
    Private Const AP_CleanupDeleteAfterUtcProperty As String =
        "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-RedInk-AutoDeleteAfterUtc"

    Private Const AP_AutoDeleteTimerIntervalSeconds As Integer = 15 * 60

    Private _apAutoDeleteTimer As System.Threading.Timer = Nothing
    Private _apAutoDeleteCheckRunning As Integer = 0

    Private Structure AutoDeleteCleanupStats
        Public ScannedCount As Integer
        Public DeletedCount As Integer
        Public ErrorCount As Integer
    End Structure

    Friend Sub StartAutoDeleteTimer()
        StopAutoDeleteTimer()

        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return

        _apAutoDeleteTimer = New System.Threading.Timer(
            AddressOf AutoDeleteTimerCallback,
            Nothing,
            dueTime:=TimeSpan.FromSeconds(30),
            period:=TimeSpan.FromSeconds(AP_AutoDeleteTimerIntervalSeconds))

        ApDashboardLog($"🗑 Auto-delete timer started ({_apConfig.AutoDeleteAfterHours}h retention, including Deleted Items).", "info")
    End Sub

    Friend Sub StopAutoDeleteTimer()
        Try : _apAutoDeleteTimer?.Dispose() : Catch : End Try
        _apAutoDeleteTimer = Nothing
    End Sub

    Private Async Sub AutoDeleteTimerCallback(state As Object)
        If Not _apActive Then Return
        If Interlocked.CompareExchange(_apAutoDeleteCheckRunning, 1, 0) <> 0 Then Return

        Try
            Dim ct = _apCts?.Token
            If ct Is Nothing OrElse ct.Value.IsCancellationRequested Then Return
            Await RunAutoDeleteCleanupAsync(ct.Value)
        Catch ex As OperationCanceledException
            ' Expected during shutdown
        Catch ex As System.Exception
            ApDashboardLog($"🗑 Auto-delete timer error: {ex.Message}", "warn")
        Finally
            Interlocked.Exchange(_apAutoDeleteCheckRunning, 0)
        End Try
    End Sub

    Friend Async Function RunAutoDeleteCleanupAsync(ct As CancellationToken) As Task
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return

        Dim pass1 As AutoDeleteCleanupStats
        Dim pass2 As AutoDeleteCleanupStats

        Await SwitchToUi(
            Sub()
                pass1 = DeleteExpiredAutoPilotItemsOutsideDeletedItems()
            End Sub)

        Await SwitchToUi(
            Sub()
                pass2 = DeleteExpiredAutoPilotItemsInsideDeletedItems()
            End Sub)

        Dim totalDeleted = pass1.DeletedCount + pass2.DeletedCount
        Dim totalErrors = pass1.ErrorCount + pass2.ErrorCount

        If totalDeleted > 0 OrElse totalErrors > 0 Then
            ApDashboardLog(
                $"🗑 Auto-delete cleanup: {totalDeleted} item(s) removed, {totalErrors} error(s), {pass1.ScannedCount + pass2.ScannedCount} item(s) scanned.",
                If(totalErrors > 0, "warn", "info"))
        End If
    End Function

    Private Function GetAutoDeleteCutoffUtc() As DateTime?
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return Nothing
        Return DateTime.UtcNow.AddHours(_apConfig.AutoDeleteAfterHours)
    End Function

    Private Function GetCleanupGroupId(mi As MailItem) As String
        If mi Is Nothing Then Return Nothing

        Try
            Dim value = CStr(mi.PropertyAccessor.GetProperty(AP_CleanupGroupIdProperty))
            If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
        Catch
        End Try

        Return Nothing
    End Function

    Friend Function GetOrCreateCleanupGroupId(mi As MailItem) As String
        If mi Is Nothing Then Return Nothing
        If _apConfig Is Nothing OrElse _apConfig.AutoDeleteAfterHours <= 0 Then Return Nothing

        Dim existing = GetCleanupGroupId(mi)
        If Not String.IsNullOrWhiteSpace(existing) Then Return existing

        Dim groupId = Guid.NewGuid().ToString("N")

        Try
            mi.PropertyAccessor.SetProperty(AP_CleanupGroupIdProperty, groupId)
            mi.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] Failed to create cleanup group ID: {ex.Message}")
        End Try

        Return groupId
    End Function

    Friend Sub StampCleanupMetadata(mi As MailItem,
                                    groupId As String,
                                    isEligible As Boolean,
                                    answeredUtc As DateTime?,
                                    deleteAfterUtc As DateTime?,
                                    Optional saveItem As Boolean = True)

        If mi Is Nothing OrElse String.IsNullOrWhiteSpace(groupId) Then Return

        Try
            Dim pa = mi.PropertyAccessor
            pa.SetProperty(AP_CleanupGroupIdProperty, groupId)
            pa.SetProperty(AP_CleanupEligibleProperty, If(isEligible, "true", "false"))
            pa.SetProperty(
                AP_CleanupAnsweredUtcProperty,
                If(answeredUtc.HasValue,
                   answeredUtc.Value.ToString("o", CultureInfo.InvariantCulture),
                   ""))
            pa.SetProperty(
                AP_CleanupDeleteAfterUtcProperty,
                If(deleteAfterUtc.HasValue,
                   deleteAfterUtc.Value.ToString("o", CultureInfo.InvariantCulture),
                   ""))

            If saveItem Then mi.Save()
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] StampCleanupMetadata error: {ex.Message}")
        End Try
    End Sub

    Friend Sub MarkMailGroupAsAnsweredAndEligible(originalMail As MailItem)
        If originalMail Is Nothing Then Return

        Dim deleteAfterUtc = GetAutoDeleteCutoffUtc()
        If Not deleteAfterUtc.HasValue Then Return

        Dim groupId = GetOrCreateCleanupGroupId(originalMail)
        If String.IsNullOrWhiteSpace(groupId) Then Return

        Dim answeredUtc = DateTime.UtcNow

        StampCleanupMetadata(
            originalMail,
            groupId,
            isEligible:=True,
            answeredUtc:=answeredUtc,
            deleteAfterUtc:=deleteAfterUtc,
            saveItem:=True)

        Dim stampedCount As Integer = 0
        ApplyEligibilityToGroupInAllStores(groupId, answeredUtc, deleteAfterUtc.Value, stampedCount)

        ApDashboardLog(
            $"🗑 Auto-delete scheduled for group {groupId} in {_apConfig.AutoDeleteAfterHours}h ({stampedCount} item(s) marked).",
            "step")
    End Sub

    Private Sub ApplyEligibilityToGroupInAllStores(groupId As String,
                                                   answeredUtc As DateTime,
                                                   deleteAfterUtc As DateTime,
                                                   ByRef stampedCount As Integer)

        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            For i As Integer = 1 To session.Stores.Count
                Dim store As Store = Nothing
                Dim root As MAPIFolder = Nothing

                Try
                    store = session.Stores(i)
                    root = store.GetRootFolder()
                    ApplyEligibilityToGroupInFolderTree(root, groupId, answeredUtc, deleteAfterUtc, stampedCount)
                Catch
                Finally
                    If root IsNot Nothing Then Try : Marshal.ReleaseComObject(root) : Catch : End Try
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] ApplyEligibilityToGroupInAllStores error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try
    End Sub

    Private Sub ApplyEligibilityToGroupInFolderTree(folder As MAPIFolder,
                                                    groupId As String,
                                                    answeredUtc As DateTime,
                                                    deleteAfterUtc As DateTime,
                                                    ByRef stampedCount As Integer)

        If folder Is Nothing Then Return

        Dim items As Items = Nothing
        Dim subFolders As Folders = Nothing

        Try
            items = folder.Items
            For i As Integer = items.Count To 1 Step -1
                Dim obj As Object = Nothing
                Dim mi As MailItem = Nothing

                Try
                    obj = items(i)
                    mi = TryCast(obj, MailItem)
                    If mi Is Nothing Then Continue For

                    Dim itemGroupId = GetCleanupGroupId(mi)
                    If itemGroupId Is Nothing OrElse
                       Not itemGroupId.Equals(groupId, StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    StampCleanupMetadata(
                        mi,
                        groupId,
                        isEligible:=True,
                        answeredUtc:=answeredUtc,
                        deleteAfterUtc:=deleteAfterUtc,
                        saveItem:=True)

                    stampedCount += 1
                Catch
                Finally
                    If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                    If obj IsNot Nothing AndAlso Not ReferenceEquals(obj, mi) Then
                        Try : Marshal.ReleaseComObject(obj) : Catch : End Try
                    End If
                End Try
            Next

            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    ApplyEligibilityToGroupInFolderTree(child, groupId, answeredUtc, deleteAfterUtc, stampedCount)
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
            If items IsNot Nothing Then Try : Marshal.ReleaseComObject(items) : Catch : End Try
        End Try
    End Sub

    Private Function DeleteExpiredAutoPilotItemsOutsideDeletedItems() As AutoDeleteCleanupStats
        Dim stats As New AutoDeleteCleanupStats()
        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            For i As Integer = 1 To session.Stores.Count
                Dim store As Store = Nothing
                Dim root As MAPIFolder = Nothing
                Dim deletedItems As MAPIFolder = Nothing

                Try
                    store = session.Stores(i)
                    root = store.GetRootFolder()
                    Try
                        deletedItems = store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                    Catch
                        deletedItems = Nothing
                    End Try

                    DeleteExpiredItemsOutsideDeletedItemsTree(
                        root,
                        If(deletedItems IsNot Nothing, deletedItems.EntryID, ""),
                        DateTime.UtcNow,
                        stats)
                Catch
                Finally
                    If deletedItems IsNot Nothing Then Try : Marshal.ReleaseComObject(deletedItems) : Catch : End Try
                    If root IsNot Nothing Then Try : Marshal.ReleaseComObject(root) : Catch : End Try
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] DeleteExpiredAutoPilotItemsOutsideDeletedItems error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try

        Return stats
    End Function

    Private Function DeleteExpiredAutoPilotItemsInsideDeletedItems() As AutoDeleteCleanupStats
        Dim stats As New AutoDeleteCleanupStats()
        Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Nothing

        Try
            session = Application.GetNamespace("MAPI")

            For i As Integer = 1 To session.Stores.Count
                Dim store As Store = Nothing
                Dim deletedItems As MAPIFolder = Nothing

                Try
                    store = session.Stores(i)
                    deletedItems = store.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems)
                    DeleteExpiredItemsInFolderTree(deletedItems, DateTime.UtcNow, stats)
                Catch
                Finally
                    If deletedItems IsNot Nothing Then Try : Marshal.ReleaseComObject(deletedItems) : Catch : End Try
                    If store IsNot Nothing Then Try : Marshal.ReleaseComObject(store) : Catch : End Try
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[AutoPilot] DeleteExpiredAutoPilotItemsInsideDeletedItems error: {ex.Message}")
        Finally
            If session IsNot Nothing Then Try : Marshal.ReleaseComObject(session) : Catch : End Try
        End Try

        Return stats
    End Function

    Private Sub DeleteExpiredItemsOutsideDeletedItemsTree(folder As MAPIFolder,
                                                          deletedItemsEntryId As String,
                                                          nowUtc As DateTime,
                                                          ByRef stats As AutoDeleteCleanupStats)

        If folder Is Nothing Then Return

        Try
            If Not String.IsNullOrWhiteSpace(deletedItemsEntryId) AndAlso
               folder.EntryID.Equals(deletedItemsEntryId, StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
        Catch
        End Try

        DeleteExpiredItemsInCurrentFolder(folder, nowUtc, stats)

        Dim subFolders As Folders = Nothing
        Try
            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    DeleteExpiredItemsOutsideDeletedItemsTree(child, deletedItemsEntryId, nowUtc, stats)
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
        End Try
    End Sub

    Private Sub DeleteExpiredItemsInFolderTree(folder As MAPIFolder,
                                               nowUtc As DateTime,
                                               ByRef stats As AutoDeleteCleanupStats)

        If folder Is Nothing Then Return

        DeleteExpiredItemsInCurrentFolder(folder, nowUtc, stats)

        Dim subFolders As Folders = Nothing
        Try
            subFolders = folder.Folders
            For i As Integer = 1 To subFolders.Count
                Dim child As MAPIFolder = Nothing
                Try
                    child = subFolders(i)
                    DeleteExpiredItemsInFolderTree(child, nowUtc, stats)
                Catch
                Finally
                    If child IsNot Nothing Then Try : Marshal.ReleaseComObject(child) : Catch : End Try
                End Try
            Next
        Catch
        Finally
            If subFolders IsNot Nothing Then Try : Marshal.ReleaseComObject(subFolders) : Catch : End Try
        End Try
    End Sub

    Private Sub DeleteExpiredItemsInCurrentFolder(folder As MAPIFolder,
                                                  nowUtc As DateTime,
                                                  ByRef stats As AutoDeleteCleanupStats)

        Dim items As Items = Nothing

        Try
            items = folder.Items

            For i As Integer = items.Count To 1 Step -1
                Dim obj As Object = Nothing
                Dim mi As MailItem = Nothing

                Try
                    obj = items(i)
                    mi = TryCast(obj, MailItem)
                    If mi Is Nothing Then Continue For

                    stats.ScannedCount += 1

                    If Not ShouldAutoDeleteMail(mi, nowUtc) Then Continue For

                    mi.Delete()
                    stats.DeletedCount += 1
                Catch
                    stats.ErrorCount += 1
                Finally
                    If mi IsNot Nothing Then Try : Marshal.ReleaseComObject(mi) : Catch : End Try
                    If obj IsNot Nothing AndAlso Not ReferenceEquals(obj, mi) Then
                        Try : Marshal.ReleaseComObject(obj) : Catch : End Try
                    End If
                End Try
            Next
        Catch
            stats.ErrorCount += 1
        Finally
            If items IsNot Nothing Then Try : Marshal.ReleaseComObject(items) : Catch : End Try
        End Try
    End Sub

    Private Function ShouldAutoDeleteMail(mi As MailItem, nowUtc As DateTime) As Boolean
        If mi Is Nothing Then Return False

        Dim groupId = GetCleanupGroupId(mi)
        If String.IsNullOrWhiteSpace(groupId) Then Return False

        Dim eligibleRaw As String = Nothing
        If Not TryGetNamedStringProperty(mi, AP_CleanupEligibleProperty, eligibleRaw) Then Return False
        If Not "true".Equals(eligibleRaw, StringComparison.OrdinalIgnoreCase) Then Return False

        Dim deleteAfterUtc As DateTime
        If Not TryGetNamedDateUtcProperty(mi, AP_CleanupDeleteAfterUtcProperty, deleteAfterUtc) Then Return False

        Return deleteAfterUtc <= nowUtc
    End Function

    Private Function TryGetNamedStringProperty(mi As MailItem, propertyName As String, ByRef value As String) As Boolean
        value = Nothing
        If mi Is Nothing Then Return False

        Try
            Dim raw = mi.PropertyAccessor.GetProperty(propertyName)
            If raw Is Nothing Then Return False
            value = CStr(raw)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function TryGetNamedDateUtcProperty(mi As MailItem, propertyName As String, ByRef value As DateTime) As Boolean
        value = DateTime.MinValue

        Dim raw As String = Nothing
        If Not TryGetNamedStringProperty(mi, propertyName, raw) Then Return False
        If String.IsNullOrWhiteSpace(raw) Then Return False

        Dim parsed As DateTime
        If DateTime.TryParse(
            raw,
            CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind Or DateTimeStyles.AssumeUniversal,
            parsed) Then

            value = parsed.ToUniversalTime()
            Return True
        End If

        Return False
    End Function

End Class