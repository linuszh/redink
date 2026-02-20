' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Commands.InboxBoard.vb
' Purpose: Visual Kanban-style board for categorized Inbox emails. Displays
'          flagged (categorized) mails as draggable tiles grouped by category.
'          Supports drag-and-drop to change categories, unmarking, opening mails
'          in Outlook, and async AI summary generation.
'
' Architecture:
'  - Mail collection: Scans default Inbox for mails with at least one category.
'    Optionally also includes flagged mails (follow-up flags). If >50 found,
'    asks user how many to load via dropdown.
'  - Flag support: Mails flagged for follow-up (olFlagMarked) can be included
'    via a settings toggle. They appear in a synthetic "⚑ Flagged" column
'    (unless they also have categories, in which case they appear in their
'    category columns with a flag badge). The FlagRequest text (e.g. "Follow up")
'    and optional due date are shown on the card. Completed flags (olFlagComplete)
'    are hidden by default via a "Hide completed flags" toggle.
'  - Conversation grouping: Mails sharing the same ConversationTopic are merged
'    into a single card showing the latest mail, with an accurate message count.
'    Toggle via a checkbox in the settings panel; persisted in My.Settings.
'  - Category columns: Dynamically built from Outlook's category list. Always-
'    present columns stored in My.Settings.InboxBoardColumns (semicolon-delimited).
'  - Board UI: Full HTML/CSS/JS board rendered in WebView2 inside a WinForms Form.
'    Supports dark/light theme, search, column filter, card field toggles, and
'    drag-and-drop reordering.
'  - JS↔VB.NET bridge: WebView2 postMessage for move/moveAdd/unmark/open/reload
'    actions. VB.NET responds by modifying MailItem categories via COM and pushing
'    updated data back to the board.
'  - Drag modes: Normal drag replaces all categories with the target column's
'    category. Ctrl+drag adds the target category while keeping existing ones
'    (like copy vs. move in file managers). Visual feedback: green drop indicator
'    and "+" badge when Ctrl is held.
'  - AI Summaries: Generated asynchronously in batches after the board is displayed.
'    Uses the existing LLM() helper with a summarization system prompt.
'    Summaries are cached in My.Settings (JSON dict, capped at 500 entries).
'    Summary language is selectable via a dropdown, persisted in My.Settings.
'  - Threading: All Outlook COM access on the UI thread; LLM calls are async.
'  - Persistence: Theme, window position/size/maximized state, card field toggles,
'    last load count, column filter, pinned columns, summary cache, summary
'    language, conversation grouping toggle, include-flagged toggle, and
'    hide-done-flags toggle are all persisted via My.Settings (not localStorage).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "InboxBoard Constants"

    ''' <summary>Number of mails per AI summary batch.</summary>
    Private Const InboxBoard_SummaryBatchSize As Integer = 5

    ''' <summary>Maximum body excerpt length sent for summarization.</summary>
    Private Const InboxBoard_SummaryBodyCap As Integer = 2000

    ''' <summary>Maximum cached summaries to persist in My.Settings.</summary>
    Private Const InboxBoard_SummaryCacheMax As Integer = 500

    ''' <summary>Column ID used for the synthetic "Flagged" column.</summary>
    Private Const InboxBoard_FlaggedColumnId As String = "⚑ Flagged"

    ''' <summary>Color used for the synthetic "Flagged" column.</summary>
    Private Const InboxBoard_FlaggedColumnColor As String = "#ef4444"

    ''' <summary>Remembers the user's last load-count choice so Refresh reuses it.</summary>
    Private _inboxBoardLastMaxToLoad As Integer = 0

    ''' <summary>In-memory summary cache: EntryID → summary text.</summary>
    Private _inboxBoardSummaryCache As Dictionary(Of String, String) = Nothing

    ''' <summary>Available languages for summary generation.</summary>
    Private Shared ReadOnly InboxBoard_SummaryLanguages As String()() = {
        New String() {"auto", "Auto-detect"},
        New String() {"en", "English"},
        New String() {"de", "Deutsch"},
        New String() {"fr", "Français"},
        New String() {"it", "Italiano"},
        New String() {"es", "Español"},
        New String() {"pt", "Português"},
        New String() {"nl", "Nederlands"},
        New String() {"pl", "Polski"},
        New String() {"ru", "Русский"},
        New String() {"ja", "日本語"},
        New String() {"zh", "中文"},
        New String() {"ko", "한국어"},
        New String() {"ar", "العربية"}
    }

    ''' <summary>Column IDs for date-bucketed flagged columns.</summary>
    Private Const InboxBoard_FlagBucket_Overdue As String = "⚑ Overdue"
    Private Const InboxBoard_FlagBucket_Today As String = "⚑ Today"
    Private Const InboxBoard_FlagBucket_Tomorrow As String = "⚑ Tomorrow"
    Private Const InboxBoard_FlagBucket_ThisWeek As String = "⚑ This Week"
    Private Const InboxBoard_FlagBucket_NextWeek As String = "⚑ Next Week"
    Private Const InboxBoard_FlagBucket_Next4Weeks As String = "⚑ Next 4 Weeks"
    Private Const InboxBoard_FlagBucket_Later As String = "⚑ Later"
    Private Const InboxBoard_FlagBucket_NoDueDate As String = "⚑ Flagged"
    Private Const InboxBoard_FlagBucket_Done As String = "✓ Done"

    ''' <summary>Colors for each flag date bucket.</summary>
    Private Shared ReadOnly InboxBoard_FlagBucketColors As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.Ordinal) From {
        {InboxBoard_FlagBucket_Overdue, "#dc2626"},
        {InboxBoard_FlagBucket_Today, "#ef4444"},
        {InboxBoard_FlagBucket_Tomorrow, "#f97316"},
        {InboxBoard_FlagBucket_ThisWeek, "#eab308"},
        {InboxBoard_FlagBucket_NextWeek, "#22c55e"},
        {InboxBoard_FlagBucket_Next4Weeks, "#3b82f6"},
        {InboxBoard_FlagBucket_Later, "#8b5cf6"},
        {InboxBoard_FlagBucket_NoDueDate, "#6b7280"},
        {InboxBoard_FlagBucket_Done, "#9ca3af"}
    }

    ''' <summary>Ordered list of all flag bucket column IDs (display order).</summary>
    Private Shared ReadOnly InboxBoard_FlagBucketOrder As String() = {
        InboxBoard_FlagBucket_Overdue,
        InboxBoard_FlagBucket_Today,
        InboxBoard_FlagBucket_Tomorrow,
        InboxBoard_FlagBucket_ThisWeek,
        InboxBoard_FlagBucket_NextWeek,
        InboxBoard_FlagBucket_Next4Weeks,
        InboxBoard_FlagBucket_Later,
        InboxBoard_FlagBucket_NoDueDate,
        InboxBoard_FlagBucket_Done
    }

#End Region

#Region "InboxBoard Data Classes"

    ''' <summary>Represents a categorized mail displayed on the board.</summary>
    Private Class InboxBoardEntry
        Public Property EntryID As String
        Public Property Subject As String
        Public Property SenderName As String
        Public Property SenderEmail As String
        Public Property Recipients As String
        Public Property ReceivedTime As DateTime
        Public Property IsRead As Boolean
        Public Property MessageCount As Integer
        Public Property Categories As String          ' First category name
        Public Property AllCategories As String       ' Semicolon-separated
        Public Property BodyExcerpt As String
        Public Property Summary As String             ' AI-generated, initially empty
        Public Property CategoryColor As String       ' Hex color for the category
        Public Property ConversationTopic As String   ' For conversation grouping
        Public Property ConversationEntryIDs As List(Of String) ' All EntryIDs in this conversation group
        Public Property IsGrouped As Boolean          ' True when this card represents a grouped conversation
        Public Property IsFlagged As Boolean          ' True when mail has FlagStatus = olFlagMarked
        Public Property IsCompleted As Boolean        ' True when mail has FlagStatus = olFlagComplete
        Public Property FlagRequest As String         ' Flag label text (e.g. "Follow up", "Reply")
        Public Property FlagDueDate As DateTime?      ' Optional due date from TaskDueDate
        Public Property FlagBucket As String          ' Which date-bucket column this card belongs to (if flagged)
    End Class

    ''' <summary>Represents a column (category) on the board.</summary>
    Private Class InboxBoardColumn
        Public Property CategoryName As String
        Public Property Color As String               ' Hex color
        Public Property Tag As String                 ' Display tag
        Public Property IsFlaggedColumn As Boolean    ' True for the synthetic "Flagged" column
    End Class

#End Region

#Region "InboxBoard Summary Cache"

    ''' <summary>
    ''' Loads the summary cache from My.Settings into the in-memory dictionary.
    ''' </summary>
    Private Sub LoadSummaryCache()
        If _inboxBoardSummaryCache IsNot Nothing Then Return
        _inboxBoardSummaryCache = New Dictionary(Of String, String)(StringComparer.Ordinal)
        Try
            Dim json As String = My.Settings.InboxBoardSummaryCache
            If Not String.IsNullOrEmpty(json) Then
                Dim obj As JObject = JObject.Parse(json)
                For Each prop In obj.Properties()
                    _inboxBoardSummaryCache(prop.Name) = CStr(prop.Value)
                Next
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Saves the in-memory summary cache to My.Settings, capping at InboxBoard_SummaryCacheMax entries.
    ''' </summary>
    Private Sub SaveSummaryCache()
        Try
            If _inboxBoardSummaryCache Is Nothing Then Return

            ' If over limit, remove oldest entries (by insertion order — Dictionary preserves order in .NET 4.x)
            While _inboxBoardSummaryCache.Count > InboxBoard_SummaryCacheMax
                Dim firstKey As String = _inboxBoardSummaryCache.Keys.First()
                _inboxBoardSummaryCache.Remove(firstKey)
            End While

            Dim obj As New JObject()
            For Each kvp In _inboxBoardSummaryCache
                obj(kvp.Key) = kvp.Value
            Next
            My.Settings.InboxBoardSummaryCache = obj.ToString(Formatting.None)
            My.Settings.Save()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Looks up a cached summary for a given EntryID.
    ''' </summary>
    Private Function GetCachedSummary(entryId As String) As String
        LoadSummaryCache()
        Dim result As String = Nothing
        If _inboxBoardSummaryCache.TryGetValue(entryId, result) Then Return result
        Return Nothing
    End Function

    ''' <summary>
    ''' Stores a summary in the cache (in-memory + persistent).
    ''' </summary>
    Private Sub CacheSummary(entryId As String, summary As String)
        LoadSummaryCache()
        _inboxBoardSummaryCache(entryId) = summary
    End Sub

    ''' <summary>
    ''' Clears the entire summary cache (in-memory + persistent).
    ''' </summary>
    Private Sub ClearSummaryCache()
        _inboxBoardSummaryCache = New Dictionary(Of String, String)(StringComparer.Ordinal)
        Try
            My.Settings.InboxBoardSummaryCache = ""
            My.Settings.Save()
        Catch
        End Try
    End Sub

#End Region


#Region "InboxBoard Entry Point"

    ''' <summary>
    ''' Main entry point for the Inbox Board feature.
    ''' </summary>
    Public Sub InboxBoard()
        Try
            ' Restore last load count from settings
            Try : _inboxBoardLastMaxToLoad = My.Settings.InboxBoardLastLoadCount : Catch : End Try

            ' Ensure summary cache is loaded
            LoadSummaryCache()

            ' 1. Collect categorized mails from Inbox
            Dim mails As List(Of InboxBoardEntry) = CollectCategorizedInboxMails()
            If mails Is Nothing Then Return

            ' 2. Build column definitions from Outlook categories
            Dim columns As List(Of InboxBoardColumn) = BuildBoardColumns(mails)
            If columns Is Nothing Then columns = New List(Of InboxBoardColumn)()

            ' 3. Show the board
            ShowInboxBoardForm(mails, columns)

        Catch ex As System.Exception
            ShowCustomMessageBox($"Inbox Board error: {ex.Message}", $"{AN} - Inbox Board")
        End Try
    End Sub

#End Region

#Region "InboxBoard Mail Collection"

    ''' <summary>
    ''' Scans the default Inbox for mails that have at least one category assigned,
    ''' and optionally mails that are flagged for follow-up.
    ''' If 1000 or more qualifying mails are found, asks the user how many to load.
    ''' </summary>
    Private Function CollectCategorizedInboxMails(Optional maxOverride As Integer = 0) As List(Of InboxBoardEntry)
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim inbox As MAPIFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox))

        If inbox Is Nothing Then
            ShowCustomMessageBox("Could not access the Inbox folder.", $"{AN} - Inbox Board")
            Return Nothing
        End If

        ' Read flag-related settings
        Dim includeFlagged As Boolean = False
        Dim hideDoneFlags As Boolean = True
        Try : includeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
        Try : hideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try

        Dim folderItems As Outlook.Items = ComRetry(Of Outlook.Items)(Function() inbox.Items)
        Dim totalItems As Integer = ComRetry(Of Integer)(Function() folderItems.Count)

        ' Count qualifying mails (manual check to avoid Restrict miscounting)
        ' Note: Flagged mails are ALWAYS counted here regardless of the includeFlagged
        ' toggle, so the board opens when flagged mails exist. The toggle only controls
        ' whether flagged mails are *displayed* (via BuildInboxBoardEntry/BuildBoardColumns).
        Dim totalQualifying As Integer = 0
        For i As Integer = 1 To totalItems
            Try
                Dim idx As Integer = i
                Dim item As Object = ComRetry(Function() folderItems.Item(idx))
                Dim mi As MailItem = TryCast(item, MailItem)
                If mi Is Nothing Then Continue For

                Dim cats As String = ""
                Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
                Dim hasCats As Boolean = Not String.IsNullOrWhiteSpace(cats)

                Dim qualifies As Boolean = hasCats
                If Not qualifies Then
                    ' Always count flagged mails so the board can open.
                    ' The includeFlagged toggle only affects display, not whether
                    ' the board launches.
                    Dim flagStatus As Integer = 0
                    Try : flagStatus = ComRetry(Of Integer)(Function() CInt(mi.FlagStatus)) : Catch : End Try
                    ' olFlagMarked = 2, olFlagComplete = 1
                    If flagStatus = 2 Then
                        qualifies = True
                    ElseIf flagStatus = 1 AndAlso Not hideDoneFlags Then
                        qualifies = True
                    End If
                End If

                If qualifies Then totalQualifying += 1
            Catch
            End Try
        Next

        If totalQualifying = 0 Then
            ' Also allow the board to open if there are pinned columns
            ' (even with zero qualifying mails, empty columns are useful as drop targets)
            Dim hasPinnedColumns As Boolean = False
            Try
                Dim pinned As String = If(My.Settings.InboxBoardColumns, "")
                If Not String.IsNullOrEmpty(pinned) Then hasPinnedColumns = True
            Catch
            End Try

            Dim hasPinnedFlagColumns As Boolean = False
            Try
                Dim pinnedFlags As String = If(My.Settings.InboxBoardPinnedFlagColumns, "")
                If Not String.IsNullOrEmpty(pinnedFlags) Then hasPinnedFlagColumns = True
            Catch
            End Try

            If Not hasPinnedColumns AndAlso Not hasPinnedFlagColumns Then
                ShowCustomMessageBox("No categorized or flagged mails found in the Inbox.", $"{AN} - Inbox Board")
                Return Nothing
            End If

            ' Return an empty list — the board will show pinned columns with no cards
            Return New List(Of InboxBoardEntry)()
        End If

        Dim maxToLoad As Integer = totalQualifying

        ' Ask how many if > threshold (and not a reload with a remembered value)
        If totalQualifying > 1000 Then
            ' On reload, reuse the user's last choice if available —
            ' but only if it was a meaningful choice (>= 1000). A small value
            ' from a previous session with few qualifying mails must not cap
            ' a larger session.
            If _inboxBoardLastMaxToLoad >= 1000 Then
                maxToLoad = _inboxBoardLastMaxToLoad
            Else
                Dim items As New List(Of SelectionItem)()
            items.Add(New SelectionItem("1000 mails", 1000))
            items.Add(New SelectionItem("2500 mails", 2500))
            items.Add(New SelectionItem("5000 mails", 5000))
            items.Add(New SelectionItem("7500 mails", 7500))
            If totalQualifying > 7500 Then
                items.Add(New SelectionItem($"All ({totalQualifying} mails)", totalQualifying))
            End If

                Dim chosen As Integer = SelectValue(items, 2500,
                    $"Found {totalQualifying} qualifying mails in Inbox. How many should be loaded?",
                    $"{AN} - Inbox Board")
                If chosen = 0 Then Return Nothing
                maxToLoad = chosen
            End If
        End If

        ' Remember the effective load count for Refresh and persist
        _inboxBoardLastMaxToLoad = maxToLoad
        Try
            My.Settings.InboxBoardLastLoadCount = maxToLoad
            My.Settings.Save()
        Catch
        End Try

        ' Build the category color map from Outlook
        Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)

        ' Extract entries (manual filter, then sort)
        Dim entries As New List(Of InboxBoardEntry)()
        For i As Integer = 1 To totalItems
            Try
                Dim idx As Integer = i
                Dim item As Object = ComRetry(Function() folderItems.Item(idx))
                Dim mi As MailItem = TryCast(item, MailItem)
                If mi IsNot Nothing Then
                    Dim entry As InboxBoardEntry = BuildInboxBoardEntry(mi, categoryColors, includeFlagged, hideDoneFlags)
                    If entry IsNot Nothing Then entries.Add(entry)
                End If
            Catch
            End Try
        Next

        ' Sort by received time descending
        entries.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

        ' Cap to maxToLoad
        If entries.Count > maxToLoad Then
            entries.RemoveRange(maxToLoad, entries.Count - maxToLoad)
        End If

        ' Apply conversation grouping if enabled
        Dim groupConversations As Boolean = False
        Try : groupConversations = My.Settings.InboxBoardGroupConversations : Catch : End Try
        If groupConversations Then
            entries = GroupByConversation(entries)
        End If

        ' Apply cached summaries
        For Each entry In entries
            Dim cached As String = GetCachedSummary(entry.EntryID)
            If Not String.IsNullOrEmpty(cached) Then
                entry.Summary = cached
            End If
        Next

        Return entries
    End Function

    ''' <summary>
    ''' Groups entries by ConversationTopic. For each conversation with 2+ mails,
    ''' keeps the latest mail as the representative card. Single-topic and no-topic
    ''' entries are preserved as individual ungrouped cards.
    ''' </summary>
    Private Function GroupByConversation(entries As List(Of InboxBoardEntry)) As List(Of InboxBoardEntry)
        ' Group by ConversationTopic (case-insensitive). Mails without a topic stay ungrouped.
        Dim grouped As New Dictionary(Of String, List(Of InboxBoardEntry))(StringComparer.OrdinalIgnoreCase)
        Dim ungrouped As New List(Of InboxBoardEntry)()

        For Each entry In entries
            If String.IsNullOrWhiteSpace(entry.ConversationTopic) Then
                entry.IsGrouped = False
                ungrouped.Add(entry)
            Else
                Dim key As String = entry.ConversationTopic.Trim()
                If Not grouped.ContainsKey(key) Then
                    grouped(key) = New List(Of InboxBoardEntry)()
                End If
                grouped(key).Add(entry)
            End If
        Next

        Dim result As New List(Of InboxBoardEntry)()

        For Each kvp In grouped
            Dim convMails As List(Of InboxBoardEntry) = kvp.Value

            ' If only one mail shares this topic, treat it as ungrouped
            If convMails.Count = 1 Then
                convMails(0).IsGrouped = False
                result.Add(convMails(0))
                Continue For
            End If

            ' Sort by ReceivedTime descending — first item is the latest
            convMails.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

            Dim representative As InboxBoardEntry = convMails(0)
            representative.MessageCount = convMails.Count
            representative.IsGrouped = True
            representative.ConversationEntryIDs = convMails.Select(Function(m) m.EntryID).ToList()

            ' If any mail in the conversation is unread, mark the card as unread
            If convMails.Any(Function(m) Not m.IsRead) Then
                representative.IsRead = False
            End If

            ' If any mail in the conversation is flagged, mark the card as flagged
            If convMails.Any(Function(m) m.IsFlagged) Then
                representative.IsFlagged = True
                ' Use the flag request from the first flagged mail
                Dim firstFlagged = convMails.FirstOrDefault(Function(m) m.IsFlagged)
                If firstFlagged IsNot Nothing Then
                    If String.IsNullOrEmpty(representative.FlagRequest) Then
                        representative.FlagRequest = firstFlagged.FlagRequest
                    End If
                    If Not representative.FlagDueDate.HasValue AndAlso firstFlagged.FlagDueDate.HasValue Then
                        representative.FlagDueDate = firstFlagged.FlagDueDate
                    End If
                End If
            End If

            ' Recompute flag bucket for the representative card
            representative.FlagBucket = GetFlagBucket(representative)

            ' Merge all categories from all mails in the conversation
            Dim allCats As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each m In convMails
                If Not String.IsNullOrEmpty(m.AllCategories) Then
                    For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = catPart.Trim()
                        If Not String.IsNullOrEmpty(t) Then allCats.Add(t)
                    Next
                End If
            Next
            representative.AllCategories = String.Join(", ", allCats)
            If allCats.Count > 0 Then representative.Categories = allCats.First()

            ' Check if any mail in the conversation has a cached summary
            If String.IsNullOrEmpty(representative.Summary) Then
                For Each m In convMails
                    Dim cached As String = GetCachedSummary(m.EntryID)
                    If Not String.IsNullOrEmpty(cached) Then
                        representative.Summary = cached
                        Exit For
                    End If
                Next
            End If

            result.Add(representative)
        Next

        ' Add all ungrouped mails — these were previously lost when toggling grouping off/on
        result.AddRange(ungrouped)

        ' Re-sort by received time descending
        result.Sort(Function(a, b) b.ReceivedTime.CompareTo(a.ReceivedTime))

        Return result
    End Function

    ''' <summary>
    ''' Builds an InboxBoardEntry from a MailItem.
    ''' Returns Nothing if the mail does not qualify for the board.
    ''' </summary>
    Private Function BuildInboxBoardEntry(mi As MailItem, categoryColors As Dictionary(Of String, String),
                                          includeFlagged As Boolean, hideDoneFlags As Boolean) As InboxBoardEntry
        Dim entry As New InboxBoardEntry()

        entry.EntryID = ComRetry(Of String)(Function() mi.EntryID)
        entry.Subject = ComRetry(Of String)(Function() If(mi.Subject, "(no subject)"))
        entry.SenderName = ComRetry(Of String)(Function() If(mi.SenderName, ""))
        entry.SenderEmail = ""
        Try : entry.SenderEmail = ComRetry(Of String)(Function() If(mi.SenderEmailAddress, "")) : Catch : End Try
        entry.Recipients = ""
        Try : entry.Recipients = ComRetry(Of String)(Function() CStr(mi.To)) : Catch : End Try
        entry.ReceivedTime = ComRetry(Of Date)(Function() mi.ReceivedTime)
        entry.IsRead = ComRetry(Of Boolean)(Function() Not mi.UnRead)

        ' MessageCount = 1 for ungrouped cards; GroupByConversation overwrites it for grouped cards.
        entry.MessageCount = 1
        entry.IsGrouped = False

        ' ConversationTopic for grouping
        entry.ConversationTopic = ""
        Try : entry.ConversationTopic = ComRetry(Of String)(Function() If(mi.ConversationTopic, "")) : Catch : End Try
        entry.ConversationEntryIDs = New List(Of String)() From {entry.EntryID}

        ' Flag properties
        entry.IsFlagged = False
        entry.IsCompleted = False
        entry.FlagRequest = ""
        entry.FlagDueDate = Nothing
        Try
            Dim flagStatus As Integer = ComRetry(Of Integer)(Function() CInt(mi.FlagStatus))
            ' olFlagMarked = 2, olFlagComplete = 1, olNoFlag = 0
            If flagStatus = 2 Then
                entry.IsFlagged = True
            ElseIf flagStatus = 1 Then
                entry.IsCompleted = True
            End If
        Catch
        End Try
        If entry.IsFlagged OrElse entry.IsCompleted Then
            Try : entry.FlagRequest = ComRetry(Of String)(Function() If(mi.FlagRequest, "")) : Catch : End Try
            Try
                Dim dueDate As DateTime = ComRetry(Of Date)(Function() mi.TaskDueDate)
                ' Outlook uses 1/1/4501 as "no date"
                If dueDate.Year < 4000 Then entry.FlagDueDate = dueDate
            Catch
            End Try
        End If

        ' Categories
        Dim cats As String = ""
        Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
        entry.AllCategories = cats

        Dim hasCats As Boolean = Not String.IsNullOrEmpty(cats)

        If hasCats Then
            Dim firstCat As String = cats.Split({","c, ";"c})(0).Trim()
            entry.Categories = firstCat
            entry.CategoryColor = "#6b7280" ' default gray
            If categoryColors.ContainsKey(firstCat) Then
                entry.CategoryColor = categoryColors(firstCat)
            End If
        Else
            ' No category — only include if flagged and includeFlagged is on
            If includeFlagged AndAlso entry.IsFlagged Then
                entry.Categories = ""
                entry.CategoryColor = InboxBoard_FlaggedColumnColor
            ElseIf includeFlagged AndAlso entry.IsCompleted AndAlso Not hideDoneFlags Then
                entry.Categories = ""
                entry.CategoryColor = "#9ca3af" ' muted gray for completed
            Else
                Return Nothing ' Does not qualify
            End If
        End If

        ' Body excerpt for AI summarization
        entry.BodyExcerpt = ""
        Try
            Dim body As String = GetMailBody(mi)
            If Not String.IsNullOrEmpty(body) Then
                body = GetLatestMailBody(body)
                If body.Length > InboxBoard_SummaryBodyCap Then
                    body = body.Substring(0, InboxBoard_SummaryBodyCap)
                End If
                entry.BodyExcerpt = body
            End If
        Catch
        End Try

        ' Compute flag bucket for date-grouped flag columns
        entry.FlagBucket = GetFlagBucket(entry)

        entry.Summary = ""
        Return entry
    End Function

    ''' <summary>
    ''' Determines which flag date-bucket a mail belongs to based on its flag status
    ''' and TaskDueDate. Returns the bucket column ID string.
    ''' </summary>
    Private Shared Function GetFlagBucket(entry As InboxBoardEntry) As String
        If entry.IsCompleted Then Return InboxBoard_FlagBucket_Done
        If Not entry.IsFlagged Then Return Nothing

        If Not entry.FlagDueDate.HasValue Then Return InboxBoard_FlagBucket_NoDueDate

        Dim dueDate As DateTime = entry.FlagDueDate.Value.Date
        Dim today As DateTime = DateTime.Today

        ' Overdue: before today
        If dueDate < today Then Return InboxBoard_FlagBucket_Overdue

        ' Today
        If dueDate = today Then Return InboxBoard_FlagBucket_Today

        ' Tomorrow
        If dueDate = today.AddDays(1) Then Return InboxBoard_FlagBucket_Tomorrow

        ' This week: up to end of current week (Sunday)
        Dim endOfWeek As DateTime = today.AddDays(7 - CInt(today.DayOfWeek))
        If dueDate <= endOfWeek Then Return InboxBoard_FlagBucket_ThisWeek

        ' Next week: Monday–Sunday of next week
        Dim endOfNextWeek As DateTime = endOfWeek.AddDays(7)
        If dueDate <= endOfNextWeek Then Return InboxBoard_FlagBucket_NextWeek

        ' Next 4 weeks: within 4 calendar weeks from today
        If dueDate <= today.AddDays(28) Then Return InboxBoard_FlagBucket_Next4Weeks

        ' Later: everything else
        Return InboxBoard_FlagBucket_Later
    End Function

    ''' <summary>
    ''' Gets a mapping of category name → hex color from Outlook's category list.
    ''' </summary>
    Private Function GetOutlookCategoryColors(ns As Outlook.NameSpace) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' Map OlCategoryColor enum values to hex colors
        Dim colorMap As New Dictionary(Of Integer, String)() From {
            {0, "#dc2626"},   ' olCategoryColorNone -> red
            {1, "#dc2626"},   ' Red
            {2, "#f97316"},   ' Orange
            {3, "#854d0e"},   ' Peach/Brown
            {4, "#eab308"},   ' Yellow
            {5, "#22c55e"},   ' Green
            {6, "#14b8a6"},   ' Teal
            {7, "#06b6d4"},   ' Olive -> Cyan
            {8, "#3b82f6"},   ' Blue
            {9, "#8b5cf6"},   ' Purple
            {10, "#d946ef"},  ' Maroon -> Fuchsia
            {11, "#64748b"},  ' Steel
            {12, "#6b7280"},  ' DarkSteel
            {13, "#9ca3af"},  ' Gray
            {14, "#475569"},  ' DarkGray
            {15, "#1e293b"},  ' Black
            {16, "#991b1b"},  ' DarkRed
            {17, "#c2410c"},  ' DarkOrange
            {18, "#a16207"},  ' DarkPeach
            {19, "#a3e635"},  ' DarkYellow -> Lime
            {20, "#15803d"},  ' DarkGreen
            {21, "#0d9488"},  ' DarkTeal
            {22, "#0284c7"},  ' DarkOlive -> Sky
            {23, "#1d4ed8"},  ' DarkBlue
            {24, "#7c3aed"},  ' DarkPurple
            {25, "#a21caf"}   ' DarkMaroon -> DarkFuchsia
        }

        Try
            Dim categories As Outlook.Categories = ns.Categories
            For Each cat As Outlook.Category In categories
                Try
                    Dim name As String = cat.Name
                    Dim colorVal As Integer = CInt(cat.Color)
                    Dim hexColor As String = "#6b7280"
                    If colorMap.ContainsKey(colorVal) Then hexColor = colorMap(colorVal)
                    result(name) = hexColor
                Catch
                End Try
            Next
        Catch
        End Try

        Return result
    End Function


#End Region

#Region "InboxBoard Column Building"

    ''' <summary>
    ''' Builds the list of board columns from found categories + always-present columns from settings.
    ''' If flagged mails are included, adds a synthetic "⚑ Flagged" column.
    ''' </summary>
    Private Function BuildBoardColumns(mails As List(Of InboxBoardEntry)) As List(Of InboxBoardColumn)
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)

        ' Collect unique categories from ALL categories on every mail (not just the first)
        Dim usedCats As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each m In mails
            If Not String.IsNullOrEmpty(m.AllCategories) Then
                For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then usedCats.Add(catName)
                Next
            End If
        Next

        ' Add always-present columns from settings
        Dim alwaysPresent As String = ""
        Try : alwaysPresent = If(My.Settings.InboxBoardColumns, "") : Catch : End Try
        If Not String.IsNullOrEmpty(alwaysPresent) Then
            For Each col In alwaysPresent.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim trimmed As String = col.Trim()
                If Not String.IsNullOrEmpty(trimmed) Then usedCats.Add(trimmed)
            Next
        End If

        ' Build column list
        Dim columns As New List(Of InboxBoardColumn)()
        For Each catName In usedCats
            Dim hexColor As String = "#6b7280"
            If categoryColors.ContainsKey(catName) Then hexColor = categoryColors(catName)

            columns.Add(New InboxBoardColumn With {
                .CategoryName = catName,
                .Color = hexColor,
                .Tag = catName,
                .IsFlaggedColumn = False
            })
        Next

        ' Add flagged column(s) when the setting is on
        Dim includeFlagged As Boolean = False
        Try : includeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
        If includeFlagged Then
            Dim groupByDate As Boolean = False
            Try : groupByDate = My.Settings.InboxBoardGroupFlagsByDate : Catch : End Try
            Dim hideDoneFlags As Boolean = True
            Try : hideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try

            If groupByDate Then
                ' Add one column per date bucket that has at least one mail, plus pinned buckets
                Dim usedBuckets As New HashSet(Of String)(StringComparer.Ordinal)
                For Each m In mails
                    If m.IsFlagged OrElse m.IsCompleted Then
                        Dim bucket = GetFlagBucket(m)
                        If bucket IsNot Nothing Then usedBuckets.Add(bucket)
                    End If
                Next

                ' Always show Today and Tomorrow even if empty, so users can drag into them
                usedBuckets.Add(InboxBoard_FlagBucket_Today)
                usedBuckets.Add(InboxBoard_FlagBucket_Tomorrow)

                ' Add pinned flag columns from settings (user-chosen always-visible buckets)
                Dim pinnedFlagCols As String = ""
                Try : pinnedFlagCols = If(My.Settings.InboxBoardPinnedFlagColumns, "") : Catch : End Try
                If Not String.IsNullOrEmpty(pinnedFlagCols) Then
                    For Each part In pinnedFlagCols.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) Then usedBuckets.Add(t)
                    Next
                End If

                ' Remove Done bucket if hidden
                If hideDoneFlags Then usedBuckets.Remove(InboxBoard_FlagBucket_Done)

                ' Insert in the defined display order, at position 0 (before category columns)
                Dim insertIdx As Integer = 0
                For Each bucketId In InboxBoard_FlagBucketOrder
                    If usedBuckets.Contains(bucketId) Then
                        Dim bucketColor As String = "#6b7280"
                        If InboxBoard_FlagBucketColors.ContainsKey(bucketId) Then bucketColor = InboxBoard_FlagBucketColors(bucketId)
                        columns.Insert(insertIdx, New InboxBoardColumn With {
                            .CategoryName = bucketId,
                            .Color = bucketColor,
                            .Tag = bucketId,
                            .IsFlaggedColumn = True
                        })
                        insertIdx += 1
                    End If
                Next
            Else
                ' Single "⚑ Flagged" column (original behavior)
                columns.Insert(0, New InboxBoardColumn With {
                    .CategoryName = InboxBoard_FlaggedColumnId,
                    .Color = InboxBoard_FlaggedColumnColor,
                    .Tag = InboxBoard_FlaggedColumnId,
                    .IsFlaggedColumn = True
                })
            End If
        End If

        ' Apply saved column order from My.Settings
        Dim savedOrder As String = ""
        Try : savedOrder = If(My.Settings.InboxBoardColumnOrder, "") : Catch : End Try
        If Not String.IsNullOrEmpty(savedOrder) Then
            Dim orderList As New List(Of String)()
            For Each part In savedOrder.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim t = part.Trim()
                If Not String.IsNullOrEmpty(t) Then orderList.Add(t)
            Next
            If orderList.Count > 0 Then
                columns.Sort(Function(a, b)
                                 Dim idxA As Integer = orderList.FindIndex(Function(o) o.Equals(a.CategoryName, StringComparison.OrdinalIgnoreCase))
                                 Dim idxB As Integer = orderList.FindIndex(Function(o) o.Equals(b.CategoryName, StringComparison.OrdinalIgnoreCase))
                                 ' Columns not in the saved order go to the end, preserving their relative position
                                 If idxA < 0 Then idxA = Integer.MaxValue
                                 If idxB < 0 Then idxB = Integer.MaxValue
                                 Return idxA.CompareTo(idxB)
                             End Function)
            End If
        End If

        Return columns
    End Function

#End Region

#Region "InboxBoard WebView2 Form"

    ''' <summary>
    ''' Shows the board in a WebView2-hosted WinForms dialog.
    ''' </summary>
    Private Sub ShowInboxBoardForm(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn))
        Dim frm As New Form()
        frm.Text = $"{AN} - Inbox Board"
        frm.MinimumSize = New Drawing.Size(900, 500)
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.MaximizeBox = True
        frm.MinimizeBox = True
        frm.ShowInTaskbar = True

        ' Restore window position, size and maximized state from My.Settings
        Dim savedX As Integer = 0, savedY As Integer = 0
        Dim savedW As Integer = 0, savedH As Integer = 0
        Dim savedMax As Integer = 0
        Try
            savedX = My.Settings.InboxBoardWindowX
            savedY = My.Settings.InboxBoardWindowY
            savedW = My.Settings.InboxBoardWindowW
            savedH = My.Settings.InboxBoardWindowH
            savedMax = My.Settings.InboxBoardWindowMax
        Catch
        End Try

        If savedW >= 900 AndAlso savedH >= 500 Then
            frm.Size = New Drawing.Size(savedW, savedH)
            ' Verify the saved position is still on a visible screen
            Dim savedRect As New Drawing.Rectangle(savedX, savedY, savedW, savedH)
            Dim onScreen As Boolean = False
            For Each scr As Screen In Screen.AllScreens
                If scr.WorkingArea.IntersectsWith(savedRect) Then
                    onScreen = True
                    Exit For
                End If
            Next
            If onScreen Then
                frm.StartPosition = FormStartPosition.Manual
                frm.Location = New Drawing.Point(savedX, savedY)
            Else
                frm.StartPosition = FormStartPosition.CenterScreen
            End If
        Else
            frm.Size = New Drawing.Size(1500, 850)
            frm.StartPosition = FormStartPosition.CenterScreen
        End If

        If savedMax <> 0 Then frm.WindowState = FormWindowState.Maximized

        Try
            Dim bmp As New Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            frm.Icon = Drawing.Icon.FromHandle(bmp.GetHicon())
            bmp.Dispose()
        Catch
        End Try

        Dim webView As New WebView2() With {
            .Dock = DockStyle.Fill
        }
        frm.Controls.Add(webView)

        ' Track mails + columns for callbacks
        Dim boardMails As List(Of InboxBoardEntry) = mails
        Dim boardColumns As List(Of InboxBoardColumn) = columns
        Dim summaryCts As New CancellationTokenSource()

        ' Temp file path for the board HTML (gives a real origin so localStorage works)
        Dim tempHtmlPath As String = ""

        AddHandler frm.FormClosing, Sub(s, e)
                                        summaryCts.Cancel()

                                        ' Persist summary cache on close
                                        SaveSummaryCache()

                                        ' Save window position, size and maximized state
                                        Try
                                            If frm.WindowState = FormWindowState.Maximized Then
                                                My.Settings.InboxBoardWindowMax = 1
                                                My.Settings.InboxBoardWindowX = frm.RestoreBounds.X
                                                My.Settings.InboxBoardWindowY = frm.RestoreBounds.Y
                                                My.Settings.InboxBoardWindowW = frm.RestoreBounds.Width
                                                My.Settings.InboxBoardWindowH = frm.RestoreBounds.Height
                                            Else
                                                My.Settings.InboxBoardWindowMax = 0
                                                My.Settings.InboxBoardWindowX = frm.Location.X
                                                My.Settings.InboxBoardWindowY = frm.Location.Y
                                                My.Settings.InboxBoardWindowW = frm.Size.Width
                                                My.Settings.InboxBoardWindowH = frm.Size.Height
                                            End If
                                            My.Settings.Save()
                                        Catch
                                        End Try

                                        ' Clean up temp file
                                        Try
                                            If Not String.IsNullOrEmpty(tempHtmlPath) AndAlso System.IO.File.Exists(tempHtmlPath) Then
                                                System.IO.File.Delete(tempHtmlPath)
                                            End If
                                        Catch
                                        End Try
                                    End Sub

        AddHandler frm.Shown, Async Sub(s, e)
                                  Try
                                      Dim userDataFolder As String = System.IO.Path.Combine(
                                          Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                          "RedInk", "WebView2")
                                      Dim env As CoreWebView2Environment = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, Nothing)
                                      Await webView.EnsureCoreWebView2Async(env)

                                      ' Handle messages from JS — use TryGetWebMessageAsString()
                                      ' because JS sends JSON.stringify(...) which is a string,
                                      ' not a raw JSON object. WebMessageAsJson would double-encode it.
                                      AddHandler webView.CoreWebView2.WebMessageReceived,
                                          Sub(sender, args)
                                              Dim raw As String = args.TryGetWebMessageAsString()
                                              If Not String.IsNullOrEmpty(raw) Then
                                                  HandleBoardMessage(raw, boardMails, boardColumns, webView, frm, summaryCts)
                                              End If
                                          End Sub

                                      ' Build HTML and write to a temp file so the page has a real
                                      ' file:// origin (about:blank blocks localStorage/sessionStorage).
                                      Dim html As String = BuildInboxBoardHtml(boardMails, boardColumns)
                                      tempHtmlPath = System.IO.Path.Combine(
                                          System.IO.Path.GetTempPath(),
                                          "RedInk_InboxBoard_" & Guid.NewGuid().ToString("N") & ".html")
                                      System.IO.File.WriteAllText(tempHtmlPath, html, Encoding.UTF8)
                                      webView.CoreWebView2.Navigate("file:///" & tempHtmlPath.Replace("\"c, "/"c))

                                      ' Start async summary generation after a short delay
                                      Await System.Threading.Tasks.Task.Delay(2000)
                                      If Not summaryCts.IsCancellationRequested Then
                                          Await GenerateSummariesAsync(boardMails, webView, summaryCts.Token)
                                      End If
                                  Catch ex As System.Exception
                                      Debug.WriteLine($"[InboxBoard] WebView2 init error: {ex.Message}")
                                  End Try
                              End Sub

        frm.Show()
    End Sub

    ''' <summary>
    ''' Handles messages sent from the board's JavaScript via postMessage.
    ''' </summary>
    Private Sub HandleBoardMessage(
    jsonMessage As String,
    mails As List(Of InboxBoardEntry),
    columns As List(Of InboxBoardColumn),
    webView As WebView2,
    frm As Form,
    ByRef summaryCts As CancellationTokenSource)

        Try
            Dim msg As JObject = JObject.Parse(jsonMessage)
            Dim action As String = If(CStr(msg("action")), "")

            Select Case action
                Case "move"
                    ' Replace: remove all existing categories, set only the new one
                    ' For grouped conversations, update ALL mails in the group
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim newCategory As String = CStr(msg("newCategory"))

                    ' If moving to any Flagged bucket column, set the flag + update due date
                    If InboxBoard_FlagBucketColors.ContainsKey(newCategory) OrElse newCategory = InboxBoard_FlaggedColumnId Then
                        Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                        If entry IsNot Nothing Then
                            ' When moving between flag buckets, only update the representative mail's flag.
                            ' The other mails in the conversation keep their own independent flag state.
                            ' Only apply to all grouped mails when moving FROM a category column TO flagged.
                            Dim movingBetweenFlagBuckets As Boolean = (entry.IsFlagged OrElse entry.IsCompleted) AndAlso String.IsNullOrEmpty(entry.AllCategories)
                            Dim idsToUpdate As List(Of String)
                            If movingBetweenFlagBuckets Then
                                idsToUpdate = New List(Of String) From {entryId}
                            Else
                                idsToUpdate = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                 entry.ConversationEntryIDs,
                                                 New List(Of String) From {entryId})
                            End If
                            Dim newDueDate As DateTime? = GetDueDateForBucket(newCategory)
                            For Each id In idsToUpdate
                                Try
                                    Dim outlookAppFlag As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                    Dim nsFlag As Outlook.NameSpace = outlookAppFlag.GetNamespace("MAPI")
                                    Dim miFlag As MailItem = TryCast(ComRetry(Of Object)(Function() nsFlag.GetItemFromID(id)), MailItem)
                                    If miFlag IsNot Nothing Then
                                        ' Only clear categories when entry actually had categories before
                                        ' (don't wipe categories when just moving between flag buckets)
                                        If Not String.IsNullOrEmpty(entry.AllCategories) Then
                                            miFlag.Categories = ""
                                        End If
                                        If newCategory = InboxBoard_FlagBucket_Done Then
                                            miFlag.FlagStatus = OlFlagStatus.olFlagComplete
                                        Else
                                            ' Always (re-)establish the flag and dates. Outlook needs
                                            ' FlagStatus set before dates, and TaskStartDate before
                                            ' TaskDueDate when the new start is after the current due.
                                            miFlag.FlagStatus = OlFlagStatus.olFlagMarked
                                            If newDueDate.HasValue Then
                                                miFlag.TaskStartDate = newDueDate.Value
                                                miFlag.TaskDueDate = newDueDate.Value
                                            Else
                                                miFlag.TaskStartDate = #1/1/4501#
                                                miFlag.TaskDueDate = #1/1/4501#
                                            End If
                                        End If
                                        miFlag.Save()
                                    End If
                                Catch ex As System.Exception
                                    Debug.WriteLine($"[InboxBoard] move-to-flag error: {ex.Message}")
                                End Try
                            Next
                            ' Only blank categories in local state if they were actually cleared in Outlook
                            If Not String.IsNullOrEmpty(entry.AllCategories) Then
                                entry.Categories = ""
                                entry.AllCategories = ""
                            End If
                            If newCategory = InboxBoard_FlagBucket_Done Then
                                entry.IsFlagged = False
                                entry.IsCompleted = True
                            Else
                                entry.IsFlagged = True
                                entry.IsCompleted = False
                            End If
                            entry.FlagDueDate = GetDueDateForBucket(newCategory)
                            entry.FlagBucket = newCategory
                            entry.CategoryColor = If(InboxBoard_FlagBucketColors.ContainsKey(newCategory), InboxBoard_FlagBucketColors(newCategory), InboxBoard_FlaggedColumnColor)
                        End If
                        PushUpdatedCardToJs(entryId, mails, webView)
                        Return
                    End If

                    Dim entryMove = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entryMove IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entryMove.IsGrouped AndAlso entryMove.ConversationEntryIDs IsNot Nothing AndAlso entryMove.ConversationEntryIDs.Count > 1,
                                                                 entryMove.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            UpdateMailCategory(id, newCategory)
                            ' If moving out of Flagged column, clear the flag
                            If entryMove.IsFlagged Then SetMailFlag(id, False)
                        Next

                        ' Update local state
                        entryMove.Categories = newCategory
                        entryMove.AllCategories = newCategory
                        If entryMove.IsFlagged Then
                            entryMove.IsFlagged = False
                            entryMove.IsCompleted = False
                        End If
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entryMove.CategoryColor = If(cc.ContainsKey(newCategory), cc(newCategory), "#6b7280")
                    End If

                    ' Push updated card data back to JS so category badges refresh
                    PushUpdatedCardToJs(entryId, mails, webView)

                Case "moveAdd"
                    ' Add: keep existing categories, append the new one
                    ' For grouped conversations, update ALL mails in the group
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim addCategory As String = CStr(msg("newCategory"))

                    ' If adding to any Flagged bucket column, set the flag + update due date
                    If InboxBoard_FlagBucketColors.ContainsKey(addCategory) OrElse addCategory = InboxBoard_FlaggedColumnId Then
                        Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                        If entry IsNot Nothing Then
                            Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                     entry.ConversationEntryIDs,
                                                                     New List(Of String) From {entryId})
                            Dim newDueDate As DateTime? = GetDueDateForBucket(addCategory)
                            For Each id In idsToUpdate
                                Try
                                    Dim outlookAppFlag As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                    Dim nsFlag As Outlook.NameSpace = outlookAppFlag.GetNamespace("MAPI")
                                    Dim miFlag As MailItem = TryCast(ComRetry(Of Object)(Function() nsFlag.GetItemFromID(id)), MailItem)
                                    If miFlag IsNot Nothing Then
                                        If addCategory = InboxBoard_FlagBucket_Done Then
                                            miFlag.FlagStatus = OlFlagStatus.olFlagComplete
                                        Else
                                            miFlag.FlagStatus = OlFlagStatus.olFlagMarked
                                            If newDueDate.HasValue Then
                                                miFlag.TaskStartDate = newDueDate.Value
                                                miFlag.TaskDueDate = newDueDate.Value
                                            Else
                                                miFlag.TaskStartDate = #1/1/4501#
                                                miFlag.TaskDueDate = #1/1/4501#
                                            End If
                                        End If
                                        miFlag.Save()
                                    End If
                                Catch ex As System.Exception
                                    Debug.WriteLine($"[InboxBoard] moveAdd-to-flag error: {ex.Message}")
                                End Try
                            Next
                            If addCategory = InboxBoard_FlagBucket_Done Then
                                entry.IsFlagged = False
                                entry.IsCompleted = True
                            Else
                                entry.IsFlagged = True
                                entry.IsCompleted = False
                            End If
                            entry.FlagDueDate = GetDueDateForBucket(addCategory)
                            entry.FlagBucket = addCategory
                        End If
                        PushUpdatedCardToJs(entryId, mails, webView)
                        Return
                    End If

                    Dim entryAdd = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entryAdd IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entryAdd.IsGrouped AndAlso entryAdd.ConversationEntryIDs IsNot Nothing AndAlso entryAdd.ConversationEntryIDs.Count > 1,
                                                                 entryAdd.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            AddMailCategory(id, addCategory)
                        Next

                        ' Update local state
                        Dim existing As New List(Of String)()
                        If Not String.IsNullOrEmpty(entryAdd.AllCategories) Then
                            For Each part In entryAdd.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                                Dim t = part.Trim()
                                If Not String.IsNullOrEmpty(t) Then existing.Add(t)
                            Next
                        End If
                        If Not existing.Any(Function(c) c.Equals(addCategory, StringComparison.OrdinalIgnoreCase)) Then
                            existing.Add(addCategory)
                        End If
                        entryAdd.AllCategories = String.Join(", ", existing)
                        entryAdd.Categories = existing(0)
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entryAdd.CategoryColor = If(cc.ContainsKey(entryAdd.Categories), cc(entryAdd.Categories), "#6b7280")
                    End If

                    ' Push updated card data back to JS so category badges refresh
                    PushUpdatedCardToJs(entryId, mails, webView)

                Case "unmark"
                    ' Remove ONE category from the mail(s). If the card has multiple categories,
                    ' only the category of the column where unmark was clicked is removed.
                    ' If it was the last category, the card is removed entirely.
                    ' For grouped conversations, all mails in the group are updated.
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim categoryToRemove As String = If(CStr(msg("category")), "")

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry Is Nothing Then Return

                    ' If unmarking from any Flagged bucket column, clear the flag
                    If InboxBoard_FlagBucketColors.ContainsKey(categoryToRemove) OrElse categoryToRemove = InboxBoard_FlaggedColumnId Then
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            SetMailFlag(id, False)
                        Next
                        entry.IsFlagged = False
                        entry.IsCompleted = False

                        ' If the card has no categories either, remove it from the board
                        If String.IsNullOrEmpty(entry.AllCategories) Then
                            mails.RemoveAll(Function(m) m.EntryID = entryId)
                            Try
                                webView.CoreWebView2.PostWebMessageAsJson(
                                    JsonConvert.SerializeObject(New With {.action = "removeCard", .entryId = entryId}))
                            Catch
                            End Try
                        Else
                            PushUpdatedCardToJs(entryId, mails, webView)
                        End If
                        Return
                    End If

                    ' Determine all entry IDs to process (group or single)
                    Dim idsToProcess As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                             entry.ConversationEntryIDs,
                                                             New List(Of String) From {entryId})

                    ' Check how many categories the card currently has
                    Dim currentCats As New List(Of String)()
                    If Not String.IsNullOrEmpty(entry.AllCategories) Then
                        For Each part In entry.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                            Dim t = part.Trim()
                            If Not String.IsNullOrEmpty(t) Then currentCats.Add(t)
                        Next
                    End If

                    Dim isLastCategory As Boolean = (currentCats.Count <= 1) OrElse
                                                     String.IsNullOrEmpty(categoryToRemove)

                    If isLastCategory Then
                        ' If flagged, keep on board (in Flagged column) but remove categories
                        If entry.IsFlagged Then
                            For Each id In idsToProcess
                                ClearMailCategories(id)
                            Next
                            entry.Categories = ""
                            entry.AllCategories = ""
                            entry.CategoryColor = InboxBoard_FlaggedColumnColor
                            PushUpdatedCardToJs(entryId, mails, webView)
                        Else
                            ' Last (or only) category: clear all categories from all mails, remove card
                            For Each id In idsToProcess
                                ClearMailCategories(id)
                                ' Also clear any flag so it doesn't linger in Outlook
                                If entry.IsFlagged OrElse entry.IsCompleted Then
                                    SetMailFlag(id, False)
                                End If
                            Next
                            entry.IsFlagged = False
                            entry.IsCompleted = False
                            mails.RemoveAll(Function(m) m.EntryID = entryId)

                            ' Tell JS to remove the card
                            Try
                                webView.CoreWebView2.PostWebMessageAsJson(
                                    JsonConvert.SerializeObject(New With {.action = "removeCard", .entryId = entryId}))
                            Catch
                            End Try
                        End If
                    Else
                        ' Multiple categories: remove only the clicked category from all mails
                        For Each id In idsToProcess
                            RemoveMailCategory(id, categoryToRemove)
                        Next

                        ' Update local state
                        Dim remaining = currentCats.Where(Function(c) Not c.Equals(categoryToRemove, StringComparison.OrdinalIgnoreCase)).ToList()
                        entry.AllCategories = String.Join(", ", remaining)
                        entry.Categories = If(remaining.Count > 0, remaining(0), "")

                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entry.CategoryColor = If(remaining.Count > 0 AndAlso cc.ContainsKey(remaining(0)), cc(remaining(0)), "#6b7280")

                        ' Push updated card to JS — renderBoard will remove it from the column
                        PushUpdatedCardToJs(entryId, mails, webView)
                    End If

                Case "open"
                    ' Open mail in Outlook Inspector
                    Dim entryId As String = CStr(msg("entryId"))
                    OpenMailInInspector(entryId)

                Case "reload"
                    ' Re-scan inbox and push new data, reusing the user's last load-count choice
                    summaryCts.Cancel()
                    Dim newCts As New CancellationTokenSource()
                    summaryCts = newCts

                    Dim reloadMax As Integer = 0
                    Dim newMails = CollectCategorizedInboxMails(reloadMax)
                    If newMails IsNot Nothing Then
                        mails.Clear()
                        mails.AddRange(newMails)
                        Dim newColumns = BuildBoardColumns(mails)
                        columns.Clear()
                        columns.AddRange(newColumns)

                        Dim dataJson As String = BuildBoardDataJson(mails, columns)
                        Try
                            webView.CoreWebView2.PostWebMessageAsJson(
                                JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJson}))
                        Catch
                        End Try

                        ' Re-generate summaries
                        Dim capturedMails As List(Of InboxBoardEntry) = mails
                        Dim capturedWebView As WebView2 = webView
                        Dim capturedToken As CancellationToken = newCts.Token
                        System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                            Await System.Threading.Tasks.Task.Delay(500)
                                                            frm.BeginInvoke(CType(Async Sub()
                                                                                      Await GenerateSummariesAsync(capturedMails, capturedWebView, capturedToken)
                                                                                  End Sub, System.Windows.Forms.MethodInvoker))
                                                        End Function)
                    End If

                Case "setPinnedColumns"
                    ' Save pinned column names to My.Settings
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardColumns = String.Join(";", names)
                            My.Settings.Save()
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setPinnedColumns error: {ex2.Message}")
                    End Try

                Case "setPinnedFlagColumns"
                    ' Save pinned flag bucket column names to My.Settings and reload
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardPinnedFlagColumns = String.Join(";", names)
                            My.Settings.Save()
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setPinnedFlagColumns error: {ex2.Message}")
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setColumnOrder"
                    ' Save user's column order to My.Settings
                    Try
                        Dim colsToken As JToken = msg("columns")
                        If colsToken IsNot Nothing AndAlso colsToken.Type = JTokenType.Array Then
                            Dim names As New List(Of String)()
                            For Each t As JToken In colsToken
                                Dim n As String = CStr(t)
                                If Not String.IsNullOrWhiteSpace(n) Then names.Add(n.Trim())
                            Next
                            My.Settings.InboxBoardColumnOrder = String.Join(";", names)
                            My.Settings.Save()

                            ' Update in-memory column order to match
                            Dim newOrder As New List(Of InboxBoardColumn)()
                            For Each name In names
                                Dim col = columns.FirstOrDefault(Function(c) c.CategoryName.Equals(name, StringComparison.OrdinalIgnoreCase))
                                If col IsNot Nothing Then newOrder.Add(col)
                            Next
                            ' Append any columns not in the saved order
                            For Each col In columns
                                If Not newOrder.Any(Function(c) c.CategoryName.Equals(col.CategoryName, StringComparison.OrdinalIgnoreCase)) Then
                                    newOrder.Add(col)
                                End If
                            Next
                            columns.Clear()
                            columns.AddRange(newOrder)
                        End If
                    Catch ex2 As System.Exception
                        Debug.WriteLine($"[InboxBoard] setColumnOrder error: {ex2.Message}")
                    End Try

                Case "setTheme"
                    ' Persist theme choice to My.Settings
                    Try
                        Dim theme As String = CStr(msg("theme"))
                        If Not String.IsNullOrEmpty(theme) Then
                            My.Settings.InboxBoardTheme = theme
                            My.Settings.Save()
                        End If
                    Catch
                    End Try

                Case "setFieldSettings"
                    ' Persist card field toggles as JSON string
                    Try
                        Dim fields As String = CStr(msg("fields"))
                        If Not String.IsNullOrEmpty(fields) Then
                            My.Settings.InboxBoardFields = fields
                            My.Settings.Save()
                        End If
                    Catch
                    End Try

                Case "setColumnFilter"
                    ' Persist column filter selection
                    Try
                        Dim filterVal As String = CStr(msg("filter"))
                        My.Settings.InboxBoardColumnFilter = If(filterVal, "")
                        My.Settings.Save()
                    Catch
                    End Try

                Case "setGroupConversations"
                    ' Persist conversation grouping toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardGroupConversations = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the grouping change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setIncludeFlagged"
                    ' Persist include-flagged toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardIncludeFlagged = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setHideDoneFlags"
                    ' Persist hide-done-flags toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardHideDoneFlags = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setGroupFlagsByDate"
                    ' Persist group-flags-by-date toggle and reload
                    Try
                        Dim enabled As Boolean = CBool(msg("enabled"))
                        My.Settings.InboxBoardGroupFlagsByDate = enabled
                        My.Settings.Save()
                    Catch
                    End Try

                    ' Trigger a reload to apply the change
                    HandleBoardReload(mails, columns, webView, frm, summaryCts)

                Case "setSummaryLanguage"
                    ' Persist summary language selection; clear cache if language changed
                    Try
                        Dim lang As String = If(CStr(msg("language")), "").Trim()
                        Dim oldLang As String = ""
                        Try : oldLang = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try

                        My.Settings.InboxBoardSummaryLanguage = lang
                        My.Settings.Save()

                        ' If the language actually changed, clear the cache and re-generate
                        If Not lang.Equals(oldLang, StringComparison.OrdinalIgnoreCase) Then
                            ClearSummaryCache()

                            ' Clear in-memory summaries so they will be re-generated
                            For Each m In mails
                                m.Summary = ""
                            Next

                            ' Push cleared summaries to JS so cards show "Generating…"
                            For Each m In mails
                                Try
                                    webView.CoreWebView2.PostWebMessageAsJson(
                                        JsonConvert.SerializeObject(New With {
                                            .action = "updateSummary",
                                            .entryId = m.EntryID,
                                            .summary = ""
                                        }))
                                Catch
                                End Try
                            Next

                            ' Restart summary generation with the new language
                            summaryCts.Cancel()
                            Dim newCtsLang As New CancellationTokenSource()
                            summaryCts = newCtsLang

                            Dim capturedMailsLang As List(Of InboxBoardEntry) = mails
                            Dim capturedWebViewLang As WebView2 = webView
                            Dim capturedTokenLang As CancellationToken = newCtsLang.Token
                            System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                                Await System.Threading.Tasks.Task.Delay(500)
                                                                frm.BeginInvoke(CType(Async Sub()
                                                                                          Await GenerateSummariesAsync(capturedMailsLang, capturedWebViewLang, capturedTokenLang)
                                                                                      End Sub, System.Windows.Forms.MethodInvoker))
                                                            End Function)
                        End If
                    Catch
                    End Try
            End Select

        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] HandleBoardMessage error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Helper to trigger a full board reload (used by multiple settings toggles).
    ''' </summary>
    Private Sub HandleBoardReload(
        mails As List(Of InboxBoardEntry),
        columns As List(Of InboxBoardColumn),
        webView As WebView2,
        frm As Form,
        ByRef summaryCts As CancellationTokenSource)

        summaryCts.Cancel()
        Dim newCts As New CancellationTokenSource()
        summaryCts = newCts

        Dim reloadMax As Integer = 0
        Dim newMails = CollectCategorizedInboxMails(reloadMax)
        If newMails IsNot Nothing Then
            mails.Clear()
            mails.AddRange(newMails)
            Dim newColumns = BuildBoardColumns(mails)
            columns.Clear()
            columns.AddRange(newColumns)
            Dim dataJson As String = BuildBoardDataJson(mails, columns)
            Try
                webView.CoreWebView2.PostWebMessageAsJson(
                    JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJson}))
            Catch
            End Try
            Dim capturedMails As List(Of InboxBoardEntry) = mails
            Dim capturedWebView As WebView2 = webView
            Dim capturedToken As CancellationToken = newCts.Token
            System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                Await System.Threading.Tasks.Task.Delay(500)
                                                frm.BeginInvoke(CType(Async Sub()
                                                                          Await GenerateSummariesAsync(capturedMails, capturedWebView, capturedToken)
                                                                      End Sub, System.Windows.Forms.MethodInvoker))
                                            End Function)
        End If
    End Sub

    ''' <summary>
    ''' Pushes updated card data (allCategories, category, color, flag info) to JS after a move/moveAdd.
    ''' </summary>
    Private Sub PushUpdatedCardToJs(entryId As String, mails As List(Of InboxBoardEntry), webView As WebView2)
        Try
            Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
            If entry Is Nothing Then Return

            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim categoryColors = GetOutlookCategoryColors(ns)

            Dim allCatsArr As New JArray()
            If Not String.IsNullOrEmpty(entry.AllCategories) Then
                For Each catPart In entry.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then
                        Dim catObj As New JObject()
                        catObj("name") = catName
                        catObj("color") = If(categoryColors.ContainsKey(catName), categoryColors(catName), "#6b7280")
                        allCatsArr.Add(catObj)
                    End If
                Next
            End If

            webView.CoreWebView2.PostWebMessageAsJson(
                JsonConvert.SerializeObject(New With {
                    .action = "updateCard",
                    .entryId = entryId,
                    .category = entry.Categories,
                    .categoryColor = entry.CategoryColor,
                    .allCategories = allCatsArr,
                    .isFlagged = entry.IsFlagged,
                    .isCompleted = entry.IsCompleted,
                    .flagRequest = If(entry.FlagRequest, ""),
                    .flagDueDate = If(entry.FlagDueDate.HasValue, entry.FlagDueDate.Value.ToString("yyyy-MM-dd"), ""),
                                   .flagBucket = If(entry.FlagBucket, "")
                }))
        Catch
        End Try
    End Sub

#End Region

#Region "InboxBoard Outlook COM Operations"

    ''' <summary>
    ''' Sets the TaskDueDate on a flagged mail item.
    ''' Pass Nothing to clear the due date.
    ''' </summary>
    Private Sub SetMailDueDate(entryId As String, dueDate As DateTime?)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                If dueDate.HasValue Then
                    mi.TaskDueDate = dueDate.Value
                    mi.TaskStartDate = dueDate.Value
                Else
                    ' Set to Outlook's "no date" sentinel (1/1/4501)
                    mi.TaskDueDate = #1/1/4501#
                    mi.TaskStartDate = #1/1/4501#
                End If
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailDueDate error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Returns a representative due date for a given flag bucket column ID.
    ''' Used when dragging a card into a date-bucketed column.
    ''' Returns Nothing for the "no due date" / generic flagged bucket.
    ''' </summary>
    Private Shared Function GetDueDateForBucket(bucketId As String) As DateTime?
        Dim today As DateTime = DateTime.Today

        Select Case bucketId
            Case InboxBoard_FlagBucket_Overdue
                ' Set to yesterday (it's already overdue)
                Return today.AddDays(-1)
            Case InboxBoard_FlagBucket_Today
                Return today
            Case InboxBoard_FlagBucket_Tomorrow
                Return today.AddDays(1)
            Case InboxBoard_FlagBucket_ThisWeek
                ' Set to end of current week (Friday)
                Dim daysUntilFriday As Integer = (5 - CInt(today.DayOfWeek) + 7) Mod 7
                If daysUntilFriday = 0 Then daysUntilFriday = 5 ' If today is Friday, keep Friday
                Return today.AddDays(daysUntilFriday)
            Case InboxBoard_FlagBucket_NextWeek
                ' Set to next Monday + 4 (next Friday)
                Dim daysUntilNextMonday As Integer = ((1 - CInt(today.DayOfWeek)) + 7) Mod 7
                If daysUntilNextMonday = 0 Then daysUntilNextMonday = 7
                Return today.AddDays(daysUntilNextMonday + 4) ' Next Friday
            Case InboxBoard_FlagBucket_Next4Weeks
                Return today.AddDays(21) ' 3 weeks out
            Case InboxBoard_FlagBucket_Later
                Return today.AddDays(42) ' 6 weeks out
            Case InboxBoard_FlagBucket_Done
                Return Nothing ' Don't change due date for completed items
            Case Else
                Return Nothing ' Generic flagged / no due date
        End Select
    End Function

    ''' <summary>
    ''' Changes the category on a mail item (replaces all existing categories).
    ''' </summary>
    Private Sub UpdateMailCategory(entryId As String, newCategory As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Categories = newCategory
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] UpdateMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Adds a category to a mail item, keeping existing categories.
    ''' </summary>
    Private Sub AddMailCategory(entryId As String, categoryToAdd As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                Dim existing As String = ComRetry(Of String)(Function() If(mi.Categories, ""))
                Dim cats As New List(Of String)()
                If Not String.IsNullOrEmpty(existing) Then
                    For Each part In existing.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) Then cats.Add(t)
                    Next
                End If
                If Not cats.Any(Function(c) c.Equals(categoryToAdd, StringComparison.OrdinalIgnoreCase)) Then
                    cats.Add(categoryToAdd)
                End If
                mi.Categories = String.Join(", ", cats)
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] AddMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Removes a single category from a mail item, keeping all other categories.
    ''' </summary>
    Private Sub RemoveMailCategory(entryId As String, categoryToRemove As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                Dim existing As String = ComRetry(Of String)(Function() If(mi.Categories, ""))
                Dim cats As New List(Of String)()
                If Not String.IsNullOrEmpty(existing) Then
                    For Each part In existing.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim t = part.Trim()
                        If Not String.IsNullOrEmpty(t) AndAlso
                           Not t.Equals(categoryToRemove, StringComparison.OrdinalIgnoreCase) Then
                            cats.Add(t)
                        End If
                    Next
                End If
                mi.Categories = String.Join(", ", cats)
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] RemoveMailCategory error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Clears all categories from a mail item.
    ''' </summary>
    Private Sub ClearMailCategories(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Categories = ""
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] ClearMailCategories error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Sets or clears the follow-up flag on a mail item.
    ''' When setting, uses olFlagMarked with "Follow up" request.
    ''' When clearing, calls ClearTaskFlag to remove all flag/task properties.
    ''' </summary>
    Private Sub SetMailFlag(entryId As String, flagOn As Boolean)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                If flagOn Then
                    mi.FlagStatus = OlFlagStatus.olFlagMarked
                    mi.FlagRequest = "Follow up"
                Else
                    mi.ClearTaskFlag()
                End If
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailFlag error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Marks a mail item's flag as complete (olFlagComplete).
    ''' </summary>
    Private Sub SetMailFlagComplete(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.FlagStatus = OlFlagStatus.olFlagComplete
                mi.Save()
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] SetMailFlagComplete error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Opens a mail in an Outlook Inspector window.
    ''' </summary>
    Private Sub OpenMailInInspector(entryId As String)
        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim mi As MailItem = TryCast(ComRetry(Of Object)(Function() ns.GetItemFromID(entryId)), MailItem)
            If mi IsNot Nothing Then
                mi.Display(False)
            End If
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] OpenMailInInspector error: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "InboxBoard AI Summaries"

    ''' <summary>
    ''' Returns the language instruction to append to the summary system prompt.
    ''' Uses the free-text language field; empty means auto-detect.
    ''' </summary>
    Private Function GetSummaryLanguageInstruction() As String
        Dim lang As String = ""
        Try : lang = If(My.Settings.InboxBoardSummaryLanguage, "").Trim() : Catch : End Try

        ' If empty, treat as auto-detect — no explicit instruction
        If String.IsNullOrEmpty(lang) Then Return ""

        Return $" Write all summaries in {lang}."
    End Function

    ''' <summary>
    ''' Generates AI summaries for mails in batches and pushes them to the board.
    ''' Skips mails that already have a cached summary.
    ''' </summary>
    Private Async Function GenerateSummariesAsync(
        mails As List(Of InboxBoardEntry),
        webView As WebView2,
        ct As CancellationToken) As System.Threading.Tasks.Task

        Dim unsummarized = mails.Where(Function(m) String.IsNullOrEmpty(m.Summary) AndAlso
                                                    Not String.IsNullOrEmpty(m.BodyExcerpt)).ToList()
        If unsummarized.Count = 0 Then Return

        Dim langInstruction As String = GetSummaryLanguageInstruction()
        Dim batchIndex As Integer = 0

        While batchIndex < unsummarized.Count
            If ct.IsCancellationRequested Then Return

            Dim batchStart As Integer = batchIndex
            Dim batchEnd As Integer = Math.Min(batchIndex + InboxBoard_SummaryBatchSize - 1, unsummarized.Count - 1)

            Dim userPrompt As New StringBuilder()
            For i As Integer = batchStart To batchEnd
                Dim mail As InboxBoardEntry = unsummarized(i)
                Dim mailNum As Integer = i - batchStart + 1
                userPrompt.AppendLine($"<EMAIL id=""{mailNum}"">")
                userPrompt.AppendLine($"Subject: {mail.Subject}")
                userPrompt.AppendLine($"From: {mail.SenderName}")
                userPrompt.AppendLine($"Date: {mail.ReceivedTime:yyyy-MM-dd}")
                If Not String.IsNullOrEmpty(mail.BodyExcerpt) Then
                    userPrompt.AppendLine($"Body: {mail.BodyExcerpt}")
                End If
                userPrompt.AppendLine("</EMAIL>")
            Next

            Try
                Dim systemPrompt As String = InterpolateAtRuntime(SP_InboxBoard) & langInstruction
                Dim response As String = Await LLM(
                    systemPrompt, userPrompt.ToString(),
                    "", "", 0, False, True)

                If Not String.IsNullOrWhiteSpace(response) Then
                    ParseAndApplySummaries(response, unsummarized, batchStart, batchEnd, webView)
                End If
            Catch ex As System.Exception
                Debug.WriteLine($"[InboxBoard] Summary batch error: {ex.Message}")
            End Try

            batchIndex = batchEnd + 1
        End While

        ' Persist cache after all batches are done
        SaveSummaryCache()
    End Function

    ''' <summary>
    ''' Parses summary JSON and pushes updates to the board via postMessage.
    ''' Also caches summaries for future use.
    ''' </summary>
    Private Sub ParseAndApplySummaries(
        response As String,
        mails As List(Of InboxBoardEntry),
        batchStart As Integer, batchEnd As Integer,
        webView As WebView2)

        Try
            Dim jsonText As String = response.Trim()
            If jsonText.StartsWith("```") Then
                Dim startIdx As Integer = jsonText.IndexOf("[")
                Dim endIdx As Integer = jsonText.LastIndexOf("]")
                If startIdx >= 0 AndAlso endIdx > startIdx Then
                    jsonText = jsonText.Substring(startIdx, endIdx - startIdx + 1)
                End If
            End If

            Dim arr As JArray = JArray.Parse(jsonText)

            For Each item As JObject In arr
                Dim id As Integer = If(item("id") IsNot Nothing, CInt(item("id")), 0)
                Dim summary As String = If(item("summary") IsNot Nothing, CStr(item("summary")), "")
                If String.IsNullOrEmpty(summary) Then Continue For

                Dim absoluteIndex As Integer = batchStart + id - 1
                If absoluteIndex < batchStart OrElse absoluteIndex > batchEnd OrElse absoluteIndex >= mails.Count Then Continue For

                mails(absoluteIndex).Summary = summary

                ' Cache the summary
                CacheSummary(mails(absoluteIndex).EntryID, summary)

                ' Push to JS
                Try
                    Dim entryId As String = mails(absoluteIndex).EntryID
                    webView.CoreWebView2.PostWebMessageAsJson(
                        JsonConvert.SerializeObject(New With {
                            .action = "updateSummary",
                            .entryId = entryId,
                            .summary = summary
                        }))
                Catch
                End Try
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"[InboxBoard] ParseAndApplySummaries error: {ex.Message}")
        End Try
    End Sub

#End Region

#Region "InboxBoard JSON Helpers"

    ''' <summary>
    ''' Builds a JSON string with card and column data for JS initialization.
    ''' </summary>
    Private Function BuildBoardDataJson(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn)) As String
        Dim categoryColors As Dictionary(Of String, String) = Nothing

        Dim cardsArr As New JArray()
        For i As Integer = 0 To mails.Count - 1
            Dim m = mails(i)
            Dim card As New JObject()
            card("entryId") = m.EntryID
            card("subject") = If(m.Subject, "")
            card("senderName") = If(m.SenderName, "")
            card("senderEmail") = If(m.SenderEmail, "")
            card("recipients") = If(m.Recipients, "")
            card("date") = m.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
            card("messages") = m.MessageCount
            card("isGrouped") = m.IsGrouped
            card("category") = If(m.Categories, "")
            card("categoryColor") = If(m.CategoryColor, "#6b7280")
            card("summary") = If(m.Summary, "")
            card("isRead") = m.IsRead
            card("isConversation") = m.IsGrouped
            card("isFlagged") = m.IsFlagged
            card("isCompleted") = m.IsCompleted
            card("flagRequest") = If(m.FlagRequest, "")
            card("flagDueDate") = If(m.FlagDueDate.HasValue, m.FlagDueDate.Value.ToString("yyyy-MM-dd"), "")
            card("flagBucket") = If(m.FlagBucket, "")

            ' Build array of all categories with colors for multi-category badges
            Dim allCatsArr As New JArray()
            If Not String.IsNullOrEmpty(m.AllCategories) Then
                If categoryColors Is Nothing Then
                    Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                    Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                    categoryColors = GetOutlookCategoryColors(ns)
                End If
                For Each catPart In m.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim catName As String = catPart.Trim()
                    If Not String.IsNullOrEmpty(catName) Then
                        Dim catObj As New JObject()
                        catObj("name") = catName
                        catObj("color") = If(categoryColors.ContainsKey(catName), categoryColors(catName), "#6b7280")
                        allCatsArr.Add(catObj)
                    End If
                Next
            End If
            card("allCategories") = allCatsArr

            cardsArr.Add(card)
        Next

        Dim colsArr As New JArray()
        For Each col In columns
            Dim c As New JObject()
            c("id") = col.CategoryName
            c("title") = col.CategoryName
            c("tag") = If(col.Tag, col.CategoryName)
            c("color") = If(col.Color, "#6b7280")
            c("isFlaggedColumn") = col.IsFlaggedColumn
            colsArr.Add(c)
        Next

        Dim root As New JObject()
        root("cards") = cardsArr
        root("columns") = colsArr
        Return root.ToString(Formatting.None)
    End Function

#End Region

#Region "InboxBoard HTML Builder"

    ''' <summary>
    ''' Builds the complete HTML page for the Inbox Board.
    ''' </summary>
    Private Function BuildInboxBoardHtml(mails As List(Of InboxBoardEntry), columns As List(Of InboxBoardColumn)) As String
        Dim dataJson As String = BuildBoardDataJson(mails, columns)

        ' Read persisted settings to inject into JS as constants
        Dim savedTheme As String = ""
        Dim savedFields As String = ""
        Dim savedColumnFilter As String = ""
        Dim savedGroupConversations As Boolean = False
        Dim savedSummaryLanguage As String = ""
        Dim savedIncludeFlagged As Boolean = False
        Dim savedHideDoneFlags As Boolean = True
        Dim savedGroupFlagsByDate As Boolean = False
        Try : savedTheme = If(My.Settings.InboxBoardTheme, "") : Catch : End Try
        Try : savedFields = If(My.Settings.InboxBoardFields, "") : Catch : End Try
        Try : savedColumnFilter = If(My.Settings.InboxBoardColumnFilter, "") : Catch : End Try
        Try : savedGroupConversations = My.Settings.InboxBoardGroupConversations : Catch : End Try
        Try : savedSummaryLanguage = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try
        Try : savedIncludeFlagged = My.Settings.InboxBoardIncludeFlagged : Catch : End Try
        Try : savedHideDoneFlags = My.Settings.InboxBoardHideDoneFlags : Catch : End Try
        Try : savedGroupFlagsByDate = My.Settings.InboxBoardGroupFlagsByDate : Catch : End Try

        ' If no saved language yet, default to INI_Language1 (full language name, e.g. "German")
        ' but do NOT persist it — let the user confirm by editing the field first.
        ' This way the field shows a sensible default without locking it in.
        If String.IsNullOrEmpty(savedSummaryLanguage) Then
            Try
                Dim lang1 As String = If(_context.INI_Language1, "")
                If Not String.IsNullOrEmpty(lang1) Then
                    savedSummaryLanguage = lang1
                End If
            Catch
            End Try
        End If

        Dim sb As New StringBuilder(32000)
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<html lang=""en"">")
        sb.AppendLine("<head>")
        sb.AppendLine("<meta charset=""UTF-8"">")
        sb.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">")
        sb.AppendLine($"<title>{AN} - Inbox Board</title>")
        sb.AppendLine("<style>")
        sb.AppendLine(GetBoardCss())
        sb.AppendLine("</style>")
        sb.AppendLine("</head>")
        sb.AppendLine("<body>")
        sb.AppendLine(GetBoardHeaderHtml())
        sb.AppendLine("<div class=""board"" id=""board""></div>")
        sb.AppendLine("<div class=""toast-container"" id=""toastContainer""></div>")
        sb.AppendLine("<div class=""drag-mode-label"" id=""dragModeLabel""><span class=""plus"">+</span> Add category (Ctrl held)</div>")
        sb.AppendLine("<script>")
        sb.AppendLine($"const INIT_DATA = {dataJson};")
        sb.AppendLine($"const SAVED_THEME = {JsonConvert.SerializeObject(savedTheme)};")
        sb.AppendLine($"const SAVED_FIELDS = {JsonConvert.SerializeObject(savedFields)};")
        sb.AppendLine($"const SAVED_COLUMN_FILTER = {JsonConvert.SerializeObject(savedColumnFilter)};")
        sb.AppendLine($"const SAVED_GROUP_CONVERSATIONS = {If(savedGroupConversations, "true", "false")};")
        sb.AppendLine($"const SAVED_SUMMARY_LANGUAGE = {JsonConvert.SerializeObject(savedSummaryLanguage)};")
        sb.AppendLine($"const SAVED_INCLUDE_FLAGGED = {If(savedIncludeFlagged, "true", "false")};")
        sb.AppendLine($"const SAVED_HIDE_DONE_FLAGS = {If(savedHideDoneFlags, "true", "false")};")
        sb.AppendLine($"const SAVED_GROUP_FLAGS_BY_DATE = {If(savedGroupFlagsByDate, "true", "false")};")
        sb.AppendLine($"const FLAGGED_COLUMN_ID = {JsonConvert.SerializeObject(InboxBoard_FlaggedColumnId)};")
        sb.AppendLine(GetBoardJavaScript())
        sb.AppendLine("</script>")
        sb.AppendLine("</body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function


    ''' <summary>Returns the CSS for the board.</summary>
    Private Function GetBoardCss() As String
        Return "
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg: #f0f2f5; --header-bg: #ffffff; --header-border: #e0e0e0;
  --col-bg: #f7f8fa; --col-border: #e2e4e8; --card-bg: #ffffff;
  --card-border: #e0e0e0; --card-shadow: rgba(0,0,0,0.06);
  --text: #1a1a1a; --text-secondary: #555; --text-muted: #888;
  --input-bg: #f0f2f5; --input-border: #d0d0d0;
  --drop-highlight: rgba(59,130,246,0.12); --drop-line: #3b82f6;
  --drop-highlight-add: rgba(34,197,94,0.12); --drop-line-add: #22c55e;
  --toast-bg: #323232; --toast-text: #fff;
  --overlay-bg: rgba(0,0,0,0.05); --settings-bg: #ffffff; --settings-border: #e0e0e0;
}
[data-theme=""dark""] {
  --bg: #1a1b1e; --header-bg: #25262b; --header-border: #373a40;
  --col-bg: #25262b; --col-border: #373a40; --card-bg: #2c2e33;
  --card-border: #373a40; --card-shadow: rgba(0,0,0,0.2);
  --text: #e0e0e0; --text-secondary: #b0b0b0; --text-muted: #777;
  --input-bg: #2c2e33; --input-border: #444;
  --drop-highlight: rgba(59,130,246,0.2);
  --drop-highlight-add: rgba(34,197,94,0.2); --drop-line-add: #22c55e;
  --toast-bg: #444; --toast-text: #fff;
  --overlay-bg: rgba(0,0,0,0.2); --settings-bg: #2c2e33; --settings-border: #444;
}
body { font-family: system-ui, 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--bg); color: var(--text); min-height: 100vh; transition: background 0.3s, color 0.3s; }
.header { background: var(--header-bg); border-bottom: 1px solid var(--header-border);
  padding: 10px 20px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
  position: sticky; top: 0; z-index: 100; transition: background 0.3s; }
.header .topline { display: flex; align-items: center; gap: 8px; min-width: 0; }
.header .topline img.logo { width: 24px; height: 24px; border-radius: 6px; display: block; }
.header .topline .brandbig { font-size: 18px; font-weight: 700; white-space: nowrap; }
.header .topline .sub { font-size: 15px; color: var(--text-muted); white-space: nowrap; }
.header-controls { display: flex; align-items: center; gap: 8px; flex: 1; justify-content: flex-end; flex-wrap: wrap; }
.search-input { padding: 6px 10px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); font-size: 13px; width: 200px; outline: none; }
.search-input:focus { border-color: #3b82f6; }
.search-input::placeholder { color: var(--text-muted); }
.filter-select { padding: 6px 8px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); font-size: 13px; outline: none; cursor: pointer; }
.icon-btn { width: 34px; height: 34px; border: 1px solid var(--input-border); border-radius: 6px;
  background: var(--input-bg); color: var(--text); cursor: pointer; display: flex;
  align-items: center; justify-content: center; font-size: 16px; position: relative; }
.icon-btn:hover { border-color: #3b82f6; }
.icon-btn:active { transform: scale(0.93); }
.settings-wrapper { position: relative; }
.settings-panel { position: absolute; top: 40px; right: 0; background: var(--settings-bg);
  border: 1px solid var(--settings-border); border-radius: 8px; padding: 10px 14px;
  min-width: 220px; max-height: 420px; overflow-y: auto;
  box-shadow: 0 4px 16px var(--card-shadow); z-index: 200; display: none; }
.settings-panel.open { display: block; }
.settings-panel h3 { font-size: 12px; font-weight: 600; margin-bottom: 6px;
  color: var(--text-secondary); text-transform: uppercase; letter-spacing: 0.5px; }
.settings-panel h3.section-divider { margin-top: 10px; padding-top: 8px;
  border-top: 1px solid var(--settings-border); }
.settings-panel label { display: flex; align-items: center; gap: 6px; padding: 3px 0;
  font-size: 13px; cursor: pointer; color: var(--text); }
.settings-panel input[type=""checkbox""] { width: 15px; height: 15px; accent-color: #3b82f6; cursor: pointer; }
.settings-panel select { padding: 4px 6px; border: 1px solid var(--input-border); border-radius: 4px;
  background: var(--input-bg); color: var(--text); font-size: 12px; outline: none; cursor: pointer; width: 100%; }
.settings-panel .cat-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.lang-input-wrapper { position: relative; width: 100%; }
.lang-input-wrapper input[type=""text""] { padding: 4px 6px; border: 1px solid var(--input-border); border-radius: 4px;
  background: var(--input-bg); color: var(--text); font-size: 12px; outline: none; width: 100%; }
.lang-input-wrapper input[type=""text""]:focus { border-color: #3b82f6; }
.lang-input-wrapper .lang-overlay { position: absolute; top: 0; left: 0; right: 0; bottom: 0;
  padding: 4px 6px; font-size: 12px; color: var(--text-muted); pointer-events: none;
  display: flex; align-items: center; opacity: 0.7; }
.lang-input-wrapper input[type=""text""]:not([data-empty=""true""]) + .lang-overlay { display: none; }
.lang-input-wrapper input[type=""text""]:focus + .lang-overlay { display: none; }
.board { display: flex; gap: 14px; padding: 16px 20px; overflow-x: auto;
  min-height: calc(100vh - 56px); align-items: flex-start; }
.column { min-width: 220px; max-width: 340px; flex: 1 1 280px; background: var(--col-bg);
  border: 1px solid var(--col-border); border-radius: 10px; display: flex; flex-direction: column; }
.column.drag-over { background: var(--drop-highlight); }
.column.drag-over-add { background: var(--drop-highlight-add); }
.column-header { padding: 10px 12px; font-weight: 600; font-size: 13px; display: flex;
  align-items: center; gap: 6px; border-bottom: 1px solid var(--col-border); user-select: none; }
.column-header .dot { width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }
.column-header .count { margin-left: auto; background: var(--input-bg); color: var(--text-muted);
  font-size: 11px; font-weight: 500; padding: 1px 7px; border-radius: 10px; }
.card-list { padding: 6px; flex: 1; min-height: 50px; display: flex; flex-direction: column; gap: 0; }
.drop-indicator { height: 3px; background: var(--drop-line); border-radius: 2px;
  margin: 2px 4px; opacity: 0; flex-shrink: 0; }
.drop-indicator.visible { opacity: 1; }
.drop-indicator.add-mode { background: var(--drop-line-add); }
.card { background: var(--card-bg); border: 1px solid var(--card-border); border-radius: 8px;
  padding: 9px 11px; cursor: grab; box-shadow: 0 1px 3px var(--card-shadow);
  user-select: none; margin-bottom: 6px; transition: box-shadow 0.2s, transform 0.15s; }
.card:hover { box-shadow: 0 3px 8px var(--card-shadow); transform: translateY(-1px); }
.card.dragging { opacity: 0.5; cursor: grabbing; transform: rotate(2deg); }
.card.hidden { display: none; }
.card.conversation { border-left: 3px solid #3b82f6; }
.card.flagged { border-left: 3px solid #ef4444; }
.card.conversation.flagged { border-left: 3px solid #ef4444; }
.card-subject { font-weight: 600; font-size: 13px; margin-bottom: 4px; color: var(--text); line-height: 1.3;
  overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.card-sender { font-size: 11px; color: var(--text-secondary); margin-bottom: 3px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.card-meta { display: flex; align-items: center; gap: 8px; font-size: 10px;
  color: var(--text-muted); margin-bottom: 3px; flex-wrap: wrap; }
.card-summary { font-size: 11px; color: var(--text-secondary); line-height: 1.3;
  margin-bottom: 4px; font-style: italic; min-height: 14px; }
.card-summary.loading { color: var(--text-muted); }
.card-footer { display: flex; justify-content: flex-end; gap: 4px; }
.card-btn { background: none; border: none; color: var(--text-muted); cursor: pointer;
  font-size: 13px; padding: 2px 4px; border-radius: 4px; display: flex;
  align-items: center; gap: 3px; position: relative; }
.card-btn:hover { color: #3b82f6; background: var(--drop-highlight); }
.card-btn .tip { position: absolute; bottom: 100%; right: 0; background: var(--toast-bg);
  color: var(--toast-text); font-size: 10px; padding: 3px 6px; border-radius: 4px;
  white-space: nowrap; opacity: 0; pointer-events: none; margin-bottom: 3px; }
.card-btn:hover .tip { opacity: 1; }
.unread-dot { width: 7px; height: 7px; border-radius: 50%; background: #3b82f6;
  display: inline-block; margin-right: 4px; flex-shrink: 0; }
.conv-icon { font-size: 10px; color: var(--text-muted); margin-right: 3px; }
.flag-badge { font-size: 10px; color: #ef4444; margin-right: 3px; }
.flag-meta { font-size: 10px; color: #ef4444; white-space: nowrap; }
.flag-meta.completed { color: #22c55e; text-decoration: line-through; }
.card-categories { display: flex; flex-wrap: wrap; gap: 3px; margin-bottom: 4px; }
.card-cat-badge { font-size: 9px; padding: 1px 6px; border-radius: 8px; color: #fff;
  white-space: nowrap; line-height: 1.4; font-weight: 500; }
.toast-container { position: fixed; bottom: 16px; right: 16px; display: flex;
  flex-direction: column-reverse; gap: 6px; z-index: 1000; pointer-events: none; }
.toast { background: var(--toast-bg); color: var(--toast-text); padding: 8px 14px;
  border-radius: 8px; font-size: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  max-width: 380px; line-height: 1.3; animation: toastIn 0.3s ease; pointer-events: auto; }
.toast.fade-out { animation: toastOut 0.3s ease forwards; }
@keyframes toastIn { from { opacity:0; transform:translateY(10px); } to { opacity:1; transform:translateY(0); } }
@keyframes toastOut { from { opacity:1; transform:translateY(0); } to { opacity:0; transform:translateY(10px); } }
.icon-svg { width: 15px; height: 15px; display: inline-block; vertical-align: middle; }
.drag-mode-label { position: fixed; bottom: 16px; left: 16px; background: var(--toast-bg);
  color: var(--toast-text); padding: 6px 12px; border-radius: 8px; font-size: 12px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.2); z-index: 1000; pointer-events: none;
  display: none; }
.drag-mode-label.visible { display: block; }
.drag-mode-label .plus { color: #22c55e; font-weight: 700; margin-right: 4px; }
.column-header { cursor: grab; }
.column-header:active { cursor: grabbing; }
.column.col-dragging { opacity: 0.5; }
.column.col-drag-over { border-left: 3px solid #3b82f6; }
"
    End Function

    ''' <summary>Returns the header HTML for the board.</summary>
    Private Function GetBoardHeaderHtml() As String
        Dim logoDataUrl As String = GetLogoDataUrl()
        Dim brandHtml As String = System.Net.WebUtility.HtmlEncode(If(Not String.IsNullOrWhiteSpace(AN), AN, "Red Ink"))

        Dim logoImg As String = ""
        If Not String.IsNullOrWhiteSpace(logoDataUrl) Then
            logoImg = $"<img class=""logo"" src=""{System.Net.WebUtility.HtmlEncode(logoDataUrl)}"" alt=""logo"">"
        End If

        ' Build pinned columns checkboxes from all Outlook categories
        Dim pinnedHtml As New StringBuilder()
        Dim pinnedSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Dim alwaysPresent As String = If(My.Settings.InboxBoardColumns, "")
            If Not String.IsNullOrEmpty(alwaysPresent) Then
                For Each col In alwaysPresent.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim trimmed As String = col.Trim()
                    If Not String.IsNullOrEmpty(trimmed) Then pinnedSet.Add(trimmed)
                Next
            End If
        Catch
        End Try

        Try
            Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
            Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
            Dim categoryColors As Dictionary(Of String, String) = GetOutlookCategoryColors(ns)
            Dim categories As Outlook.Categories = ns.Categories
            For Each cat As Outlook.Category In categories
                Try
                    Dim catName As String = cat.Name
                    Dim escapedName As String = System.Net.WebUtility.HtmlEncode(catName)
                    Dim hexColor As String = "#6b7280"
                    If categoryColors.ContainsKey(catName) Then hexColor = categoryColors(catName)
                    Dim isChecked As String = If(pinnedSet.Contains(catName), " checked", "")
                    pinnedHtml.AppendLine($"        <label><input type=""checkbox"" data-pinned-cat=""{escapedName}""{isChecked}><span class=""cat-dot"" style=""background:{hexColor}""></span> {escapedName}</label>")
                Catch
                End Try
            Next
        Catch
        End Try

        ' Conversation grouping checkbox
        Dim groupChecked As String = ""
        Try : If My.Settings.InboxBoardGroupConversations Then groupChecked = " checked"
        Catch : End Try

        ' Include flagged checkbox
        Dim flaggedChecked As String = ""
        Try : If My.Settings.InboxBoardIncludeFlagged Then flaggedChecked = " checked"
        Catch : End Try

        ' Hide done flags checkbox
        Dim hideDoneChecked As String = ""
        Try : If My.Settings.InboxBoardHideDoneFlags Then hideDoneChecked = " checked"
        Catch : End Try

        ' Group flagged by date checkbox
        Dim groupFlagsByDateChecked As String = ""
        Try : If My.Settings.InboxBoardGroupFlagsByDate Then groupFlagsByDateChecked = " checked"
        Catch : End Try

        ' Summary language: free text input with "Auto-detect" overlay
        Dim savedLang As String = ""
        Try : savedLang = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try
        Dim escapedLang As String = System.Net.WebUtility.HtmlEncode(savedLang)
        Dim dataEmpty As String = If(String.IsNullOrEmpty(savedLang), " data-empty=""true""", "")

        Return $"
<div class=""header"">
  <div class=""topline"">
    {logoImg}
    <div class=""brandbig"">{brandHtml}</div>
    <div class=""sub"">Inbox Board</div>
  </div>
  <div class=""header-controls"">
    <input type=""text"" class=""search-input"" id=""searchInput"" placeholder=""Search subjects, senders…"">
    <select class=""filter-select"" id=""columnFilter""><option value=""all"">All Columns</option></select>
    <div class=""settings-wrapper"">
      <button class=""icon-btn"" id=""settingsBtn"" title=""Card display settings"">
        <svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><circle cx=""12"" cy=""12"" r=""3""/><path d=""M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.32 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z""/></svg>
      </button>
      <div class=""settings-panel"" id=""settingsPanel"">
        <h3>Card Fields</h3>
        <label><input type=""checkbox"" data-field=""sender"" checked> Sender</label>
        <label><input type=""checkbox"" data-field=""date"" checked> Date</label>
        <label><input type=""checkbox"" data-field=""messageCount"" checked> Message count</label>
        <label><input type=""checkbox"" data-field=""summary"" checked> AI Summary</label>
        <h3 class=""section-divider"">Grouping</h3>
        <label><input type=""checkbox"" id=""groupConversationsChk""{groupChecked}> Group conversations</label>
        <h3 class=""section-divider"">Flags</h3>
        <label><input type=""checkbox"" id=""includeFlaggedChk""{flaggedChecked}> Include flagged mails</label>
        <label><input type=""checkbox"" id=""hideDoneFlagsChk""{hideDoneChecked}> Hide completed flags</label>
        <label><input type=""checkbox"" id=""groupFlagsByDateChk""{groupFlagsByDateChecked}> Group flagged by due date</label>
        <h3 class=""section-divider"">Summary Language</h3>
        <div class=""lang-input-wrapper"">
          <input type=""text"" id=""summaryLanguageInput"" value=""{escapedLang}""{dataEmpty}>
          <div class=""lang-overlay"">Auto-detect</div>
        </div>
        <h3 class=""section-divider"">Pinned Columns</h3>
{pinnedHtml.ToString().TrimEnd()}
        <h3 class=""section-divider"">Pinned Flag Columns</h3>
        <div id=""pinnedFlagColumnsSection"" style=""display:{If(groupFlagsByDateChecked <> "", "block", "none")}"">
{GetPinnedFlagColumnsHtml()}
        </div>
        <div id=""pinnedFlagColumnsHint"" style=""display:{If(groupFlagsByDateChecked = "", "block", "none")};font-size:11px;color:var(--text-muted);padding:2px 0"">
          Enable &quot;Group flagged by due date&quot; first
        </div>
        <h3 class=""section-divider"">Drag &amp; Drop</h3>
        <div style=""font-size:11px;color:var(--text-muted);line-height:1.4;padding:2px 0"">
          <b>Drag</b> → replace all categories<br>
          <b>Ctrl + Drag</b> → add category
        </div>
      </div>
    </div>
    <button class=""icon-btn"" id=""reloadBtn"" title=""Reload mails from Inbox"">
      <svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><polyline points=""23 4 23 10 17 10""/><polyline points=""1 20 1 14 7 14""/><path d=""M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15""/></svg>
    </button>
    <button class=""icon-btn"" id=""themeToggle"" title=""Toggle theme"">
      <svg class=""icon-svg"" id=""themeIcon"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><circle cx=""12"" cy=""12"" r=""5""/><line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/><line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/><line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/><line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/></svg>
    </button>
  </div>
</div>"
    End Function

    ''' <summary>Builds checkboxes for pinning individual flag date-bucket columns.</summary>
    Private Function GetPinnedFlagColumnsHtml() As String
        Dim sb As New StringBuilder()
        Dim pinnedSet As New HashSet(Of String)(StringComparer.Ordinal)
        Try
            Dim saved As String = If(My.Settings.InboxBoardPinnedFlagColumns, "")
            If Not String.IsNullOrEmpty(saved) Then
                For Each part In saved.Split({";"c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t = part.Trim()
                    If Not String.IsNullOrEmpty(t) Then pinnedSet.Add(t)
                Next
            End If
        Catch
        End Try

        For Each bucketId In InboxBoard_FlagBucketOrder
            Dim escapedId As String = System.Net.WebUtility.HtmlEncode(bucketId)
            Dim bucketColor As String = "#6b7280"
            If InboxBoard_FlagBucketColors.ContainsKey(bucketId) Then bucketColor = InboxBoard_FlagBucketColors(bucketId)
            Dim isChecked As String = If(pinnedSet.Contains(bucketId), " checked", "")
            sb.AppendLine($"          <label><input type=""checkbox"" data-pinned-flag=""{escapedId}""{isChecked}><span class=""cat-dot"" style=""background:{bucketColor}""></span> {escapedId}</label>")
        Next

        Return sb.ToString().TrimEnd()
    End Function


    ''' <summary>Returns the JavaScript for the board.</summary>
    Private Function GetBoardJavaScript() As String
        Dim js As New StringBuilder(20000)
        js.AppendLine("// --- State ---")
        js.AppendLine("var cards = [];")
        js.AppendLine("var COLUMNS = [];")
        js.AppendLine("var fieldSettings = { sender: true, date: true, messageCount: true, summary: true };")
        js.AppendLine("var isDragging = false;")
        js.AppendLine("var ctrlHeld = false;")
        js.AppendLine("var langCommitTimer = null;")
        js.AppendLine()
        js.AppendLine("function init(data) {")
        js.AppendLine("  try {")
        js.AppendLine("    var d = (typeof data === 'string') ? JSON.parse(data) : data;")
        js.AppendLine("    COLUMNS = d.columns || [];")
        js.AppendLine("    cards = (d.cards || []).map(function(c, i) { return Object.assign({}, c, { id: i + 1 }); });")
        js.AppendLine("    console.log('InboxBoard init: ' + cards.length + ' cards, ' + COLUMNS.length + ' columns');")
        js.AppendLine("    populateColumnFilter();")
        js.AppendLine("    renderBoard();")
        js.AppendLine("    loadFieldSettings();")
        js.AppendLine("    applyFieldVisibility();")
        js.AppendLine("    restoreColumnFilter();")
        js.AppendLine("  } catch(err) { document.body.innerHTML = '<pre style=""color:red;padding:20px"">' + err.message + '\n' + err.stack + '</pre>'; }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function populateColumnFilter() {")
        js.AppendLine("  var sel = document.getElementById('columnFilter');")
        js.AppendLine("  sel.innerHTML = '<option value=""all"">All Columns</option>';")
        js.AppendLine("  COLUMNS.forEach(function(col) {")
        js.AppendLine("    var opt = document.createElement('option');")
        js.AppendLine("    opt.value = col.id;")
        js.AppendLine("    opt.textContent = col.title;")
        js.AppendLine("    sel.appendChild(opt);")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function restoreColumnFilter() {")
        js.AppendLine("  if (SAVED_COLUMN_FILTER) {")
        js.AppendLine("    var sel = document.getElementById('columnFilter');")
        js.AppendLine("    for (var i = 0; i < sel.options.length; i++) {")
        js.AppendLine("      if (sel.options[i].value === SAVED_COLUMN_FILTER) { sel.selectedIndex = i; break; }")
        js.AppendLine("    }")
        js.AppendLine("    applySearchFilter();")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Persistence via My.Settings (postMessage to VB.NET) ---")
        js.AppendLine("function saveState() {")
        js.AppendLine("  try {")
        js.AppendLine("    var state = COLUMNS.map(function(col) {")
        js.AppendLine("      var colEl = document.querySelector('.column[data-column=""' + CSS.escape(col.id) + '""]');")
        js.AppendLine("      if (!colEl) return { id: col.id, cardIds: [] };")
        js.AppendLine("      var cids = [];")
        js.AppendLine("      colEl.querySelector('.card-list').querySelectorAll('.card').forEach(function(c) { cids.push(c.dataset.entryId); });")
        js.AppendLine("      return { id: col.id, cardIds: cids };")
        js.AppendLine("    });")
        js.AppendLine("    localStorage.setItem('rib-columns', JSON.stringify(state));")
        js.AppendLine("  } catch(err) { console.error('saveState error:', err); }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function saveFieldSettings() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setFieldSettings', fields: JSON.stringify(fieldSettings) }));")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function loadFieldSettings() {")
        js.AppendLine("  if (SAVED_FIELDS) { try { fieldSettings = JSON.parse(SAVED_FIELDS); } catch(e) {} }")
        js.AppendLine("  document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-field]').forEach(function(cb) {")
        js.AppendLine("    cb.checked = fieldSettings[cb.dataset.field] !== false;")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Theme (unchanged) ---
        js.AppendLine("function applyTheme(theme) {")
        js.AppendLine("  document.documentElement.setAttribute('data-theme', theme);")
        js.AppendLine("  var icon = document.getElementById('themeIcon');")
        js.AppendLine("  if (theme === 'dark') {")
        js.AppendLine("    icon.innerHTML = '<path d=""M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z""/>';")
        js.AppendLine("  } else {")
        js.AppendLine("    icon.innerHTML = '<circle cx=""12"" cy=""12"" r=""5""/>' +")
        js.AppendLine("      '<line x1=""12"" y1=""1"" x2=""12"" y2=""3""/><line x1=""12"" y1=""21"" x2=""12"" y2=""23""/>' +")
        js.AppendLine("      '<line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/><line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/>' +")
        js.AppendLine("      '<line x1=""1"" y1=""12"" x2=""3"" y2=""12""/><line x1=""21"" y1=""12"" x2=""23"" y2=""12""/>' +")
        js.AppendLine("      '<line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/><line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/>';")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine("function initTheme() {")
        js.AppendLine("  var theme = SAVED_THEME || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');")
        js.AppendLine("  applyTheme(theme);")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("document.getElementById('themeToggle').addEventListener('click', function() {")
        js.AppendLine("  var cur = document.documentElement.getAttribute('data-theme');")
        js.AppendLine("  var next = cur === 'dark' ? 'light' : 'dark';")
        js.AppendLine("  applyTheme(next);")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setTheme', theme: next }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Settings panel ---
        js.AppendLine("document.getElementById('settingsBtn').addEventListener('click', function(e) {")
        js.AppendLine("  e.stopPropagation(); document.getElementById('settingsPanel').classList.toggle('open');")
        js.AppendLine("});")
        js.AppendLine("document.addEventListener('click', function(e) {")
        js.AppendLine("  var p = document.getElementById('settingsPanel');")
        js.AppendLine("  if (!p.contains(e.target) && e.target !== document.getElementById('settingsBtn')) p.classList.remove('open');")
        js.AppendLine("});")
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-field]').forEach(function(cb) {")
        js.AppendLine("  cb.addEventListener('change', function() {")
        js.AppendLine("    fieldSettings[cb.dataset.field] = cb.checked;")
        js.AppendLine("    saveFieldSettings(); applyFieldVisibility();")
        js.AppendLine("  });")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Group conversations toggle ---
        js.AppendLine("document.getElementById('groupConversationsChk').addEventListener('change', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setGroupConversations', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Include flagged toggle ---
        js.AppendLine("document.getElementById('includeFlaggedChk').addEventListener('change', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setIncludeFlagged', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Hide done flags toggle ---
        js.AppendLine("document.getElementById('hideDoneFlagsChk').addEventListener('change', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setHideDoneFlags', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Summary language (free text input with debounce) ---
        js.AppendLine("(function() {")
        js.AppendLine("  var langInput = document.getElementById('summaryLanguageInput');")
        js.AppendLine("  var lastCommitted = SAVED_SUMMARY_LANGUAGE || '';")
        js.AppendLine("  if (SAVED_SUMMARY_LANGUAGE) { langInput.value = SAVED_SUMMARY_LANGUAGE; }")
        js.AppendLine("  function updateEmpty() { langInput.setAttribute('data-empty', langInput.value.trim() === '' ? 'true' : 'false'); }")
        js.AppendLine("  updateEmpty();")
        js.AppendLine("  function commitLang() {")
        js.AppendLine("    if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; }")
        js.AppendLine("    var val = langInput.value.trim();")
        js.AppendLine("    if (val !== lastCommitted) {")
        js.AppendLine("      lastCommitted = val;")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'setSummaryLanguage', language: val }));")
        js.AppendLine("      if (val) { showToast('Summary language: ' + val); } else { showToast('Summary language: Auto-detect'); }")
        js.AppendLine("    }")
        js.AppendLine("  }")
        js.AppendLine("  langInput.addEventListener('input', function() { updateEmpty(); if (langCommitTimer) clearTimeout(langCommitTimer); langCommitTimer = setTimeout(commitLang, 1500); });")
        js.AppendLine("  langInput.addEventListener('blur', function() { if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } commitLang(); });")
        js.AppendLine("  langInput.addEventListener('keydown', function(e) { if (e.key === 'Enter') { e.preventDefault(); if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } commitLang(); langInput.blur(); } });")
        js.AppendLine("})();")
        js.AppendLine()
        ' --- Pinned columns ---
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(cb) {")
        js.AppendLine("  cb.addEventListener('change', function() {")
        js.AppendLine("    var pinned = [];")
        js.AppendLine("    document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(c) { if (c.checked) pinned.push(c.dataset.pinnedCat); });")
        js.AppendLine("    window.chrome.webview.postMessage(JSON.stringify({ action: 'setPinnedColumns', columns: pinned }));")
        js.AppendLine("    showToast('Pinned columns updated');")
        js.AppendLine("  });")
        js.AppendLine("});")
        ' --- Pinned flag columns ---
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-flag]').forEach(function(cb) {")
        js.AppendLine("  cb.addEventListener('change', function() {")
        js.AppendLine("    var pinned = [];")
        js.AppendLine("    document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-flag]').forEach(function(c) { if (c.checked) pinned.push(c.dataset.pinnedFlag); });")
        js.AppendLine("    window.chrome.webview.postMessage(JSON.stringify({ action: 'setPinnedFlagColumns', columns: pinned }));")
        js.AppendLine("    showToast('Pinned flag columns updated');")
        js.AppendLine("  });")
        js.AppendLine("});")
        js.AppendLine()

        js.AppendLine()
        js.AppendLine("function applyFieldVisibility() {")
        js.AppendLine("  document.querySelectorAll('.card-sender').forEach(function(el) { el.style.display = fieldSettings.sender ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-date').forEach(function(el) { el.style.display = fieldSettings.date ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-msg-count').forEach(function(el) { el.style.display = fieldSettings.messageCount ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-summary').forEach(function(el) { el.style.display = fieldSettings.summary ? '' : 'none'; });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Reload ---
        js.AppendLine("document.getElementById('groupFlagsByDateChk').addEventListener('change', function() {")
        js.AppendLine("  var sec = document.getElementById('pinnedFlagColumnsSection'); var hint = document.getElementById('pinnedFlagColumnsHint');")
        js.AppendLine("  if (sec) sec.style.display = this.checked ? 'block' : 'none';")
        js.AppendLine("  if (hint) hint.style.display = this.checked ? 'none' : 'block';")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setGroupFlagsByDate', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Reload button ---
        js.AppendLine("document.getElementById('reloadBtn').addEventListener('click', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'reload' }));")
        js.AppendLine("});")
        js.AppendLine()
        ' --- Ctrl key tracking for drag mode ---
        js.AppendLine("document.addEventListener('keydown', function(e) { if (e.key === 'Control' && !ctrlHeld) { ctrlHeld = true; updateDragModeVisuals(); } });")
        js.AppendLine("document.addEventListener('keyup', function(e) { if (e.key === 'Control') { ctrlHeld = false; updateDragModeVisuals(); } });")
        js.AppendLine("window.addEventListener('blur', function() { ctrlHeld = false; updateDragModeVisuals(); });")
        js.AppendLine()
        js.AppendLine("function updateDragModeVisuals() {")
        js.AppendLine("  var label = document.getElementById('dragModeLabel');")
        js.AppendLine("  if (isDragging && ctrlHeld) {")
        js.AppendLine("    label.classList.add('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over').forEach(function(el) { el.classList.add('drag-over-add'); });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.visible').forEach(function(el) { el.classList.add('add-mode'); });")
        js.AppendLine("  } else {")
        js.AppendLine("    label.classList.remove('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over-add').forEach(function(el) { el.classList.remove('drag-over-add'); });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.add-mode').forEach(function(el) { el.classList.remove('add-mode'); });")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Render Board ---
        js.AppendLine("function renderBoard() {")
        js.AppendLine("  var board = document.getElementById('board');")
        js.AppendLine("  board.innerHTML = '';")
        js.AppendLine("  COLUMNS.forEach(function(col) {")
        js.AppendLine("    var colEl = document.createElement('div');")
        js.AppendLine("    colEl.className = 'column';")
        js.AppendLine("    colEl.dataset.column = col.id;")
        js.AppendLine("    var colCards = cards.filter(function(c) {")
        js.AppendLine("      if (col.isFlaggedColumn) {")
        js.AppendLine("        if (!(c.isFlagged || c.isCompleted)) return false;")
        js.AppendLine("        if (col.id === FLAGGED_COLUMN_ID) return true;")
        js.AppendLine("        if (c.flagBucket) return c.flagBucket === col.id;")
        js.AppendLine("        return false;")
        js.AppendLine("      }")
        js.AppendLine("      if (c.allCategories && c.allCategories.length > 0) {")
        js.AppendLine("        return c.allCategories.some(function(ac) { return ac.name === col.id; });")
        js.AppendLine("      }")
        js.AppendLine("      return c.category === col.id;")
        js.AppendLine("    });")
        js.AppendLine("    var headerDiv = document.createElement('div');")
        js.AppendLine("    headerDiv.className = 'column-header';")
        js.AppendLine("    var dot = document.createElement('span');")
        js.AppendLine("    dot.className = 'dot';")
        js.AppendLine("    dot.style.background = col.color;")
        js.AppendLine("    var titleSpan = document.createElement('span');")
        js.AppendLine("    titleSpan.textContent = col.title;")
        js.AppendLine("    var countSpan = document.createElement('span');")
        js.AppendLine("    countSpan.className = 'count';")
        js.AppendLine("    countSpan.textContent = colCards.length;")
        js.AppendLine("    headerDiv.appendChild(dot);")
        js.AppendLine("    headerDiv.appendChild(titleSpan);")
        js.AppendLine("    headerDiv.appendChild(countSpan);")
        js.AppendLine("    colEl.appendChild(headerDiv);")
        js.AppendLine("    var list = document.createElement('div');")
        js.AppendLine("    list.className = 'card-list';")
        js.AppendLine("    list.dataset.column = col.id;")
        js.AppendLine("    colEl.appendChild(list);")
        js.AppendLine("    colCards.forEach(function(card) { list.appendChild(createCardEl(card, col)); });")
        js.AppendLine("    setupDropZone(list);")
        js.AppendLine("    board.appendChild(colEl);")
        js.AppendLine("  });")
        js.AppendLine("  applyFieldVisibility();")
        js.AppendLine("  applySearchFilter();")
        js.AppendLine("  setupColumnDrag();")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Column drag (unchanged) ---
        js.AppendLine("var colDragSrc=null;")
        js.AppendLine("function setupColumnDrag(){")
        js.AppendLine("  document.querySelectorAll('.column').forEach(function(colEl){")
        js.AppendLine("    var header=colEl.querySelector('.column-header');")
        js.AppendLine("    header.draggable=true;")
        js.AppendLine("    header.addEventListener('dragstart',function(e){ if(e.target.closest('.card')){e.preventDefault();return;} colDragSrc=colEl; colEl.classList.add('col-dragging'); e.dataTransfer.effectAllowed='move'; e.dataTransfer.setData('text/x-column',colEl.dataset.column); });")
        js.AppendLine("    header.addEventListener('dragend',function(){ colDragSrc=null; document.querySelectorAll('.column').forEach(function(c){c.classList.remove('col-dragging','col-drag-over');}); });")
        js.AppendLine("    colEl.addEventListener('dragover',function(e){ if(!colDragSrc||colDragSrc===colEl)return; e.preventDefault(); e.dataTransfer.dropEffect='move'; document.querySelectorAll('.column.col-drag-over').forEach(function(c){c.classList.remove('col-drag-over');}); colEl.classList.add('col-drag-over'); });")
        js.AppendLine("    colEl.addEventListener('dragleave',function(e){ if(!colEl.contains(e.relatedTarget))colEl.classList.remove('col-drag-over'); });")
        js.AppendLine("    colEl.addEventListener('drop',function(e){")
        js.AppendLine("      if(!colDragSrc||colDragSrc===colEl)return; e.preventDefault();")
        js.AppendLine("      var board=document.getElementById('board'); var allCols=[...board.querySelectorAll('.column')];")
        js.AppendLine("      var fromIdx=allCols.indexOf(colDragSrc); var toIdx=allCols.indexOf(colEl);")
        js.AppendLine("      if(fromIdx<0||toIdx<0)return;")
        js.AppendLine("      if(fromIdx<toIdx){board.insertBefore(colDragSrc,colEl.nextSibling);}else{board.insertBefore(colDragSrc,colEl);}")
        js.AppendLine("      document.querySelectorAll('.column').forEach(function(c){c.classList.remove('col-dragging','col-drag-over');});")
        js.AppendLine("      colDragSrc=null;")
        js.AppendLine("      var newOrder=[]; board.querySelectorAll('.column').forEach(function(c){newOrder.push(c.dataset.column);});")
        js.AppendLine("      COLUMNS.sort(function(a,b){return newOrder.indexOf(a.id)-newOrder.indexOf(b.id);});")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({action:'setColumnOrder',columns:newOrder}));")
        js.AppendLine("      showToast('Column order updated');")
        js.AppendLine("    });")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Create card element (with flag support) ---
        js.AppendLine("function createCardEl(card, col) {")
        js.AppendLine("  var el = document.createElement('div');")
        js.AppendLine("  var cls = 'card';")
        js.AppendLine("  if (card.isConversation) cls += ' conversation';")
        js.AppendLine("  if (card.isFlagged) cls += ' flagged';")
        js.AppendLine("  el.className = cls;")
        js.AppendLine("  el.dataset.entryId = card.entryId;")
        js.AppendLine("  el.dataset.id = card.id;")
        js.AppendLine("  el.draggable = true;")
        js.AppendLine()
        js.AppendLine("  var subjectDiv = document.createElement('div');")
        js.AppendLine("  subjectDiv.className = 'card-subject';")
        js.AppendLine("  if (!card.isRead) { var dot = document.createElement('span'); dot.className = 'unread-dot'; subjectDiv.appendChild(dot); }")
        js.AppendLine("  if (card.isFlagged) { var fb = document.createElement('span'); fb.className = 'flag-badge'; fb.textContent = '\u2691 '; subjectDiv.appendChild(fb); }")
        js.AppendLine("  if (card.isConversation) { var ci = document.createElement('span'); ci.className = 'conv-icon'; ci.textContent = '\uD83D\uDCEC '; subjectDiv.appendChild(ci); }")
        js.AppendLine("  subjectDiv.appendChild(document.createTextNode(card.subject));")
        js.AppendLine("  el.appendChild(subjectDiv);")
        js.AppendLine()
        js.AppendLine("  if (card.allCategories && card.allCategories.length > 1) {")
        js.AppendLine("    var catsDiv = document.createElement('div'); catsDiv.className = 'card-categories';")
        js.AppendLine("    card.allCategories.forEach(function(ac) {")
        js.AppendLine("      var badge = document.createElement('span'); badge.className = 'card-cat-badge'; badge.textContent = ac.name; badge.style.background = ac.color;")
        js.AppendLine("      if (ac.name !== col.id) badge.style.opacity = '0.55';")
        js.AppendLine("      catsDiv.appendChild(badge);")
        js.AppendLine("    });")
        js.AppendLine("    el.appendChild(catsDiv);")
        js.AppendLine("  }")
        js.AppendLine()
        js.AppendLine("  var senderDiv = document.createElement('div'); senderDiv.className = 'card-sender';")
        js.AppendLine("  senderDiv.textContent = card.senderName || card.senderEmail;")
        js.AppendLine("  el.appendChild(senderDiv);")
        js.AppendLine()
        js.AppendLine("  var metaDiv = document.createElement('div'); metaDiv.className = 'card-meta';")
        js.AppendLine("  var dateSpan = document.createElement('span'); dateSpan.className = 'card-date'; dateSpan.textContent = card.date; metaDiv.appendChild(dateSpan);")
        js.AppendLine("  var msgSpan = document.createElement('span'); msgSpan.className = 'card-msg-count';")
        js.AppendLine("  if (card.isGrouped && card.messages > 1) { msgSpan.textContent = card.messages + ' mail' + (card.messages !== 1 ? 's' : ''); }")
        js.AppendLine("  metaDiv.appendChild(msgSpan);")
        ' Flag metadata (request + due date)
        js.AppendLine("  if (card.isFlagged || card.isCompleted) {")
        js.AppendLine("    var flagSpan = document.createElement('span');")
        js.AppendLine("    flagSpan.className = 'flag-meta' + (card.isCompleted ? ' completed' : '');")
        js.AppendLine("    var flagText = card.flagRequest || (card.isFlagged ? 'Flagged' : 'Done');")
        js.AppendLine("    if (card.flagDueDate) flagText += ' \u2022 ' + card.flagDueDate;")
        js.AppendLine("    flagSpan.textContent = flagText;")
        js.AppendLine("    metaDiv.appendChild(flagSpan);")
        js.AppendLine("  }")
        js.AppendLine("  el.appendChild(metaDiv);")
        js.AppendLine()
        js.AppendLine("  var summaryDiv = document.createElement('div');")
        js.AppendLine("  summaryDiv.className = card.summary ? 'card-summary' : 'card-summary loading';")
        js.AppendLine("  summaryDiv.dataset.entryId = card.entryId;")
        js.AppendLine("  summaryDiv.textContent = card.summary || 'Generating summary\u2026';")
        js.AppendLine("  el.appendChild(summaryDiv);")
        js.AppendLine()
        js.AppendLine("  var footerDiv = document.createElement('div'); footerDiv.className = 'card-footer';")
        js.AppendLine("  var openBtn = document.createElement('button'); openBtn.className = 'card-btn'; openBtn.dataset.action = 'open'; openBtn.dataset.entryId = card.entryId; openBtn.title = 'Open in Outlook';")
        js.AppendLine("  openBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><path d=""M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z""/><polyline points=""22,6 12,13 2,6""/></svg><span class=""tip"">Open in Outlook</span>';")
        js.AppendLine("  footerDiv.appendChild(openBtn);")
        js.AppendLine("  var unmarkBtn = document.createElement('button'); unmarkBtn.className = 'card-btn'; unmarkBtn.dataset.action = 'unmark'; unmarkBtn.dataset.entryId = card.entryId; unmarkBtn.dataset.column = col.id;")
        js.AppendLine("  unmarkBtn.title = 'Remove from ' + col.title;")
        js.AppendLine("  unmarkBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><line x1=""18"" y1=""6"" x2=""6"" y2=""18""/><line x1=""6"" y1=""6"" x2=""18"" y2=""18""/></svg><span class=""tip"">Remove from ' + escHtml(col.title) + '</span>';")
        js.AppendLine("  footerDiv.appendChild(unmarkBtn);")
        js.AppendLine("  el.appendChild(footerDiv);")
        js.AppendLine()
        js.AppendLine("  el.addEventListener('dragstart', function(e) { isDragging = true; el.classList.add('dragging'); e.dataTransfer.effectAllowed = 'copyMove'; e.dataTransfer.setData('text/plain', card.entryId); updateDragModeVisuals(); });")
        js.AppendLine("  el.addEventListener('dragend', function() { isDragging = false; el.classList.remove('dragging'); clearAllIndicators(); document.getElementById('dragModeLabel').classList.remove('visible'); });")
        js.AppendLine()
        js.AppendLine("  el.querySelectorAll('.card-btn').forEach(function(btn) {")
        js.AppendLine("    btn.addEventListener('click', function(e) {")
        js.AppendLine("      e.stopPropagation(); var act = btn.dataset.action; var eid = btn.dataset.entryId;")
        js.AppendLine("      if (act === 'open') { window.chrome.webview.postMessage(JSON.stringify({ action: 'open', entryId: eid })); }")
        js.AppendLine("      else if (act === 'unmark') { var colId = btn.dataset.column || ''; window.chrome.webview.postMessage(JSON.stringify({ action: 'unmark', entryId: eid, category: colId })); }")
        js.AppendLine("    });")
        js.AppendLine("  });")
        js.AppendLine("  return el;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function escHtml(str) { var d = document.createElement('div'); d.textContent = str; return d.innerHTML; }")
        js.AppendLine()
        ' --- Drag & Drop (unchanged logic) ---
        js.AppendLine("function setupDropZone(list) {")
        js.AppendLine("  list.addEventListener('dragover', function(e) { e.preventDefault(); ctrlHeld = e.ctrlKey; e.dataTransfer.dropEffect = ctrlHeld ? 'copy' : 'move'; var col = list.closest('.column'); col.classList.add('drag-over'); if (ctrlHeld) col.classList.add('drag-over-add'); else col.classList.remove('drag-over-add'); showDropIndicator(list, getDragAfterElement(list, e.clientY)); });")
        js.AppendLine("  list.addEventListener('dragleave', function(e) { if (!list.contains(e.relatedTarget)) { var col = list.closest('.column'); col.classList.remove('drag-over'); col.classList.remove('drag-over-add'); clearIndicators(list); } });")
        js.AppendLine("  list.addEventListener('drop', function(e) {")
        js.AppendLine("    e.preventDefault(); var isAddMode = e.ctrlKey; ctrlHeld = e.ctrlKey;")
        js.AppendLine("    var column = list.closest('.column'); column.classList.remove('drag-over'); column.classList.remove('drag-over-add');")
        js.AppendLine("    var entryId = e.dataTransfer.getData('text/plain');")
        js.AppendLine("    var cardEl = document.querySelector('.card[data-entry-id=""' + CSS.escape(entryId) + '""]');")
        js.AppendLine("    if (!cardEl) return;")
        js.AppendLine("    var oldColumnId = cardEl.closest('.card-list').dataset.column;")
        js.AppendLine("    var newColumnId = list.dataset.column;")
        js.AppendLine("    clearAllIndicators();")
        js.AppendLine("    if (oldColumnId === newColumnId && !isAddMode) { var after = getDragAfterElement(list, e.clientY); if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl); saveState(); return; }")
        js.AppendLine("    if (isAddMode) {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; });")
        js.AppendLine("      if (card && (card.isFlagged || card.isCompleted)) { var targetCol = COLUMNS.find(function(cc) { return cc.id === newColumnId; }); if (targetCol && targetCol.isFlaggedColumn && card.flagBucket === newColumnId) { showToast('Already in ' + newColumnId); return; } }")
        js.AppendLine("      if (card && card.allCategories && card.allCategories.some(function(ac) { return ac.name === newColumnId; })) { showToast('Already in ' + newColumnId); return; }")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'moveAdd', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var addCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('+ Added to ' + (addCol ? addCol.title : newColumnId));")
        js.AppendLine("    } else {")
        js.AppendLine("      var after = getDragAfterElement(list, e.clientY); if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl);")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; }); if (card) card.category = newColumnId;")
        js.AppendLine("      updateCounts(); saveState();")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'move', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var newCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('Moved to ' + (newCol ? newCol.title : newColumnId));")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function getDragAfterElement(list, y) { var els = Array.from(list.querySelectorAll('.card:not(.dragging)')); var closest = null, closestOffset = Number.POSITIVE_INFINITY; els.forEach(function(child) { var box = child.getBoundingClientRect(); var offset = y - box.top - box.height / 2; if (offset < 0 && -offset < closestOffset) { closestOffset = -offset; closest = child; } }); return closest; }")
        js.AppendLine("function showDropIndicator(list, after) { clearIndicators(list); var ind = document.createElement('div'); ind.className = 'drop-indicator visible' + (ctrlHeld ? ' add-mode' : ''); if (after) list.insertBefore(ind, after); else list.appendChild(ind); }")
        js.AppendLine("function clearIndicators(list) { list.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); }); }")
        js.AppendLine("function clearAllIndicators() { document.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); }); document.querySelectorAll('.column.drag-over').forEach(function(el) { el.classList.remove('drag-over'); }); document.querySelectorAll('.column.drag-over-add').forEach(function(el) { el.classList.remove('drag-over-add'); }); }")
        js.AppendLine()
        js.AppendLine("function updateCounts() { document.querySelectorAll('.column').forEach(function(colEl) { var list = colEl.querySelector('.card-list'); if (!list) return; var count = list.querySelectorAll('.card:not(.hidden)').length; var countEl = colEl.querySelector('.count'); if (countEl) countEl.textContent = count; }); }")
        js.AppendLine()
        js.AppendLine("function showToast(message) { var container = document.getElementById('toastContainer'); var toast = document.createElement('div'); toast.className = 'toast'; toast.textContent = message; container.appendChild(toast); setTimeout(function() { toast.classList.add('fade-out'); setTimeout(function() { toast.remove(); }, 300); }, 2500); }")
        js.AppendLine()
        ' --- Search & Filter ---
        js.AppendLine("document.getElementById('searchInput').addEventListener('input', applySearchFilter);")
        js.AppendLine("document.getElementById('columnFilter').addEventListener('change', function() { applySearchFilter(); var val = document.getElementById('columnFilter').value; window.chrome.webview.postMessage(JSON.stringify({ action: 'setColumnFilter', filter: val })); });")
        js.AppendLine()
        js.AppendLine("function applySearchFilter() {")
        js.AppendLine("  var query = document.getElementById('searchInput').value.toLowerCase().trim();")
        js.AppendLine("  var colFilter = document.getElementById('columnFilter').value;")
        js.AppendLine("  document.querySelectorAll('.column').forEach(function(colEl) { var colId = colEl.dataset.column; colEl.style.display = (colFilter === 'all' || colFilter === colId) ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card').forEach(function(cardEl) {")
        js.AppendLine("    var card = cards.find(function(c) { return c.entryId === cardEl.dataset.entryId; }); if (!card) return;")
        js.AppendLine("    if (!query) { cardEl.classList.remove('hidden'); return; }")
        js.AppendLine("    var match = (card.subject || '').toLowerCase().indexOf(query) >= 0 || (card.senderName || '').toLowerCase().indexOf(query) >= 0 || (card.senderEmail || '').toLowerCase().indexOf(query) >= 0;")
        js.AppendLine("    if (match) cardEl.classList.remove('hidden'); else cardEl.classList.add('hidden');")
        js.AppendLine("  });")
        js.AppendLine("  updateCounts();")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Messages from VB.NET ---
        js.AppendLine("if (window.chrome && window.chrome.webview) {")
        js.AppendLine("  window.chrome.webview.addEventListener('message', function(event) {")
        js.AppendLine("    var msg = (typeof event.data === 'string') ? JSON.parse(event.data) : event.data;")
        js.AppendLine("    if (msg.action === 'updateSummary') {")
        js.AppendLine("      document.querySelectorAll('.card-summary').forEach(function(el) { if (el.dataset.entryId === msg.entryId) { if (msg.summary) { el.textContent = msg.summary; el.classList.remove('loading'); } else { el.textContent = 'Generating summary\u2026'; el.classList.add('loading'); } } });")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; }); if (card) card.summary = msg.summary;")
        js.AppendLine("    } else if (msg.action === 'removeCard') {")
        js.AppendLine("      document.querySelectorAll('.card').forEach(function(el) { if (el.dataset.entryId === msg.entryId) el.remove(); });")
        js.AppendLine("      updateCounts(); cards = cards.filter(function(c) { return c.entryId !== msg.entryId; }); showToast('Category removed');")
        js.AppendLine("    } else if (msg.action === 'reloadData') {")
        js.AppendLine("      var d = (typeof msg.data === 'string') ? JSON.parse(msg.data) : msg.data; init(d); showToast('Board reloaded');")
        js.AppendLine("    } else if (msg.action === 'updateCard') {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; });")
        js.AppendLine("      if (card) { card.category = msg.category; card.categoryColor = msg.categoryColor; card.allCategories = msg.allCategories || [];")
        js.AppendLine("        card.isFlagged = msg.isFlagged || false; card.isCompleted = msg.isCompleted || false;")
        js.AppendLine("        card.flagRequest = msg.flagRequest || ''; card.flagDueDate = msg.flagDueDate || '';")
        js.AppendLine("        card.flagBucket = msg.flagBucket || ''; }")
        js.AppendLine("      renderBoard();")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        ' --- Init ---
        js.AppendLine("function startBoard() {")
        js.AppendLine("  try { initTheme(); } catch(e) { console.error('initTheme error:', e); }")
        js.AppendLine("  requestAnimationFrame(function() { requestAnimationFrame(function() { init(INIT_DATA); }); });")
        js.AppendLine("}")
        js.AppendLine("if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', startBoard); } else { startBoard(); }")

        Return js.ToString()
    End Function

#End Region

End Class