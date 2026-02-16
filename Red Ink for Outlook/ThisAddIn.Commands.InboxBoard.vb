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
'    If >50 found, asks user how many to load via dropdown.
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
'    language, and conversation grouping toggle are all persisted via
'    My.Settings (not localStorage).
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

    ''' <summary>Default maximum mails to load without prompting.</summary>
    Private Const InboxBoard_DefaultMax As Integer = 50

    ''' <summary>Number of mails per AI summary batch.</summary>
    Private Const InboxBoard_SummaryBatchSize As Integer = 5

    ''' <summary>Maximum body excerpt length sent for summarization.</summary>
    Private Const InboxBoard_SummaryBodyCap As Integer = 2000

    ''' <summary>Maximum cached summaries to persist in My.Settings.</summary>
    Private Const InboxBoard_SummaryCacheMax As Integer = 500

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
    End Class

    ''' <summary>Represents a column (category) on the board.</summary>
    Private Class InboxBoardColumn
        Public Property CategoryName As String
        Public Property Color As String               ' Hex color
        Public Property Tag As String                 ' Display tag
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
            If columns Is Nothing OrElse columns.Count = 0 Then
                ShowCustomMessageBox("No categories found.", $"{AN} - Inbox Board")
                Return
            End If

            ' 3. Show the board
            ShowInboxBoardForm(mails, columns)

        Catch ex As System.Exception
            ShowCustomMessageBox($"Inbox Board error: {ex.Message}", $"{AN} - Inbox Board")
        End Try
    End Sub

#End Region

#Region "InboxBoard Mail Collection"

    ''' <summary>
    ''' Scans the default Inbox for mails that have at least one category assigned.
    ''' If 50 or more are found, asks the user how many to load.
    ''' </summary>
    Private Function CollectCategorizedInboxMails(Optional maxOverride As Integer = 0) As List(Of InboxBoardEntry)
        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
        Dim inbox As MAPIFolder = ComRetry(Of MAPIFolder)(Function() ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox))

        If inbox Is Nothing Then
            ShowCustomMessageBox("Could not access the Inbox folder.", $"{AN} - Inbox Board")
            Return Nothing
        End If

        Dim folderItems As Outlook.Items = ComRetry(Of Outlook.Items)(Function() inbox.Items)
        Dim totalItems As Integer = ComRetry(Of Integer)(Function() folderItems.Count)

        ' Count categorized mails (manual check to avoid Restrict miscounting)
        Dim totalCategorized As Integer = 0
        For i As Integer = 1 To totalItems
            Try
                Dim idx As Integer = i
                Dim item As Object = ComRetry(Function() folderItems.Item(idx))
                Dim mi As MailItem = TryCast(item, MailItem)
                If mi Is Nothing Then Continue For

                Dim cats As String = ""
                Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
                If Not String.IsNullOrWhiteSpace(cats) Then totalCategorized += 1
            Catch
            End Try
        Next

        If totalCategorized = 0 Then
            ShowCustomMessageBox("No categorized mails found in the Inbox.", $"{AN} - Inbox Board")
            Return Nothing
        End If

        Dim maxToLoad As Integer = If(maxOverride > 0, maxOverride, totalCategorized)

        ' Ask how many if > threshold (and not a reload with a remembered value)
        If maxOverride = 0 AndAlso totalCategorized > InboxBoard_DefaultMax Then
            Dim items As New List(Of SelectionItem)()
            items.Add(New SelectionItem("1000 mails", 1000))
            items.Add(New SelectionItem("2500 mails", 2500))
            items.Add(New SelectionItem("5000 mails", 5000))
            items.Add(New SelectionItem("7500 mails", 7500))
            If totalCategorized > 7500 Then
                items.Add(New SelectionItem($"All ({totalCategorized} mails)", totalCategorized))
            End If

            Dim chosen As Integer = SelectValue(items, 2500,
                $"Found {totalCategorized} categorized mails in Inbox. How many should be loaded?",
                $"{AN} - Inbox Board")
            If chosen = 0 Then Return Nothing
            maxToLoad = chosen
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
                    Dim entry As InboxBoardEntry = BuildInboxBoardEntry(mi, categoryColors)
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
    ''' </summary>
    Private Function BuildInboxBoardEntry(mi As MailItem, categoryColors As Dictionary(Of String, String)) As InboxBoardEntry
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
        ' We no longer derive a count from ConversationIndex because that number (thread depth)
        ' is confusing when shown alongside the grouped message count.
        entry.MessageCount = 1
        entry.IsGrouped = False

        ' ConversationTopic for grouping
        entry.ConversationTopic = ""
        Try : entry.ConversationTopic = ComRetry(Of String)(Function() If(mi.ConversationTopic, "")) : Catch : End Try
        entry.ConversationEntryIDs = New List(Of String)() From {entry.EntryID}

        ' Categories
        Dim cats As String = ""
        Try : cats = ComRetry(Of String)(Function() If(mi.Categories, "")) : Catch : End Try
        entry.AllCategories = cats

        If Not String.IsNullOrEmpty(cats) Then
            Dim firstCat As String = cats.Split({","c, ";"c})(0).Trim()
            entry.Categories = firstCat
            entry.CategoryColor = "#6b7280" ' default gray
            If categoryColors.ContainsKey(firstCat) Then
                entry.CategoryColor = categoryColors(firstCat)
            End If
        Else
            Return Nothing ' No category - should not happen due to filter
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

        entry.Summary = ""
        Return entry
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
                .Tag = catName
            })
        Next

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

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            UpdateMailCategory(id, newCategory)
                        Next

                        ' Update local state
                        entry.Categories = newCategory
                        entry.AllCategories = newCategory
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entry.CategoryColor = If(cc.ContainsKey(newCategory), cc(newCategory), "#6b7280")
                    End If

                    ' Push updated card data back to JS so category badges refresh
                    PushUpdatedCardToJs(entryId, mails, webView)

                Case "moveAdd"
                    ' Add: keep existing categories, append the new one
                    ' For grouped conversations, update ALL mails in the group
                    Dim entryId As String = CStr(msg("entryId"))
                    Dim addCategory As String = CStr(msg("newCategory"))

                    Dim entry = mails.FirstOrDefault(Function(m) m.EntryID = entryId)
                    If entry IsNot Nothing Then
                        ' Update all mails in the conversation group (or just the single mail)
                        Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
                                                                 entry.ConversationEntryIDs,
                                                                 New List(Of String) From {entryId})
                        For Each id In idsToUpdate
                            AddMailCategory(id, addCategory)
                        Next

                        ' Update local state
                        Dim existing As New List(Of String)()
                        If Not String.IsNullOrEmpty(entry.AllCategories) Then
                            For Each part In entry.AllCategories.Split({","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                                Dim t = part.Trim()
                                If Not String.IsNullOrEmpty(t) Then existing.Add(t)
                            Next
                        End If
                        If Not existing.Any(Function(c) c.Equals(addCategory, StringComparison.OrdinalIgnoreCase)) Then
                            existing.Add(addCategory)
                        End If
                        entry.AllCategories = String.Join(", ", existing)
                        entry.Categories = existing(0)
                        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                        Dim ns As Outlook.NameSpace = outlookApp.GetNamespace("MAPI")
                        Dim cc = GetOutlookCategoryColors(ns)
                        entry.CategoryColor = If(cc.ContainsKey(entry.Categories), cc(entry.Categories), "#6b7280")
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

                    ' Determine all entry IDs to process (group or single)
                    Dim idsToUpdate As List(Of String) = If(entry.IsGrouped AndAlso entry.ConversationEntryIDs IsNot Nothing AndAlso entry.ConversationEntryIDs.Count > 1,
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
                        ' Last (or only) category: clear all categories from all mails, remove card
                        For Each id In idsToUpdate
                            ClearMailCategories(id)
                        Next
                        mails.RemoveAll(Function(m) m.EntryID = entryId)

                        ' Tell JS to remove the card
                        Try
                            webView.CoreWebView2.PostWebMessageAsJson(
                                JsonConvert.SerializeObject(New With {.action = "removeCard", .entryId = entryId}))
                        Catch
                        End Try
                    Else
                        ' Multiple categories: remove only the clicked category from all mails
                        For Each id In idsToUpdate
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

                    Dim reloadMax As Integer = If(_inboxBoardLastMaxToLoad > 0, _inboxBoardLastMaxToLoad, 0)
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
                    summaryCts.Cancel()
                    Dim newCts2 As New CancellationTokenSource()
                    summaryCts = newCts2

                    Dim reloadMax2 As Integer = If(_inboxBoardLastMaxToLoad > 0, _inboxBoardLastMaxToLoad, 0)
                    Dim newMails2 = CollectCategorizedInboxMails(reloadMax2)
                    If newMails2 IsNot Nothing Then
                        mails.Clear()
                        mails.AddRange(newMails2)
                        Dim newColumns2 = BuildBoardColumns(mails)
                        columns.Clear()
                        columns.AddRange(newColumns2)
                        Dim dataJson2 As String = BuildBoardDataJson(mails, columns)
                        Try
                            webView.CoreWebView2.PostWebMessageAsJson(
                                JsonConvert.SerializeObject(New With {.action = "reloadData", .data = dataJson2}))
                        Catch
                        End Try
                        Dim capturedMails2 As List(Of InboxBoardEntry) = mails
                        Dim capturedWebView2 As WebView2 = webView
                        Dim capturedToken2 As CancellationToken = newCts2.Token
                        System.Threading.Tasks.Task.Run(Async Function() As System.Threading.Tasks.Task
                                                            Await System.Threading.Tasks.Task.Delay(500)
                                                            frm.BeginInvoke(CType(Async Sub()
                                                                                      Await GenerateSummariesAsync(capturedMails2, capturedWebView2, capturedToken2)
                                                                                  End Sub, System.Windows.Forms.MethodInvoker))
                                                        End Function)
                    End If

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
    ''' Pushes updated card data (allCategories, category, color) to JS after a move/moveAdd.
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
                    .allCategories = allCatsArr
                }))
        Catch
        End Try
    End Sub

#End Region

#Region "InboxBoard Outlook COM Operations"

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
        Try : savedTheme = If(My.Settings.InboxBoardTheme, "") : Catch : End Try
        Try : savedFields = If(My.Settings.InboxBoardFields, "") : Catch : End Try
        Try : savedColumnFilter = If(My.Settings.InboxBoardColumnFilter, "") : Catch : End Try
        Try : savedGroupConversations = My.Settings.InboxBoardGroupConversations : Catch : End Try
        Try : savedSummaryLanguage = If(My.Settings.InboxBoardSummaryLanguage, "") : Catch : End Try

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
.card-subject { font-weight: 600; font-size: 13px; margin-bottom: 4px; color: var(--text); line-height: 1.3;
  overflow: hidden; text-overflow: ellipsis; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
.card-sender { font-size: 11px; color: var(--text-secondary); margin-bottom: 3px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.card-meta { display: flex; align-items: center; gap: 8px; font-size: 10px;
  color: var(--text-muted); margin-bottom: 3px; }
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
        <h3 class=""section-divider"">Summary Language</h3>
        <div class=""lang-input-wrapper"">
          <input type=""text"" id=""summaryLanguageInput"" value=""{escapedLang}""{dataEmpty}>
          <div class=""lang-overlay"">Auto-detect</div>
        </div>
        <h3 class=""section-divider"">Pinned Columns</h3>
{pinnedHtml.ToString().TrimEnd()}
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

    ''' <summary>Returns the JavaScript for the board.</summary>
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
        js.AppendLine("// --- Theme ---")
        js.AppendLine("function applyTheme(theme) {")
        js.AppendLine("  document.documentElement.setAttribute('data-theme', theme);")
        js.AppendLine("  var icon = document.getElementById('themeIcon');")
        js.AppendLine("  if (theme === 'dark') {")
        js.AppendLine("    icon.innerHTML = '<path d=""M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z""/>';")
        js.AppendLine("  } else {")
        js.AppendLine("    icon.innerHTML = '<circle cx=""12"" cy=""12"" r=""5""/>' +")
        js.AppendLine("      '<line x1=""12"" y1=""1"" x2=""12"" y2=""3""/>' +")
        js.AppendLine("      '<line x1=""12"" y1=""21"" x2=""12"" y2=""23""/>' +")
        js.AppendLine("      '<line x1=""4.22"" y1=""4.22"" x2=""5.64"" y2=""5.64""/>' +")
        js.AppendLine("      '<line x1=""18.36"" y1=""18.36"" x2=""19.78"" y2=""19.78""/>' +")
        js.AppendLine("      '<line x1=""1"" y1=""12"" x2=""3"" y2=""12""/>' +")
        js.AppendLine("      '<line x1=""21"" y1=""12"" x2=""23"" y2=""12""/>' +")
        js.AppendLine("      '<line x1=""4.22"" y1=""19.78"" x2=""5.64"" y2=""18.36""/>' +")
        js.AppendLine("      '<line x1=""18.36"" y1=""5.64"" x2=""19.78"" y2=""4.22""/>';")
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
        js.AppendLine("// --- Settings panel ---")
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
        js.AppendLine("// --- Group conversations toggle ---")
        js.AppendLine("document.getElementById('groupConversationsChk').addEventListener('change', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setGroupConversations', enabled: this.checked }));")
        js.AppendLine("});")
        js.AppendLine()
        js.AppendLine("// --- Summary language (free text input with debounce) ---")
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
        js.AppendLine("  langInput.addEventListener('input', function() {")
        js.AppendLine("    updateEmpty();")
        js.AppendLine("    if (langCommitTimer) clearTimeout(langCommitTimer);")
        js.AppendLine("    langCommitTimer = setTimeout(commitLang, 1500);")
        js.AppendLine("  });")
        js.AppendLine("  langInput.addEventListener('blur', function() {")
        js.AppendLine("    if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; }")
        js.AppendLine("    commitLang();")
        js.AppendLine("  });")
        js.AppendLine("  langInput.addEventListener('keydown', function(e) {")
        js.AppendLine("    if (e.key === 'Enter') { e.preventDefault(); if (langCommitTimer) { clearTimeout(langCommitTimer); langCommitTimer = null; } commitLang(); langInput.blur(); }")
        js.AppendLine("  });")
        js.AppendLine("})();")
        js.AppendLine()
        js.AppendLine("// --- Pinned columns ---")
        js.AppendLine("document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(cb) {")
        js.AppendLine("  cb.addEventListener('change', function() {")
        js.AppendLine("    var pinned = [];")
        js.AppendLine("    document.querySelectorAll('#settingsPanel input[type=""checkbox""][data-pinned-cat]').forEach(function(c) {")
        js.AppendLine("      if (c.checked) pinned.push(c.dataset.pinnedCat);")
        js.AppendLine("    });")
        js.AppendLine("    window.chrome.webview.postMessage(JSON.stringify({ action: 'setPinnedColumns', columns: pinned }));")
        js.AppendLine("    showToast('Pinned columns updated');")
        js.AppendLine("  });")
        js.AppendLine("});")
        js.AppendLine()
        js.AppendLine("function applyFieldVisibility() {")
        js.AppendLine("  document.querySelectorAll('.card-sender').forEach(function(el) { el.style.display = fieldSettings.sender ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-date').forEach(function(el) { el.style.display = fieldSettings.date ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-meta .card-msg-count').forEach(function(el) { el.style.display = fieldSettings.messageCount ? '' : 'none'; });")
        js.AppendLine("  document.querySelectorAll('.card-summary').forEach(function(el) { el.style.display = fieldSettings.summary ? '' : 'none'; });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Reload ---")
        js.AppendLine("document.getElementById('reloadBtn').addEventListener('click', function() {")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'reload' }));")
        js.AppendLine("});")
        js.AppendLine()
        js.AppendLine("// --- Ctrl key tracking for drag mode ---")
        js.AppendLine("document.addEventListener('keydown', function(e) {")
        js.AppendLine("  if (e.key === 'Control' && !ctrlHeld) { ctrlHeld = true; updateDragModeVisuals(); }")
        js.AppendLine("});")
        js.AppendLine("document.addEventListener('keyup', function(e) {")
        js.AppendLine("  if (e.key === 'Control') { ctrlHeld = false; updateDragModeVisuals(); }")
        js.AppendLine("});")
        js.AppendLine("window.addEventListener('blur', function() { ctrlHeld = false; updateDragModeVisuals(); });")
        js.AppendLine()
        js.AppendLine("function updateDragModeVisuals() {")
        js.AppendLine("  var label = document.getElementById('dragModeLabel');")
        js.AppendLine("  if (isDragging && ctrlHeld) {")
        js.AppendLine("    label.classList.add('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over').forEach(function(el) {")
        js.AppendLine("      el.classList.add('drag-over-add');")
        js.AppendLine("    });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.visible').forEach(function(el) {")
        js.AppendLine("      el.classList.add('add-mode');")
        js.AppendLine("    });")
        js.AppendLine("  } else {")
        js.AppendLine("    label.classList.remove('visible');")
        js.AppendLine("    document.querySelectorAll('.column.drag-over-add').forEach(function(el) {")
        js.AppendLine("      el.classList.remove('drag-over-add');")
        js.AppendLine("    });")
        js.AppendLine("    document.querySelectorAll('.drop-indicator.add-mode').forEach(function(el) {")
        js.AppendLine("      el.classList.remove('add-mode');")
        js.AppendLine("    });")
        js.AppendLine("  }")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Render Board ---")
        js.AppendLine("function renderBoard() {")
        js.AppendLine("  var board = document.getElementById('board');")
        js.AppendLine("  board.innerHTML = '';")
        js.AppendLine("  COLUMNS.forEach(function(col) {")
        js.AppendLine("    var colEl = document.createElement('div');")
        js.AppendLine("    colEl.className = 'column';")
        js.AppendLine("    colEl.dataset.column = col.id;")
        js.AppendLine("    var colCards = cards.filter(function(c) {")
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
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function createCardEl(card, col) {")
        js.AppendLine("  var el = document.createElement('div');")
        js.AppendLine("  el.className = 'card' + (card.isConversation ? ' conversation' : '');")
        js.AppendLine("  el.dataset.entryId = card.entryId;")
        js.AppendLine("  el.dataset.id = card.id;")
        js.AppendLine("  el.draggable = true;")
        js.AppendLine()
        js.AppendLine("  var subjectDiv = document.createElement('div');")
        js.AppendLine("  subjectDiv.className = 'card-subject';")
        js.AppendLine("  if (!card.isRead) {")
        js.AppendLine("    var dot = document.createElement('span');")
        js.AppendLine("    dot.className = 'unread-dot';")
        js.AppendLine("    subjectDiv.appendChild(dot);")
        js.AppendLine("  }")
        js.AppendLine("  if (card.isConversation) {")
        js.AppendLine("    var convIcon = document.createElement('span');")
        js.AppendLine("    convIcon.className = 'conv-icon';")
        js.AppendLine("    convIcon.textContent = '\uD83D\uDCEC ';")
        js.AppendLine("    subjectDiv.appendChild(convIcon);")
        js.AppendLine("  }")
        js.AppendLine("  subjectDiv.appendChild(document.createTextNode(card.subject));")
        js.AppendLine("  el.appendChild(subjectDiv);")
        js.AppendLine()
        js.AppendLine("  if (card.allCategories && card.allCategories.length > 1) {")
        js.AppendLine("    var catsDiv = document.createElement('div');")
        js.AppendLine("    catsDiv.className = 'card-categories';")
        js.AppendLine("    card.allCategories.forEach(function(ac) {")
        js.AppendLine("      var badge = document.createElement('span');")
        js.AppendLine("      badge.className = 'card-cat-badge';")
        js.AppendLine("      badge.textContent = ac.name;")
        js.AppendLine("      badge.style.background = ac.color;")
        js.AppendLine("      if (ac.name !== col.id) badge.style.opacity = '0.55';")
        js.AppendLine("      catsDiv.appendChild(badge);")
        js.AppendLine("    });")
        js.AppendLine("    el.appendChild(catsDiv);")
        js.AppendLine("  }")
        js.AppendLine()
        js.AppendLine("  var senderDiv = document.createElement('div');")
        js.AppendLine("  senderDiv.className = 'card-sender';")
        js.AppendLine("  senderDiv.textContent = card.senderName || card.senderEmail;")
        js.AppendLine("  el.appendChild(senderDiv);")
        js.AppendLine()
        js.AppendLine("  var metaDiv = document.createElement('div');")
        js.AppendLine("  metaDiv.className = 'card-meta';")
        js.AppendLine("  var dateSpan = document.createElement('span');")
        js.AppendLine("  dateSpan.className = 'card-date';")
        js.AppendLine("  dateSpan.textContent = card.date;")
        js.AppendLine("  metaDiv.appendChild(dateSpan);")
        js.AppendLine("  var msgSpan = document.createElement('span');")
        js.AppendLine("  msgSpan.className = 'card-msg-count';")
        js.AppendLine("  if (card.isGrouped && card.messages > 1) {")
        js.AppendLine("    msgSpan.textContent = card.messages + ' mail' + (card.messages !== 1 ? 's' : '');")
        js.AppendLine("  }")
        js.AppendLine("  metaDiv.appendChild(msgSpan);")
        js.AppendLine("  el.appendChild(metaDiv);")
        js.AppendLine()
        js.AppendLine("  var summaryDiv = document.createElement('div');")
        js.AppendLine("  summaryDiv.className = card.summary ? 'card-summary' : 'card-summary loading';")
        js.AppendLine("  summaryDiv.dataset.entryId = card.entryId;")
        js.AppendLine("  summaryDiv.textContent = card.summary || 'Generating summary\u2026';")
        js.AppendLine("  el.appendChild(summaryDiv);")
        js.AppendLine()
        js.AppendLine("  var footerDiv = document.createElement('div');")
        js.AppendLine("  footerDiv.className = 'card-footer';")
        js.AppendLine()
        js.AppendLine("  var openBtn = document.createElement('button');")
        js.AppendLine("  openBtn.className = 'card-btn';")
        js.AppendLine("  openBtn.dataset.action = 'open';")
        js.AppendLine("  openBtn.dataset.entryId = card.entryId;")
        js.AppendLine("  openBtn.title = 'Open in Outlook';")
        js.AppendLine("  openBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"">' +")
        js.AppendLine("    '<path d=""M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z""/>' +")
        js.AppendLine("    '<polyline points=""22,6 12,13 2,6""/></svg><span class=""tip"">Open in Outlook</span>';")
        js.AppendLine("  footerDiv.appendChild(openBtn);")
        js.AppendLine()
        js.AppendLine("  var unmarkBtn = document.createElement('button');")
        js.AppendLine("  unmarkBtn.className = 'card-btn';")
        js.AppendLine("  unmarkBtn.dataset.action = 'unmark';")
        js.AppendLine("  unmarkBtn.dataset.entryId = card.entryId;")
        js.AppendLine("  unmarkBtn.dataset.column = col.id;")
        js.AppendLine("  unmarkBtn.title = 'Remove from ' + col.title;")
        js.AppendLine("  unmarkBtn.innerHTML = '<svg class=""icon-svg"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round"">' +")
        js.AppendLine("    '<line x1=""18"" y1=""6"" x2=""6"" y2=""18""/><line x1=""6"" y1=""6"" x2=""18"" y2=""18""/></svg><span class=""tip"">Remove from ' + escHtml(col.title) + '</span>';")
        js.AppendLine("  footerDiv.appendChild(unmarkBtn);")
        js.AppendLine("  el.appendChild(footerDiv);")
        js.AppendLine()
        js.AppendLine("  el.addEventListener('dragstart', function(e) {")
        js.AppendLine("    isDragging = true;")
        js.AppendLine("    el.classList.add('dragging');")
        js.AppendLine("    e.dataTransfer.effectAllowed = 'copyMove';")
        js.AppendLine("    e.dataTransfer.setData('text/plain', card.entryId);")
        js.AppendLine("    updateDragModeVisuals();")
        js.AppendLine("  });")
        js.AppendLine("  el.addEventListener('dragend', function() {")
        js.AppendLine("    isDragging = false;")
        js.AppendLine("    el.classList.remove('dragging');")
        js.AppendLine("    clearAllIndicators();")
        js.AppendLine("    document.getElementById('dragModeLabel').classList.remove('visible');")
        js.AppendLine("  });")
        js.AppendLine()
        js.AppendLine("  el.querySelectorAll('.card-btn').forEach(function(btn) {")
        js.AppendLine("    btn.addEventListener('click', function(e) {")
        js.AppendLine("      e.stopPropagation();")
        js.AppendLine("      var act = btn.dataset.action;")
        js.AppendLine("      var eid = btn.dataset.entryId;")
        js.AppendLine("      if (act === 'open') {")
        js.AppendLine("        window.chrome.webview.postMessage(JSON.stringify({ action: 'open', entryId: eid }));")
        js.AppendLine("      } else if (act === 'unmark') {")
        js.AppendLine("        var colId = btn.dataset.column || '';")
        js.AppendLine("        window.chrome.webview.postMessage(JSON.stringify({ action: 'unmark', entryId: eid, category: colId }));")
        js.AppendLine("      }")
        js.AppendLine("    });")
        js.AppendLine("  });")
        js.AppendLine("  return el;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function escHtml(str) { var d = document.createElement('div'); d.textContent = str; return d.innerHTML; }")
        js.AppendLine()
        js.AppendLine("// --- Drag & Drop ---")
        js.AppendLine("function setupDropZone(list) {")
        js.AppendLine("  list.addEventListener('dragover', function(e) {")
        js.AppendLine("    e.preventDefault();")
        js.AppendLine("    ctrlHeld = e.ctrlKey;")
        js.AppendLine("    e.dataTransfer.dropEffect = ctrlHeld ? 'copy' : 'move';")
        js.AppendLine("    var col = list.closest('.column');")
        js.AppendLine("    col.classList.add('drag-over');")
        js.AppendLine("    if (ctrlHeld) col.classList.add('drag-over-add'); else col.classList.remove('drag-over-add');")
        js.AppendLine("    showDropIndicator(list, getDragAfterElement(list, e.clientY));")
        js.AppendLine("  });")
        js.AppendLine("  list.addEventListener('dragleave', function(e) {")
        js.AppendLine("    if (!list.contains(e.relatedTarget)) {")
        js.AppendLine("      var col = list.closest('.column');")
        js.AppendLine("      col.classList.remove('drag-over');")
        js.AppendLine("      col.classList.remove('drag-over-add');")
        js.AppendLine("      clearIndicators(list);")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("  list.addEventListener('drop', function(e) {")
        js.AppendLine("    e.preventDefault();")
        js.AppendLine("    var isAddMode = e.ctrlKey;")
        js.AppendLine("    ctrlHeld = e.ctrlKey;")
        js.AppendLine("    var column = list.closest('.column');")
        js.AppendLine("    column.classList.remove('drag-over');")
        js.AppendLine("    column.classList.remove('drag-over-add');")
        js.AppendLine("    var entryId = e.dataTransfer.getData('text/plain');")
        js.AppendLine("    var cardEl = document.querySelector('.card[data-entry-id=""' + CSS.escape(entryId) + '""]');")
        js.AppendLine("    if (!cardEl) return;")
        js.AppendLine("    var oldColumnId = cardEl.closest('.card-list').dataset.column;")
        js.AppendLine("    var newColumnId = list.dataset.column;")
        js.AppendLine("    clearAllIndicators();")
        js.AppendLine()
        js.AppendLine("    if (oldColumnId === newColumnId && !isAddMode) {")
        js.AppendLine("      var after = getDragAfterElement(list, e.clientY);")
        js.AppendLine("      if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl);")
        js.AppendLine("      saveState();")
        js.AppendLine("      return;")
        js.AppendLine("    }")
        js.AppendLine()
        js.AppendLine("    if (isAddMode) {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; });")
        js.AppendLine("      if (card) {")
        js.AppendLine("        var alreadyHas = card.allCategories && card.allCategories.some(function(ac) { return ac.name === newColumnId; });")
        js.AppendLine("        if (alreadyHas) {")
        js.AppendLine("          showToast('Already in ' + newColumnId);")
        js.AppendLine("          return;")
        js.AppendLine("        }")
        js.AppendLine("      }")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'moveAdd', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var addCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('+ Added to ' + (addCol ? addCol.title : newColumnId));")
        js.AppendLine("    } else {")
        js.AppendLine("      var after = getDragAfterElement(list, e.clientY);")
        js.AppendLine("      if (after) list.insertBefore(cardEl, after); else list.appendChild(cardEl);")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === entryId; });")
        js.AppendLine("      if (card) card.category = newColumnId;")
        js.AppendLine("      updateCounts();")
        js.AppendLine("      saveState();")
        js.AppendLine("      window.chrome.webview.postMessage(JSON.stringify({ action: 'move', entryId: entryId, newCategory: newColumnId }));")
        js.AppendLine("      var newCol = COLUMNS.find(function(c) { return c.id === newColumnId; });")
        js.AppendLine("      showToast('Moved to ' + (newCol ? newCol.title : newColumnId));")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function getDragAfterElement(list, y) {")
        js.AppendLine("  var els = Array.from(list.querySelectorAll('.card:not(.dragging)'));")
        js.AppendLine("  var closest = null, closestOffset = Number.POSITIVE_INFINITY;")
        js.AppendLine("  els.forEach(function(child) {")
        js.AppendLine("    var box = child.getBoundingClientRect();")
        js.AppendLine("    var offset = y - box.top - box.height / 2;")
        js.AppendLine("    if (offset < 0 && -offset < closestOffset) { closestOffset = -offset; closest = child; }")
        js.AppendLine("  });")
        js.AppendLine("  return closest;")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function showDropIndicator(list, after) {")
        js.AppendLine("  clearIndicators(list);")
        js.AppendLine("  var ind = document.createElement('div');")
        js.AppendLine("  ind.className = 'drop-indicator visible' + (ctrlHeld ? ' add-mode' : '');")
        js.AppendLine("  if (after) list.insertBefore(ind, after); else list.appendChild(ind);")
        js.AppendLine("}")
        js.AppendLine("function clearIndicators(list) { list.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); }); }")
        js.AppendLine("function clearAllIndicators() {")
        js.AppendLine("  document.querySelectorAll('.drop-indicator').forEach(function(el) { el.remove(); });")
        js.AppendLine("  document.querySelectorAll('.column.drag-over').forEach(function(el) { el.classList.remove('drag-over'); });")
        js.AppendLine("  document.querySelectorAll('.column.drag-over-add').forEach(function(el) { el.classList.remove('drag-over-add'); });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("function updateCounts() {")
        js.AppendLine("  document.querySelectorAll('.column').forEach(function(colEl) {")
        js.AppendLine("    var list = colEl.querySelector('.card-list');")
        js.AppendLine("    if (!list) return;")
        js.AppendLine("    var count = list.querySelectorAll('.card:not(.hidden)').length;")
        js.AppendLine("    var countEl = colEl.querySelector('.count');")
        js.AppendLine("    if (countEl) countEl.textContent = count;")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Toast ---")
        js.AppendLine("function showToast(message) {")
        js.AppendLine("  var container = document.getElementById('toastContainer');")
        js.AppendLine("  var toast = document.createElement('div'); toast.className = 'toast'; toast.textContent = message;")
        js.AppendLine("  container.appendChild(toast);")
        js.AppendLine("  setTimeout(function() { toast.classList.add('fade-out'); setTimeout(function() { toast.remove(); }, 300); }, 2500);")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Search & Filter ---")
        js.AppendLine("document.getElementById('searchInput').addEventListener('input', applySearchFilter);")
        js.AppendLine("document.getElementById('columnFilter').addEventListener('change', function() {")
        js.AppendLine("  applySearchFilter();")
        js.AppendLine("  var val = document.getElementById('columnFilter').value;")
        js.AppendLine("  window.chrome.webview.postMessage(JSON.stringify({ action: 'setColumnFilter', filter: val }));")
        js.AppendLine("});")
        js.AppendLine()
        js.AppendLine("function applySearchFilter() {")
        js.AppendLine("  var query = document.getElementById('searchInput').value.toLowerCase().trim();")
        js.AppendLine("  var colFilter = document.getElementById('columnFilter').value;")
        js.AppendLine("  document.querySelectorAll('.column').forEach(function(colEl) {")
        js.AppendLine("    var colId = colEl.dataset.column;")
        js.AppendLine("    colEl.style.display = (colFilter === 'all' || colFilter === colId) ? '' : 'none';")
        js.AppendLine("  });")
        js.AppendLine("  document.querySelectorAll('.card').forEach(function(cardEl) {")
        js.AppendLine("    var card = cards.find(function(c) { return c.entryId === cardEl.dataset.entryId; });")
        js.AppendLine("    if (!card) return;")
        js.AppendLine("    if (!query) { cardEl.classList.remove('hidden'); return; }")
        js.AppendLine("    var match = (card.subject || '').toLowerCase().indexOf(query) >= 0 ||")
        js.AppendLine("                (card.senderName || '').toLowerCase().indexOf(query) >= 0 ||")
        js.AppendLine("                (card.senderEmail || '').toLowerCase().indexOf(query) >= 0;")
        js.AppendLine("    if (match) cardEl.classList.remove('hidden'); else cardEl.classList.add('hidden');")
        js.AppendLine("  });")
        js.AppendLine("  updateCounts();")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Messages from VB.NET ---")
        js.AppendLine("if (window.chrome && window.chrome.webview) {")
        js.AppendLine("  window.chrome.webview.addEventListener('message', function(event) {")
        js.AppendLine("    var msg = (typeof event.data === 'string') ? JSON.parse(event.data) : event.data;")
        js.AppendLine("    if (msg.action === 'updateSummary') {")
        js.AppendLine("      var summaryEls = document.querySelectorAll('.card-summary');")
        js.AppendLine("      summaryEls.forEach(function(el) {")
        js.AppendLine("        if (el.dataset.entryId === msg.entryId) {")
        js.AppendLine("          if (msg.summary) { el.textContent = msg.summary; el.classList.remove('loading'); }")
        js.AppendLine("          else { el.textContent = 'Generating summary\u2026'; el.classList.add('loading'); }")
        js.AppendLine("        }")
        js.AppendLine("      });")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; });")
        js.AppendLine("      if (card) card.summary = msg.summary;")
        js.AppendLine("    } else if (msg.action === 'removeCard') {")
        js.AppendLine("      document.querySelectorAll('.card').forEach(function(el) {")
        js.AppendLine("        if (el.dataset.entryId === msg.entryId) el.remove();")
        js.AppendLine("      });")
        js.AppendLine("      updateCounts();")
        js.AppendLine("      cards = cards.filter(function(c) { return c.entryId !== msg.entryId; });")
        js.AppendLine("      showToast('Category removed');")
        js.AppendLine("    } else if (msg.action === 'reloadData') {")
        js.AppendLine("      var d = (typeof msg.data === 'string') ? JSON.parse(msg.data) : msg.data;")
        js.AppendLine("      init(d);")
        js.AppendLine("      showToast('Board reloaded');")
        js.AppendLine("    } else if (msg.action === 'updateCard') {")
        js.AppendLine("      var card = cards.find(function(c) { return c.entryId === msg.entryId; });")
        js.AppendLine("      if (card) {")
        js.AppendLine("        card.category = msg.category;")
        js.AppendLine("        card.categoryColor = msg.categoryColor;")
        js.AppendLine("        card.allCategories = msg.allCategories || [];")
        js.AppendLine("      }")
        js.AppendLine("      renderBoard();")
        js.AppendLine("    }")
        js.AppendLine("  });")
        js.AppendLine("}")
        js.AppendLine()
        js.AppendLine("// --- Init ---")
        js.AppendLine("try { initTheme(); } catch(e) { console.error('initTheme error:', e); }")
        js.AppendLine("init(INIT_DATA);")

        Return js.ToString()
    End Function

#End Region

End Class