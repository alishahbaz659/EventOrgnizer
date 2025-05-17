Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json
Imports Renci.SshNet
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Class Form1
    Private selectedFolderPath As String = ""
    Private logFilePath As String = ""
    Private roomDataApi As String = "https://konferenz.mediaquotes.ch/wp-json/eventin/v2/events"
    Private scheduleDataApi As String = "https://konferenz.mediaquotes.ch/wp-json/eventin/v2/schedules"
    Private speakersDataApi As String = "https://konferenz.mediaquotes.ch/wp-json/eventin/v2/speakers"
    Private sftpHost As String = "danwin.de"            ' SFTP server
    Private sftpPort As Integer = 2244                  ' SFTP port
    Private sftpUser As String = "videokonferenz"       ' SFTP username
    Private sftpPassword As String = "7MinDF2eYSK7U6tuDW$iruWhZe%^Tb$S"  ' SFTP password
    Private remoteDirectory As String = "/srv/videokonferenz"  ' Correct SFTP directory
    Private selectedRoom As String = ""
    Private WithEvents updateTimer As New Timer()
    Private WithEvents countdownTimer As New Timer()
    Private apiUsername As String = "winapp"
    Private apiPassword As String = "me8wRk3bwlCNxuSbvQtGxxb1"
    Private eventsList As New List(Of Dictionary(Of String, Object))
    Private schedulesList As New List(Of Dictionary(Of String, Object))
    Private speakersList As New List(Of Dictionary(Of String, Object))
    Private labelLocation As Label
    Private labelEventLocation As Label
    
    ' Tab control for multiple pages
    Private mainTabControl As TabControl
    Private tabPageMain As TabPage
    Private tabPageLogs As TabPage
    Private tabPageEntries As TabPage
    
    ' DataGridView for showing entries
    Private entriesGridView As DataGridView

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        Try
            ' Set form properties with modern styling
        Me.Text = "Meetings File Sync"
            Me.MinimumSize = New Size(1024, 768)
            Me.BackColor = Color.FromArgb(248, 249, 250)  ' Light gray background
            Me.Font = New Font("Segoe UI", 10.0F)
            Me.Padding = New Padding(15)

            ' Create the tab control with modern styling
            mainTabControl = New TabControl()
            mainTabControl.Dock = DockStyle.Fill
            mainTabControl.Font = New Font("Segoe UI", 11.0F, FontStyle.Regular)
            mainTabControl.Appearance = TabAppearance.Normal
            mainTabControl.SizeMode = TabSizeMode.Fixed
            mainTabControl.ItemSize = New Size(120, 35)
            mainTabControl.Padding = New Point(15, 5)

            ' Style tab pages
            tabPageMain = New TabPage("Dashboard")
            tabPageMain.BackColor = Color.FromArgb(248, 249, 250)
            tabPageMain.Padding = New Padding(20)

            tabPageLogs = New TabPage("System Logs")
            tabPageLogs.BackColor = Color.FromArgb(248, 249, 250)
            tabPageLogs.Padding = New Padding(20)

            tabPageEntries = New TabPage("Events Data")
            tabPageEntries.BackColor = Color.FromArgb(248, 249, 250)
            tabPageEntries.Padding = New Padding(20)
            
            ' Add tab pages to tab control
            mainTabControl.Controls.Add(tabPageMain)
            mainTabControl.Controls.Add(tabPageLogs)
            mainTabControl.Controls.Add(tabPageEntries)
            
            ' Add tab control to form
            Controls.Add(mainTabControl)
            
            ' Create modern status labels
            labelEventLocation = New Label()
            labelEventLocation.Font = New Font("Segoe UI", 10.0F)
            labelEventLocation.ForeColor = Color.FromArgb(73, 80, 87)
            labelEventLocation.AutoSize = True
            labelEventLocation.Text = "?? Location: Not selected"
            
            ' Create responsive TableLayoutPanel for main tab
            Dim mainLayout As New TableLayoutPanel()
            mainLayout.Dock = DockStyle.Fill
            mainLayout.ColumnCount = 1
            mainLayout.RowCount = 7
            mainLayout.Padding = New Padding(10)
            mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
            
            ' Adjust row heights for better spacing
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 120)) ' Logo row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 60))  ' Room selector row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 70))  ' Buttons row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))  ' Status label row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))  ' Location label row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 40))  ' Folder label row
            mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100))  ' Logs row
            
            ' Logo Panel with shadow effect
                Dim logoPanel As New Panel()
                logoPanel.Dock = DockStyle.Fill
            logoPanel.BackColor = Color.White
            logoPanel.Margin = New Padding(0, 0, 0, 20)

            ' Create shadow effect for logo panel using custom paint handler
            AddHandler logoPanel.Paint, Sub(sender As Object, e As PaintEventArgs)
                Dim shadowRect As New Rectangle(0, 0, logoPanel.Width, logoPanel.Height)
                Using shadowBrush As New LinearGradientBrush(shadowRect,
                                                           Color.FromArgb(10, 0, 0, 0),
                                                           Color.Transparent,
                                                           90.0F)
                    e.Graphics.FillRectangle(shadowBrush, shadowRect)
                End Using
            End Sub

            Try
            Dim logoBox As New PictureBox()
                logoBox.Size = New Size(250, 100)
            logoBox.SizeMode = PictureBoxSizeMode.Zoom
                logoBox.Anchor = AnchorStyles.None

                ' Load logo from various possible locations
            Dim logoPath As String = Path.Combine(Application.StartupPath, "Resources", "mch_logo.png")
            If File.Exists(logoPath) Then
                logoBox.Image = Image.FromFile(logoPath)
                ElseIf File.Exists("mch_logo.png") Then
                    logoBox.Image = Image.FromFile("mch_logo.png")
                ElseIf File.Exists(Path.Combine(Application.StartupPath, "mch_logo.png")) Then
                    logoBox.Image = Image.FromFile(Path.Combine(Application.StartupPath, "mch_logo.png"))
                Else
                    LogMessage("Warning: Could not find logo file. The application will continue without displaying the logo.")
            End If
                
                logoPanel.Controls.Add(logoBox)
                logoBox.Location = New Point((logoPanel.Width - logoBox.Width) \ 2, (logoPanel.Height - logoBox.Height) \ 2)
                mainLayout.Controls.Add(logoPanel, 0, 0)
        Catch ex As Exception
            LogMessage("Error loading logo: " & ex.Message)
        End Try

            ' Modern Room Selector
            Dim roomPanel As New Panel()
            roomPanel.Dock = DockStyle.Fill
            
        comboBoxRooms.DropDownStyle = ComboBoxStyle.DropDownList
            comboBoxRooms.Font = New Font("Segoe UI", 11.0F)
            comboBoxRooms.Size = New Size(450, 35)
        comboBoxRooms.BackColor = Color.White
            comboBoxRooms.FlatStyle = FlatStyle.Flat
            comboBoxRooms.Dock = DockStyle.None
            comboBoxRooms.Anchor = AnchorStyles.None
            
            roomPanel.Controls.Add(comboBoxRooms)
            mainLayout.Controls.Add(roomPanel, 0, 1)
            
            ' Modern Button Panel
            Dim buttonPanel As New FlowLayoutPanel()
            buttonPanel.Dock = DockStyle.Fill
            buttonPanel.FlowDirection = FlowDirection.LeftToRight
            buttonPanel.WrapContents = False
            buttonPanel.AutoSize = True
            buttonPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink
            buttonPanel.Anchor = AnchorStyles.None

            ' Modern button styling
        Dim buttonStyle = Sub(btn As Button)
                              btn.FlatStyle = FlatStyle.Flat
                                btn.BackColor = ColorTranslator.FromHtml("#dc3545")  ' Bootstrap danger red
                              btn.ForeColor = Color.White
                                btn.Font = New Font("Segoe UI", 11.0F)
                                btn.Size = New Size(200, 40)
                              btn.Cursor = Cursors.Hand
                                btn.FlatAppearance.BorderSize = 0
                                  btn.Margin = New Padding(10, 0, 10, 0)
                                AddHandler btn.MouseEnter, Sub() btn.BackColor = ControlPaint.Dark(ColorTranslator.FromHtml("#dc3545"), 0.1)
                                AddHandler btn.MouseLeave, Sub() btn.BackColor = ColorTranslator.FromHtml("#dc3545")
                          End Sub

            ' Apply modern button styles
        buttonStyle(buttonConnect)
        buttonStyle(buttonSelectFolder)
        buttonStyle(buttonSyncFiles)

            ' Add buttons to panel
            buttonPanel.Controls.Add(buttonConnect)
            buttonPanel.Controls.Add(buttonSelectFolder)
            buttonPanel.Controls.Add(buttonSyncFiles)
            
            mainLayout.Controls.Add(buttonPanel, 0, 2)
            
            ' Modern status labels with icons
        labelStatus.Font = New Font("Segoe UI", 10.0F)
            labelStatus.ForeColor = Color.FromArgb(73, 80, 87)
        labelStatus.AutoSize = True
            labelStatus.Text = "?? Status: Not connected"
            
            Dim labelsPanel As New Panel()
            labelsPanel.Dock = DockStyle.Fill
            labelsPanel.Padding = New Padding(20, 5, 0, 0)
            labelsPanel.Controls.Add(labelStatus)
            mainLayout.Controls.Add(labelsPanel, 0, 3)
            
            ' Location label
            Dim locationPanel As New Panel()
            locationPanel.Dock = DockStyle.Fill
            locationPanel.Padding = New Padding(20, 5, 0, 0)
            labelEventLocation.Location = New Point(0, 0)
            locationPanel.Controls.Add(labelEventLocation)
            mainLayout.Controls.Add(locationPanel, 0, 4)
            
            ' Folder label
            Dim folderPanel As New Panel()
            folderPanel.Dock = DockStyle.Fill
            folderPanel.Padding = New Padding(20, 5, 0, 0)
        labelFolder.Font = New Font("Segoe UI", 10.0F)
            labelFolder.ForeColor = Color.FromArgb(73, 80, 87)
        labelFolder.AutoSize = True
            labelFolder.Text = "?? Storage Location: Not selected"
            folderPanel.Controls.Add(labelFolder)
            mainLayout.Controls.Add(folderPanel, 0, 5)
            
            ' Modern logs panel
            Dim logsPanel As New Panel()
            logsPanel.Dock = DockStyle.Fill
            logsPanel.Padding = New Padding(20, 10, 20, 10)

            ' Add shadow effect to logs panel using custom paint handler
            AddHandler logsPanel.Paint, Sub(sender As Object, e As PaintEventArgs)
                Dim shadowRect As New Rectangle(0, 0, logsPanel.Width, logsPanel.Height)
                Using shadowBrush As New LinearGradientBrush(shadowRect,
                                                           Color.FromArgb(10, 0, 0, 0),
                                                           Color.Transparent,
                                                           90.0F)
                    e.Graphics.FillRectangle(shadowBrush, shadowRect)
                End Using
            End Sub

            textBoxLogs.Font = New Font("Consolas", 10.0F)
            textBoxLogs.Dock = DockStyle.Fill
        textBoxLogs.BackColor = Color.White
            textBoxLogs.BorderStyle = BorderStyle.None
        textBoxLogs.Multiline = True
        textBoxLogs.ScrollBars = ScrollBars.Vertical
        textBoxLogs.ReadOnly = True

            logsPanel.Controls.Add(textBoxLogs)
            mainLayout.Controls.Add(logsPanel, 0, 6)
            
            ' Add the main layout to the main tab
            tabPageMain.Controls.Add(mainLayout)
            
            ' Setup full-size logs tab
            Dim fullLogsPanel As New Panel()
            fullLogsPanel.Dock = DockStyle.Fill
            fullLogsPanel.Padding = New Padding(15)
            
            Dim fullLogsTextBox As New TextBox()
            fullLogsTextBox.Font = New Font("Consolas", 10.0F)
            fullLogsTextBox.Dock = DockStyle.Fill
            fullLogsTextBox.BackColor = Color.White
            fullLogsTextBox.BorderStyle = BorderStyle.None
            fullLogsTextBox.Multiline = True
            fullLogsTextBox.ScrollBars = ScrollBars.Both
            fullLogsTextBox.ReadOnly = True
            fullLogsTextBox.WordWrap = False
            fullLogsTextBox.Name = "textBoxFullLogs"
            
            fullLogsPanel.Controls.Add(fullLogsTextBox)
            tabPageLogs.Controls.Add(fullLogsPanel)
            
            ' Setup entries tab with DataGridView and SplitContainer
            Dim entriesSplitContainer As New SplitContainer()
            entriesSplitContainer.Dock = DockStyle.Fill
            entriesSplitContainer.Orientation = Orientation.Vertical
            entriesSplitContainer.SplitterDistance = 250
            entriesSplitContainer.Panel1MinSize = 150
            entriesSplitContainer.Panel2MinSize = 150
            
            ' Initialize entriesGridView
            entriesGridView = New DataGridView()
            entriesGridView.Dock = DockStyle.Fill
            entriesGridView.BackgroundColor = Color.White
            entriesGridView.BorderStyle = BorderStyle.None
            entriesGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            entriesGridView.RowHeadersVisible = False
            entriesGridView.AllowUserToAddRows = False
            entriesGridView.AllowUserToDeleteRows = False
            entriesGridView.ReadOnly = True
            entriesGridView.Font = New Font("Segoe UI", 9.0F)
            entriesGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            
            ' Add columns for entries grid
            entriesGridView.Columns.Add("EventId", "ID")
            entriesGridView.Columns.Add("Title", "Title")
            entriesGridView.Columns.Add("EventDate", "Date")
            entriesGridView.Columns.Add("StartTime", "Start Time")
            entriesGridView.Columns.Add("EndTime", "End Time")
            entriesGridView.Columns.Add("Room", "Room")
            entriesGridView.Columns.Add("Speakers", "Speakers")
            entriesGridView.Columns.Add("Countdown", "Countdown")
            
            ' Hide the EventId column as it's just for reference
            entriesGridView.Columns("EventId").Visible = False
            
            ' Add header to top panel
            Dim scheduleHeaderPanel As New Panel()
            scheduleHeaderPanel.Dock = DockStyle.Top
            scheduleHeaderPanel.Height = 30
            scheduleHeaderPanel.BackColor = ColorTranslator.FromHtml("#9c1c13")
            
            Dim scheduleHeaderLabel As New Label()
            scheduleHeaderLabel.Text = "Event Schedules"
            scheduleHeaderLabel.ForeColor = Color.White
            scheduleHeaderLabel.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
            scheduleHeaderLabel.AutoSize = False
            scheduleHeaderLabel.TextAlign = ContentAlignment.MiddleLeft
            scheduleHeaderLabel.Dock = DockStyle.Fill
            scheduleHeaderLabel.Padding = New Padding(10, 0, 0, 0)
            
            scheduleHeaderPanel.Controls.Add(scheduleHeaderLabel)
            
            ' Create container for header and grid
            Dim schedulesPanel As New Panel()
            schedulesPanel.Dock = DockStyle.Fill
            schedulesPanel.Controls.Add(entriesGridView)
            schedulesPanel.Controls.Add(scheduleHeaderPanel)
            
            entriesSplitContainer.Panel1.Controls.Add(schedulesPanel)
            
            ' Bottom panel - Details view
            Dim detailPanel As New Panel()
            detailPanel.Dock = DockStyle.Fill
            detailPanel.Padding = New Padding(10)
            detailPanel.BackColor = Color.FromArgb(245, 245, 245)
            
            ' Create header for details panel
            Dim detailHeaderPanel As New Panel()
            detailHeaderPanel.Dock = DockStyle.Top
            detailHeaderPanel.Height = 30
            detailHeaderPanel.BackColor = ColorTranslator.FromHtml("#9c1c13")
            
            Dim detailHeaderLabel As New Label()
            detailHeaderLabel.Text = "Event Details"
            detailHeaderLabel.ForeColor = Color.White
            detailHeaderLabel.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
            detailHeaderLabel.AutoSize = False
            detailHeaderLabel.TextAlign = ContentAlignment.MiddleLeft
            detailHeaderLabel.Dock = DockStyle.Fill
            detailHeaderLabel.Padding = New Padding(10, 0, 0, 0)
            
            detailHeaderPanel.Controls.Add(detailHeaderLabel)
            
            ' Create detail content panel
            Dim detailContentPanel As New Panel()
            detailContentPanel.Dock = DockStyle.Fill
            detailContentPanel.Padding = New Padding(10)
            
            ' Create labels for details
            Dim detailTitleLabel As New Label()
            detailTitleLabel.Text = "No event selected"
            detailTitleLabel.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
            detailTitleLabel.ForeColor = Color.FromArgb(50, 50, 50)
            detailTitleLabel.Location = New Point(10, 10)
            detailTitleLabel.AutoSize = True
            detailTitleLabel.Name = "detailTitleLabel"
            
            Dim detailTimeLabel As New Label()
            detailTimeLabel.Text = "Time: -"
            detailTimeLabel.Font = New Font("Segoe UI", 10.0F)
            detailTimeLabel.ForeColor = Color.FromArgb(80, 80, 80)
            detailTimeLabel.Location = New Point(10, 40)
            detailTimeLabel.AutoSize = True
            detailTimeLabel.Name = "detailTimeLabel"
            
            Dim detailRoomLabel As New Label()
            detailRoomLabel.Text = "Room: -"
            detailRoomLabel.Font = New Font("Segoe UI", 10.0F)
            detailRoomLabel.ForeColor = Color.FromArgb(80, 80, 80)
            detailRoomLabel.Location = New Point(10, 65)
            detailRoomLabel.AutoSize = True
            detailRoomLabel.Name = "detailRoomLabel"
            
            Dim detailSpeakersLabel As New Label()
            detailSpeakersLabel.Text = "Speakers: -"
            detailSpeakersLabel.Font = New Font("Segoe UI", 10.0F)
            detailSpeakersLabel.ForeColor = Color.FromArgb(80, 80, 80)
            detailSpeakersLabel.Location = New Point(10, 90)
            detailSpeakersLabel.AutoSize = True
            detailSpeakersLabel.Name = "detailSpeakersLabel"
            
            Dim detailDescriptionLabel As New Label()
            detailDescriptionLabel.Text = "Description:"
            detailDescriptionLabel.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
            detailDescriptionLabel.ForeColor = Color.FromArgb(80, 80, 80)
            detailDescriptionLabel.Location = New Point(10, 125)
            detailDescriptionLabel.AutoSize = True
            
            Dim detailDescriptionText As New TextBox()
            detailDescriptionText.Multiline = True
            detailDescriptionText.ReadOnly = True
            detailDescriptionText.BorderStyle = BorderStyle.FixedSingle
            detailDescriptionText.BackColor = Color.White
            detailDescriptionText.Font = New Font("Segoe UI", 9.0F)
            detailDescriptionText.Location = New Point(10, 150)
            detailDescriptionText.Size = New Size(detailContentPanel.Width - 20, 100)
            detailDescriptionText.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            detailDescriptionText.Name = "detailDescriptionText"
            
            ' Add current time display
            Dim currentTimeLabel As New Label()
            currentTimeLabel.Text = "Current Time: " & DateTime.Now.ToString("HH:mm:ss")
            currentTimeLabel.Font = New Font("Segoe UI", 9.0F)
            currentTimeLabel.ForeColor = Color.FromArgb(100, 100, 100)
            currentTimeLabel.Location = New Point(10, detailContentPanel.Height - 30)
            currentTimeLabel.AutoSize = True
            currentTimeLabel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
            currentTimeLabel.Name = "currentTimeLabel"
            
            ' Create a timer to update the current time
            Dim clockTimer As New Timer()
            clockTimer.Interval = 1000 ' 1 second
            AddHandler clockTimer.Tick, Sub(sender, e)
                If currentTimeLabel IsNot Nothing Then
                    currentTimeLabel.Text = "Current Time: " & DateTime.Now.ToString("HH:mm:ss")
                End If
            End Sub
            clockTimer.Start()
            
            ' Add controls to detail content panel
            detailContentPanel.Controls.Add(detailTitleLabel)
            detailContentPanel.Controls.Add(detailTimeLabel)
            detailContentPanel.Controls.Add(detailRoomLabel)
            detailContentPanel.Controls.Add(detailSpeakersLabel)
            detailContentPanel.Controls.Add(detailDescriptionLabel)
            detailContentPanel.Controls.Add(detailDescriptionText)
            detailContentPanel.Controls.Add(currentTimeLabel)
            
            detailPanel.Controls.Add(detailContentPanel)
            detailPanel.Controls.Add(detailHeaderPanel)
            
            entriesSplitContainer.Panel2.Controls.Add(detailPanel)
            
            ' Add event handler for grid selection changed
            AddHandler entriesGridView.SelectionChanged, AddressOf EntriesGridView_SelectionChanged
            
            ' Add the split container to entries tab
            Dim entriesPanel As New Panel()
            entriesPanel.Dock = DockStyle.Fill
            entriesPanel.Padding = New Padding(10)
            entriesPanel.Controls.Add(entriesSplitContainer)
            
            tabPageEntries.Controls.Add(entriesPanel)
            
            ' Set up timer for auto-sync
            updateTimer.Interval = 60000  ' 60 seconds
            
            ' Set up countdown timer
            countdownTimer.Interval = 1000  ' 1 second
            countdownTimer.Start()            
        Catch ex As Exception
            LogMessage("Error during form initialization: " & ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            LogMessage("Application started.")
            labelStatus.Text = "Status: Not connected"
            
            ' Make sure all panels and controls are sized properly on first load
            Me.PerformLayout()
            
            ' Center the combobox if visible
            Try
                Dim roomPanel = tabPageMain.Controls(0).Controls(1)
                If roomPanel IsNot Nothing AndAlso comboBoxRooms IsNot Nothing Then
                    comboBoxRooms.Location = New Point((roomPanel.Width - comboBoxRooms.Width) \ 2, (roomPanel.Height - comboBoxRooms.Height) \ 2)
                End If
                
                ' Center logo
                Dim logoPanel = tabPageMain.Controls(0).Controls(0)
                If logoPanel IsNot Nothing AndAlso logoPanel.Controls.Count > 0 Then
                    Dim logoBox = DirectCast(logoPanel.Controls(0), PictureBox)
                    logoBox.Location = New Point((logoPanel.Width - logoBox.Width) \ 2, (logoPanel.Height - logoBox.Height) \ 2)
                End If
                
                ' Center buttons
                Dim buttonPanel = tabPageMain.Controls(0).Controls(2)
                If TypeOf buttonPanel Is FlowLayoutPanel Then
                    Dim flow = DirectCast(buttonPanel, FlowLayoutPanel)
                    flow.Padding = New Padding((flow.Width - flow.Controls.Cast(Of Control).Sum(Function(c) c.Width) - (flow.Controls.Count * 20)) \ 2, 5, 0, 0)
                End If
            Catch ex As Exception
                LogMessage("Layout adjustment warning: " & ex.Message)
            End Try
            
            ' Log application version info
            LogMessage("Application version: 1.1.0")
            LogMessage("Current working directory: " & Application.StartupPath)
            
            ' Set focus to connect button
            buttonConnect.Focus()
        Catch ex As Exception
            LogMessage("Error during form load: " & ex.Message)
        End Try
    End Sub
    


    Private Sub ButtonConnect_Click(sender As Object, e As EventArgs) Handles buttonConnect.Click
        labelStatus.Text = "Status: Connecting..."
        LogMessage("Connecting...")
        FetchRoomData()
    End Sub

    Private Async Sub FetchRoomData()
        Using client As New HttpClient()
            Try
                ' Add Basic Authentication
                Dim byteArray As Byte() = Encoding.ASCII.GetBytes($"{apiUsername}:{apiPassword}")
                client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(byteArray))

                Dim response As HttpResponseMessage = Await client.GetAsync(roomDataApi)
                response.EnsureSuccessStatusCode()

                Dim result As String = Await response.Content.ReadAsStringAsync()
                ' Parse the JSON response
                Dim jsonResponse = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(result)
                eventsList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(jsonResponse("items").ToString())

                comboBoxRooms.Items.Clear()
                For Each eventItem In eventsList
                    ' Add only the title to ComboBox
                    comboBoxRooms.Items.Add(eventItem("title").ToString())
                Next

                labelStatus.Text = "Status: Connected"
                LogMessage("Event data successfully loaded.")
            Catch ex As Exception
                labelStatus.Text = "Status: Connection failed"
                LogMessage("Error: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub ButtonSelectFolder_Click(sender As Object, e As EventArgs) Handles buttonSelectFolder.Click
        Using folderDialog As New FolderBrowserDialog()
            folderDialog.Description = "Select a folder to store meeting files"
            
            ' Set initial directory if a folder was previously selected
            If Not String.IsNullOrEmpty(selectedFolderPath) AndAlso Directory.Exists(selectedFolderPath) Then
                folderDialog.SelectedPath = selectedFolderPath
            End If
            
            If folderDialog.ShowDialog() = DialogResult.OK Then
                selectedFolderPath = folderDialog.SelectedPath
                
                ' Update label with truncated path if very long
                Dim displayPath = selectedFolderPath
                If displayPath.Length > 40 Then
                    displayPath = "..." + displayPath.Substring(displayPath.Length - 37)
                End If
                
                labelFolder.Text = "Storage: " & displayPath
                
                ' Create log file path
                logFilePath = Path.Combine(selectedFolderPath, "MeetingsSync.log")
                
                ' Create the log file or append to it if it exists
                Try
                    If Not File.Exists(logFilePath) Then
                        File.WriteAllText(logFilePath, $"--- Log started: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---{Environment.NewLine}")
                    Else
                        File.AppendAllText(logFilePath, $"{Environment.NewLine}--- Log continued: {DateTime.Now:yyyy-MM-dd HH:mm:ss} ---{Environment.NewLine}")
                    End If
                    
                    LogMessage("Folder selected: " & selectedFolderPath)
                    
                    ' Create a subfolder for meeting files
                    Dim conferenceFolder = Path.Combine(selectedFolderPath, "MeetingFiles")
                    If Not Directory.Exists(conferenceFolder) Then
                        Directory.CreateDirectory(conferenceFolder)
                        LogMessage("Created subfolder: MeetingFiles")
                    End If
                    
                    ' If we have a selected room, trigger a sync
                    If Not String.IsNullOrEmpty(selectedRoom) Then
                        LogMessage("Triggering sync for: " & selectedRoom)
                        SyncFiles()
                    End If
                    
                Catch ex As Exception
                    LogMessage("Error setting up log file: " & ex.Message)
                End Try
            End If
        End Using
    End Sub

    Private Sub ButtonSyncFiles_Click(sender As Object, e As EventArgs) Handles buttonSyncFiles.Click
        LogMessage("Starting SFTP synchronization...")
        SyncFiles()
    End Sub

    Private Sub SyncFiles()
        If String.IsNullOrEmpty(selectedFolderPath) OrElse String.IsNullOrEmpty(selectedRoom) Then
            LogMessage("Error: No folder selected or no room selected.")
            Return
        End If

        ' Get the location for the selected event
        Dim selectedEvent = eventsList.Find(Function(evt) evt("title").ToString() = selectedRoom)
        Dim eventLocation As String = ""

        ' Extract location from selected event
        If selectedEvent IsNot Nothing AndAlso
           selectedEvent.ContainsKey("location") AndAlso
           TypeOf selectedEvent("location") Is Newtonsoft.Json.Linq.JObject Then

            Dim locationObj = DirectCast(selectedEvent("location"), Newtonsoft.Json.Linq.JObject)
            If locationObj.ContainsKey("address") Then
                eventLocation = locationObj("address").ToString().Trim()
            End If
        End If

        If String.IsNullOrEmpty(eventLocation) Then
            LogMessage("Error: No location found for selected event.")
            Return
        End If

        Try
            Using sftp = New SftpClient(sftpHost, sftpPort, sftpUser, sftpPassword)
                sftp.Connect()
                LogMessage("SFTP connection established.")

                ' List and log all files in the directory
                Dim remoteFiles = sftp.ListDirectory(remoteDirectory)
                LogMessage("--- All Files on SFTP Server ---")
                For Each file In remoteFiles
                    If Not file.IsDirectory AndAlso Not file.IsSymbolicLink Then
                        LogMessage($"Found file: {file.Name} (Last modified: {file.LastWriteTime})")
                    End If
                Next
                LogMessage("--- End of File List ---")

                ' Filter files by location name instead of event title
                Dim locationFiles = remoteFiles.Where(Function(f)
                                                          Return Not f.IsDirectory AndAlso
                           Not f.IsSymbolicLink AndAlso
                           f.Name.Contains(eventLocation, StringComparison.OrdinalIgnoreCase)
                                                      End Function).ToList()

                LogMessage($"Found {locationFiles.Count} files matching location: {eventLocation}")

                Dim downloadedFiles As Integer = 0
                Dim errors As Integer = 0

                For Each file In locationFiles
                    ' Sanitize the file name before creating the local path
                    Dim sanitizedFileName = SanitizeFileName(file.Name)
                    Dim localFilePath = Path.Combine(selectedFolderPath, sanitizedFileName)

                    ' Check if file exists and get its timestamp
                    Dim shouldDownload = True
                    If System.IO.File.Exists(localFilePath) Then
                        Dim localFile = New FileInfo(localFilePath)
                        ' Download only if remote file is newer
                        shouldDownload = file.LastWriteTime > localFile.LastWriteTime
                    End If

                    If shouldDownload Then
                        Try
                            Using fileStream As New FileStream(localFilePath, FileMode.Create)
                                sftp.DownloadFile(file.FullName, fileStream)
                            End Using
                            LogMessage($"File downloaded/updated: {sanitizedFileName}")
                            downloadedFiles += 1
                        Catch ex As Exception
                            LogMessage($"Error downloading file {file.Name}: {ex.Message}")
                            errors += 1
                        End Try
                    End If
                Next

                sftp.Disconnect()
                LogMessage($"SFTP sync completed. Downloaded/updated {downloadedFiles} files with {errors} errors.")
                
                ' Update the entries tab after sync
                If mainTabControl.SelectedTab Is tabPageMain Then
                    FetchSchedulesData()
                End If
            End Using
        Catch ex As Exception
            LogMessage("SFTP sync error: " & ex.Message)
        End Try
    End Sub

    Private Function SanitizeFileName(fileName As String) As String
        ' Replace invalid characters with safe alternatives
        Dim invalidChars As Char() = Path.GetInvalidFileNameChars()
        Dim result As String = fileName

        For Each invalidChar In invalidChars
            result = result.Replace(invalidChar, "-"c)
        Next

        ' Replace multiple consecutive dashes with a single dash
        While result.Contains("--")
            result = result.Replace("--", "-")
        End While

        Return result.Trim("-"c)
    End Function

    Private Sub LogMessage(message As String)
        Dim timestamp As String = DateTime.Now.ToString("HH:mm:ss")
        Dim logMessage = $"{timestamp} � {message}{Environment.NewLine}"
        
        ' Update main tab log
        textBoxLogs.AppendText(logMessage)
        textBoxLogs.SelectionStart = textBoxLogs.TextLength
        textBoxLogs.ScrollToCaret()
        
        ' Update full logs tab if it exists
        Try
            Dim fullLogsTextBox = DirectCast(tabPageLogs.Controls(0).Controls(0), TextBox)
            If fullLogsTextBox IsNot Nothing Then
                fullLogsTextBox.AppendText(logMessage)
                fullLogsTextBox.SelectionStart = fullLogsTextBox.TextLength
                fullLogsTextBox.ScrollToCaret()
            End If
        Catch ex As Exception
            ' Silently ignore if the control doesn't exist yet
        End Try

        If Not String.IsNullOrEmpty(logFilePath) Then
            File.AppendAllText(logFilePath, logMessage)
        End If
    End Sub

    Private Sub ComboBoxRooms_SelectedIndexChanged(sender As Object, e As EventArgs) Handles comboBoxRooms.SelectedIndexChanged
        If comboBoxRooms.SelectedItem IsNot Nothing Then
            selectedRoom = comboBoxRooms.SelectedItem.ToString()

            ' Find the selected event in our stored events list
            Dim selectedEvent = eventsList.Find(Function(evt) evt("title").ToString() = selectedRoom)

            ' Update location labels if location exists
            If selectedEvent IsNot Nothing AndAlso
               selectedEvent.ContainsKey("location") AndAlso
               TypeOf selectedEvent("location") Is Newtonsoft.Json.Linq.JObject Then

                Dim locationObj = DirectCast(selectedEvent("location"), Newtonsoft.Json.Linq.JObject)
                If locationObj.ContainsKey("address") Then
                    Dim locationText = locationObj("address").ToString()
                    labelEventLocation.Text = "Location: " & locationText
                    labelFolder.Text = "Storage: " & If(String.IsNullOrEmpty(selectedFolderPath), locationText, selectedFolderPath)
                Else
                    labelEventLocation.Text = "Location: Not specified"
                    labelFolder.Text = "Storage: " & If(String.IsNullOrEmpty(selectedFolderPath), "Not specified", selectedFolderPath)
                End If
            Else
                labelEventLocation.Text = "Location: Not specified"
                labelFolder.Text = "Storage: " & If(String.IsNullOrEmpty(selectedFolderPath), "Not specified", selectedFolderPath)
            End If

            LogMessage($"Selected event: {selectedRoom}")
            
            ' Ensure entriesGridView is initialized
            EnsureEntriesGridViewInitialized()
            
            ' Fetch schedule data for the selected event
            FetchSchedulesData()
            
            ' Start auto-update when room is selected
            updateTimer.Start()
        End If
    End Sub

    ' Method to ensure entriesGridView is initialized
    Private Sub EnsureEntriesGridViewInitialized()
        ' Check if entriesGridView is null or not properly initialized
        If entriesGridView Is Nothing Then
            LogMessage("Reinitializing entriesGridView...")
            
            ' Create new instance
            entriesGridView = New DataGridView()
            entriesGridView.Name = "entriesGridView"  ' Add a name to make it easier to find
            entriesGridView.Dock = DockStyle.Fill
            entriesGridView.BackgroundColor = Color.White
            entriesGridView.BorderStyle = BorderStyle.None
            entriesGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            entriesGridView.RowHeadersVisible = False
            entriesGridView.AllowUserToAddRows = False
            entriesGridView.AllowUserToDeleteRows = False
            entriesGridView.ReadOnly = True
            entriesGridView.Font = New Font("Segoe UI", 9.0F)
            entriesGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            
            ' Add columns for entries grid
            entriesGridView.Columns.Add("EventId", "ID")
            entriesGridView.Columns.Add("Title", "Title")
            entriesGridView.Columns.Add("EventDate", "Date")
            entriesGridView.Columns.Add("StartTime", "Start Time")
            entriesGridView.Columns.Add("EndTime", "End Time")
            entriesGridView.Columns.Add("Room", "Room")
            entriesGridView.Columns.Add("Speakers", "Speakers")
            entriesGridView.Columns.Add("Countdown", "Countdown")
            
            ' Hide the EventId column as it's just for reference
            entriesGridView.Columns("EventId").Visible = False
            
            ' Add event handler for grid selection changed
            AddHandler entriesGridView.SelectionChanged, AddressOf EntriesGridView_SelectionChanged
            
            ' Find the container for the grid and add it if it exists
            If tabPageEntries IsNot Nothing AndAlso tabPageEntries.Controls.Count > 0 Then
                LogMessage($"tabPageEntries has {tabPageEntries.Controls.Count} controls")
                Dim entriesPanel = tabPageEntries.Controls(0)
                If entriesPanel IsNot Nothing AndAlso entriesPanel.Controls.Count > 0 Then
                    LogMessage($"entriesPanel has {entriesPanel.Controls.Count} controls")
                    Dim splitContainer = TryCast(entriesPanel.Controls(0), SplitContainer)
                    If splitContainer IsNot Nothing Then
                        LogMessage("Found SplitContainer in entriesPanel")
                        If splitContainer.Panel1.Controls.Count > 0 Then
                            LogMessage($"SplitContainer.Panel1 has {splitContainer.Panel1.Controls.Count} controls")
                            Dim schedulesPanel = splitContainer.Panel1.Controls(0)
                            If schedulesPanel IsNot Nothing Then
                                LogMessage($"schedulesPanel found, type: {schedulesPanel.GetType().Name}, control count: {schedulesPanel.Controls.Count}")
                                
                                ' Remove existing DataGridView if there is one
                                For i As Integer = schedulesPanel.Controls.Count - 1 To 0 Step -1
                                    If TypeOf schedulesPanel.Controls(i) Is DataGridView Then
                                        LogMessage($"Removing existing DataGridView: {schedulesPanel.Controls(i).Name}")
                                        schedulesPanel.Controls.RemoveAt(i)
                                    End If
                                Next
                                
                                ' Add the new grid
                                schedulesPanel.Controls.Add(entriesGridView)
                                
                                ' Ensure proper Z-order if there's a header panel
                                For i As Integer = 0 To schedulesPanel.Controls.Count - 1
                                    If TypeOf schedulesPanel.Controls(i) Is Panel AndAlso
                                       schedulesPanel.Controls(i).Dock = DockStyle.Top Then
                                        schedulesPanel.Controls(i).BringToFront()
                                    End If
                                Next
                                
                                LogMessage("entriesGridView successfully reinitialized and added to the UI.")
                            Else
                                LogMessage("ERROR: schedulesPanel is null.")
                            End If
                        Else
                            LogMessage("ERROR: SplitContainer.Panel1 has no controls.")
                        End If
                    Else
                        LogMessage("ERROR: Could not find SplitContainer in entriesPanel.")
                    End If
                Else
                    LogMessage("ERROR: entriesPanel is null or has no controls.")
                End If
            Else
                LogMessage("ERROR: tabPageEntries is null or has no controls.")
            End If
        End If
    End Sub

    Private Sub UpdateTimer_Tick(sender As Object, e As EventArgs) Handles updateTimer.Tick
        If Not String.IsNullOrEmpty(selectedRoom) AndAlso Not String.IsNullOrEmpty(selectedFolderPath) Then
            LogMessage("Auto-sync triggered for: " & selectedRoom)
            SyncFiles()
            FetchSchedulesData()
        End If
    End Sub

    ' Timer to update countdown display in real-time
    Private Sub CountdownTimer_Tick(sender As Object, e As EventArgs) Handles countdownTimer.Tick
        ' Update the countdown column if visible
        If mainTabControl.SelectedTab Is tabPageEntries AndAlso entriesGridView IsNot Nothing Then
            Try
                Dim now = DateTime.Now
                
                For Each row As DataGridViewRow In entriesGridView.Rows
                    If row.Cells("EventDate").Value IsNot Nothing AndAlso row.Cells("StartTime").Value IsNot Nothing Then
                        ' Parse the event date and time
                        Dim eventDate As DateTime
                        Dim eventTime As DateTime
                        Dim eventDateTime As DateTime
                        
                        If DateTime.TryParse(row.Cells("EventDate").Value.ToString(), eventDate) AndAlso 
                           DateTime.TryParse(row.Cells("StartTime").Value.ToString(), eventTime) Then
                            
                            ' Combine date and time
                            eventDateTime = New DateTime(
                                eventDate.Year, 
                                eventDate.Month, 
                                eventDate.Day, 
                                eventTime.Hour, 
                                eventTime.Minute, 
                                0)
                            
                            ' Calculate time remaining
                            Dim timeRemaining As TimeSpan = eventDateTime - now
                            
                            ' Update countdown display
                            If timeRemaining.TotalMinutes < 0 Then
                                ' Event has already started
                                If eventDate.Date = DateTime.Today Then
                                    row.Cells("Countdown").Value = "In progress"
                                Else
                                    row.Cells("Countdown").Value = "Past"
                                End If
                            Else
                                ' Event is in the future
                                If timeRemaining.TotalDays >= 1 Then
                                    ' More than a day away
                                    row.Cells("Countdown").Value = $"{Math.Floor(timeRemaining.TotalDays)}d {timeRemaining.Hours}h"
                                ElseIf timeRemaining.TotalHours >= 1 Then
                                    ' Hours away
                                    row.Cells("Countdown").Value = $"{Math.Floor(timeRemaining.TotalHours)}h {timeRemaining.Minutes}m"
                                Else
                                    ' Minutes away
                                    row.Cells("Countdown").Value = $"{timeRemaining.Minutes}m {timeRemaining.Seconds}s"
                                End If
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                ' Ignore errors in countdown update
                ' LogMessage("Countdown update error: " & ex.Message)
            End Try
        End If
    End Sub

    ' Stop timer when form is closing
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        updateTimer.Stop()
        countdownTimer.Stop()
    End Sub

    ' Add this method to handle form resize
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        ' Recenter any controls that need it when form is resized
        Try
            Dim roomPanel = tabPageMain.Controls(0).Controls(1)
            If roomPanel IsNot Nothing AndAlso comboBoxRooms IsNot Nothing Then
                comboBoxRooms.Location = New Point((roomPanel.Width - comboBoxRooms.Width) \ 2, comboBoxRooms.Location.Y)
        End If

            Dim logoPanel = tabPageMain.Controls(0).Controls(0)
            If logoPanel IsNot Nothing AndAlso logoPanel.Controls.Count > 0 Then
                Dim logoBox = DirectCast(logoPanel.Controls(0), PictureBox)
                logoBox.Location = New Point((logoPanel.Width - logoBox.Width) \ 2, (logoPanel.Height - logoBox.Height) \ 2)
        End If

            ' Adjust SplitContainer distance on resize
            If tabPageEntries IsNot Nothing AndAlso tabPageEntries.Controls.Count > 0 Then
                Dim entriesPanel = tabPageEntries.Controls(0)
                If entriesPanel IsNot Nothing AndAlso entriesPanel.Controls.Count > 0 Then
                    Dim splitContainer = TryCast(entriesPanel.Controls(0), SplitContainer)
                    If splitContainer IsNot Nothing Then
                        ' Set the splitter distance to 40% of the container height
                        splitContainer.SplitterDistance = Math.Max(150, CInt(splitContainer.Height * 0.4))
                        
                        ' Adjust description text box width in detail panel
                        Dim detailContentPanel = FindControlRecursive(splitContainer.Panel2, GetType(Panel), "")
                        If detailContentPanel IsNot Nothing Then
                            Dim descriptionTextBox = FindControlByName(detailContentPanel, "detailDescriptionText")
                            If descriptionTextBox IsNot Nothing Then
                                descriptionTextBox.Width = detailContentPanel.Width - 20
        End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            ' Log resize errors instead of silently ignoring
            LogMessage("Warning: Error during form resize: " & ex.Message)
        End Try
    End Sub
    
    ' Helper function to find a control recursively by type
    Private Function FindControlRecursive(container As Control, controlType As Type, nameFilter As String) As Control
        For Each ctrl As Control In container.Controls
            If ctrl.GetType() Is controlType AndAlso (String.IsNullOrEmpty(nameFilter) OrElse ctrl.Name.Contains(nameFilter)) Then
                Return ctrl
            End If
            
            If ctrl.Controls.Count > 0 Then
                Dim found = FindControlRecursive(ctrl, controlType, nameFilter)
                If found IsNot Nothing Then
                    Return found
                End If
            End If
        Next
        
        Return Nothing
    End Function

    ' Handle grid selection changed to update detail view
    Private Sub EntriesGridView_SelectionChanged(sender As Object, e As EventArgs)
        Try
            If entriesGridView.SelectedRows.Count > 0 Then
                Dim selectedRow = entriesGridView.SelectedRows(0)
                
                ' Get the schedule ID from the selected row
                Dim scheduleId = selectedRow.Cells("EventId").Value.ToString()
                LogMessage($"Schedule selected: ID={scheduleId}")
                
                ' Update schedule details for the selected schedule
                UpdateScheduleDetails(scheduleId)
            End If
        Catch ex As Exception
            LogMessage("Error in selection changed event: " & ex.Message)
        End Try
    End Sub
    
    ' Helper function to find a control by name in a container
    Private Function FindControlByName(container As Control, name As String) As Control
        For Each ctrl As Control In container.Controls
            If ctrl.Name = name Then
                Return ctrl
            End If
            
            If ctrl.Controls.Count > 0 Then
                Dim found = FindControlByName(ctrl, name)
                If found IsNot Nothing Then
                    Return found
                End If
            End If
        Next
        
        Return Nothing
    End Function
    
    ' Helper function to find a schedule by ID
    Private Function FindScheduleById(id As String) As Dictionary(Of String, Object)
        If schedulesList Is Nothing Then
            Return Nothing
        End If

        Return schedulesList.Find(Function(s) s.ContainsKey("id") AndAlso s("id").ToString() = id)
    End Function
    
    Private Async Sub FetchSchedulesData()
        If String.IsNullOrEmpty(selectedRoom) Then
            LogMessage("Error: No room selected.")
            Return
        End If
        
        ' Ensure eventsList is not null
        If eventsList Is Nothing Then
            LogMessage("Error: eventsList is null.")
            Return
        End If
        
        Try
            ' Find the selected event ID
            Dim selectedEvent = eventsList.Find(Function(evt) evt("title").ToString() = selectedRoom)
            If selectedEvent Is Nothing OrElse Not selectedEvent.ContainsKey("id") Then
                LogMessage("Error: Could not find event ID for selected room")
                Return
            End If
            
            Dim eventId = selectedEvent("id").ToString()
            LogMessage($"Found event ID: {eventId} for room: {selectedRoom}")
            
            ' Check if the event has schedules listed
            Dim schedulesIds As String = ""
            If selectedEvent.ContainsKey("schedules") Then
                schedulesIds = selectedEvent("schedules").ToString().Trim()
                LogMessage($"Event has schedules: {schedulesIds}")
            Else
                LogMessage("Event has no schedules listed")
            End If
            
            ' Create a new DataGridView for displaying schedules (simpler approach)
            CreateSimpleScheduleDisplay()
            
            Using client As New HttpClient()
                Try
                    ' Add Basic Authentication
                    Dim byteArray As Byte() = Encoding.ASCII.GetBytes($"{apiUsername}:{apiPassword}")
                    client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(byteArray))
                    
                    LogMessage($"Fetching schedule data from: {scheduleDataApi}")
                    
                    ' Fetch schedule data
                    Dim response As HttpResponseMessage = Await client.GetAsync(scheduleDataApi)
                    LogMessage($"API response status: {response.StatusCode}")
                    
                    response.EnsureSuccessStatusCode()
                    
                    Dim result As String = Await response.Content.ReadAsStringAsync()
                    LogMessage("Successfully fetched schedule data")
                    
                    ' Parse the JSON response
                    LogMessage("Parsing JSON response...")
                    Dim jsonResponse = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(result)
                    
                    If jsonResponse Is Nothing Then
                        LogMessage("Error: JSON response couldn't be parsed")
                        Return
                    End If
                    
                    If Not jsonResponse.ContainsKey("items") Then
                        LogMessage("Error: JSON response doesn't contain 'items' key")
                        LogMessage($"Available keys: {String.Join(", ", jsonResponse.Keys)}")
                        Return
                    End If
                    
                    LogMessage($"Found items in response: {jsonResponse("items").ToString().Length} characters")
                    schedulesList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(jsonResponse("items").ToString())
                    
                    ' Ensure schedulesList is not null
                    If schedulesList Is Nothing Then
                        LogMessage("Error: schedulesList is null after parsing")
                        Return
                    End If
                    
                    LogMessage($"Parsed {schedulesList.Count} schedules")
                    
                    ' Filter schedules based on scheduleIds listed in the event, or by date range if no specific schedules
                    Dim filteredSchedules As New List(Of Dictionary(Of String, Object))
                    
                    If Not String.IsNullOrEmpty(schedulesIds) Then
                        ' Clean up the schedule IDs for proper parsing
                        schedulesIds = schedulesIds.Replace("[", "").Replace("]", "").Trim()
                        
                        ' Split the scheduleIds string into an array, making sure to filter out empty entries
                        Dim scheduleIdArray = schedulesIds.Split(New Char() {" "c, ","c}, StringSplitOptions.RemoveEmptyEntries)
                        
                        LogMessage($"Filtering {schedulesList.Count} schedules by {scheduleIdArray.Length} IDs: {String.Join(", ", scheduleIdArray)}")
                        
                        For Each scheduleId In scheduleIdArray
                            Dim id = scheduleId.Trim()
                            If Not String.IsNullOrEmpty(id) Then
                                Dim schedule = schedulesList.Find(Function(s) s.ContainsKey("id") AndAlso s("id").ToString() = id)
                                If schedule IsNot Nothing Then
                                    filteredSchedules.Add(schedule)
                                    LogMessage($"Added schedule ID {id} to filtered list")
                                Else
                                    LogMessage($"Could not find schedule with ID {id}")
                                End If
                            End If
                        Next
                    End If
                    
                    ' If no schedules were found by ID, try by date range
                    If filteredSchedules.Count = 0 Then
                        LogMessage("No schedules found by ID, trying by date range...")
                        ' Filter by event date range as fallback
                        Dim eventStartDate As DateTime
                        Dim eventEndDate As DateTime
                        
                        If DateTime.TryParse(selectedEvent("start_date").ToString(), eventStartDate) AndAlso
                           DateTime.TryParse(selectedEvent("end_date").ToString(), eventEndDate) Then
                            
                            LogMessage($"Filtering schedules by date range: {eventStartDate.ToString("yyyy-MM-dd")} to {eventEndDate.ToString("yyyy-MM-dd")}")
                            
                            For Each schedule In schedulesList
                                If schedule.ContainsKey("date") Then
                                    Dim scheduleDate As DateTime
                                    If DateTime.TryParse(schedule("date").ToString(), scheduleDate) Then
                                        If scheduleDate >= eventStartDate AndAlso scheduleDate <= eventEndDate Then
                                            filteredSchedules.Add(schedule)
                                            LogMessage($"Added schedule with date {scheduleDate.ToString("yyyy-MM-dd")} to filtered list")
                                        End If
                                    End If
                                End If
                            Next
                        Else
                            LogMessage("Warning: Could not parse event start/end dates for filtering")
                        End If
                    End If
                    
                    ' If we still have no schedules, add all of them as a last resort
                    If filteredSchedules.Count = 0 Then
                        LogMessage("No schedules found by ID or date range, adding all schedules")
                        filteredSchedules = schedulesList
                    End If
                    
                    LogMessage($"Found {filteredSchedules.Count} relevant schedules for event ID: {eventId}")
                    
                    ' Try to fetch speakers but continue even if it fails
                    Try
                        LogMessage("Fetching speakers data...")
                        Await FetchSpeakersData()
                    Catch ex As Exception
                        LogMessage("Error fetching speakers (continuing anyway): " & ex.Message)
                    End Try
                    
                    ' Create a default event start/end time from the event data
                    Dim defaultStartTime = If(selectedEvent.ContainsKey("start_time"), selectedEvent("start_time").ToString(), "")
                    Dim defaultEndTime = If(selectedEvent.ContainsKey("end_time"), selectedEvent("end_time").ToString(), "")
                    
                    ' Populate the grid with schedule data - using our SIMPLE approach
                    If entriesGridView IsNot Nothing AndAlso filteredSchedules.Count > 0 Then
                        LogMessage("Populating grid with schedule data using enhanced approach...")
                        
                        ' Clear any existing rows
                        entriesGridView.Rows.Clear()
                        
                        ' Add rows for each schedule
                        For Each schedule In filteredSchedules
                            Try
                                Dim title As String = If(schedule.ContainsKey("program_title"), schedule("program_title").ToString(), "Unknown")
                                Dim eventDate As String = If(schedule.ContainsKey("date"), schedule("date").ToString(), "Unknown")
                                Dim startTime As String = defaultStartTime
                                Dim endTime As String = defaultEndTime
                                Dim room As String = If(selectedEvent.ContainsKey("location") AndAlso 
                                                        TypeOf selectedEvent("location") Is Newtonsoft.Json.Linq.JObject AndAlso
                                                        DirectCast(selectedEvent("location"), Newtonsoft.Json.Linq.JObject).ContainsKey("address"),
                                                    DirectCast(selectedEvent("location"), Newtonsoft.Json.Linq.JObject)("address").ToString(),
                                                    "Main Room")
                                Dim speakers As String = "None"
                                Dim topicTitle As String = ""
                                Dim objective As String = ""
                                Dim scheduleId As String = If(schedule.ContainsKey("id"), schedule("id").ToString(), "0")
                                
                                ' Extract additional information from schedule_slot if available
                                If schedule.ContainsKey("schedule_slot") AndAlso 
                                  TypeOf schedule("schedule_slot") Is Newtonsoft.Json.Linq.JArray Then
                                    
                                    Dim scheduleSlot = DirectCast(schedule("schedule_slot"), Newtonsoft.Json.Linq.JArray)
                                    
                                    If scheduleSlot.Count > 0 Then
                                        ' Process the first schedule slot (can be expanded to handle multiple slots)
                                        Dim slotData = scheduleSlot(0).ToString()
                                        LogMessage($"Processing schedule slot: {slotData}")
                                        
                                        ' Extract data using regular expressions for better reliability
                                        Dim topicMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_schedule_topic=([^;]+)")
                                        If topicMatch.Success Then
                                            topicTitle = topicMatch.Groups(1).Value.Trim()
                                        End If
                                        
                                        Dim startTimeMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_start_time=([^;]+)")
                                        If startTimeMatch.Success Then
                                            startTime = startTimeMatch.Groups(1).Value.Trim()
                                        End If
                                        
                                        Dim endTimeMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_end_time=([^;]+)")
                                        If endTimeMatch.Success Then
                                            endTime = endTimeMatch.Groups(1).Value.Trim()
                                        End If
                                        
                                        Dim roomMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_room=([^;]+)")
                                        If roomMatch.Success Then
                                            room = roomMatch.Groups(1).Value.Trim()
                                        End If
                                        
                                        Dim objectiveMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_objective=([^;]+)")
                                        If objectiveMatch.Success Then
                                            objective = objectiveMatch.Groups(1).Value.Trim()
                                        End If
                                        
                                        ' Extract speakers information
                                        Dim speakersMatch = System.Text.RegularExpressions.Regex.Match(slotData, "speakers=([^;]+)")
                                        If speakersMatch.Success Then
                                            Dim speakersData = speakersMatch.Groups(1).Value.Trim()
                                            LogMessage($"Found speakers data: {speakersData}")
                                            
                                            ' If we have speaker IDs, try to look them up
                                            If speakersList IsNot Nothing AndAlso speakersList.Count > 0 Then
                                                ' First, check if it's a direct array of integers
                                                Try
                                                    ' Try to parse as a JSON array directly
                                                    Dim speakerIdsArray = JsonConvert.DeserializeObject(Of List(Of Integer))(speakersData)
                                                    If speakerIdsArray IsNot Nothing AndAlso speakerIdsArray.Count > 0 Then
                                                        Dim speakerNames As New List(Of String)
                                                        
                                                        For Each speakerId In speakerIdsArray
                                                            LogMessage($"Looking up speaker ID: {speakerId}")
                                                            Dim speaker = GetSpeakerById(speakerId.ToString())
                                                            
                                                            If speaker IsNot Nothing AndAlso speaker.ContainsKey("name") Then
                                                                speakerNames.Add(speaker("name").ToString())
                                                                LogMessage($"Found speaker: {speaker("name")}")
                                                            End If
                                                        Next
                                                        
                                                        If speakerNames.Count > 0 Then
                                                            speakers = String.Join(", ", speakerNames)
                                                            LogMessage($"Set speakers to: {speakers}")
                                                        End If
                                                    End If
                                                Catch jsonEx As Exception
                                                    LogMessage($"Not a JSON array, trying regex patterns: {jsonEx.Message}")
                                                    
                                                    ' Try to extract IDs using regex patterns
                                                    ' First pattern: Check for ID=number pattern
                                                    Dim idMatches = System.Text.RegularExpressions.Regex.Matches(speakersData, "id=(\d+)")
                                                    If idMatches.Count > 0 Then
                                                        Dim speakerNames As New List(Of String)
                                                        
                                                        For Each match In idMatches
                                                            Dim speakerId = match.Groups(1).Value
                                                            LogMessage($"Regex found speaker ID: {speakerId}")
                                                            Dim speaker = GetSpeakerById(speakerId)
                                                            
                                                            If speaker IsNot Nothing AndAlso speaker.ContainsKey("name") Then
                                                                speakerNames.Add(speaker("name").ToString())
                                                                LogMessage($"Found speaker: {speaker("name")}")
                                                            End If
                                                        Next
                                                        
                                                        If speakerNames.Count > 0 Then
                                                            speakers = String.Join(", ", speakerNames)
                                                            LogMessage($"Set speakers to: {speakers}")
                                                        End If
                                                    Else
                                                        ' Second pattern: Check for plain numbers separated by spaces
                                                        Dim numberMatches = System.Text.RegularExpressions.Regex.Matches(speakersData, "(\d+)")
                                                        If numberMatches.Count > 0 Then
                                                            Dim speakerNames As New List(Of String)
                                                            
                                                            For Each match In numberMatches
                                                                Dim speakerId = match.Groups(1).Value
                                                                LogMessage($"Found plain number speaker ID: {speakerId}")
                                                                Dim speaker = GetSpeakerById(speakerId)
                                                                
                                                                If speaker IsNot Nothing AndAlso speaker.ContainsKey("name") Then
                                                                    speakerNames.Add(speaker("name").ToString())
                                                                    LogMessage($"Found speaker: {speaker("name")}")
                                                                End If
                                                            Next
                                                            
                                                            If speakerNames.Count > 0 Then
                                                                speakers = String.Join(", ", speakerNames)
                                                                LogMessage($"Set speakers to: {speakers}")
                                                            End If
                                                        End If
                                                    End If
                                                End Try
                                            Else
                                                LogMessage("No speakers list available for lookup")
                                            End If
                                        Else
                                            ' If no speakers in slot, try to use the event's speakers
                                            If selectedEvent.ContainsKey("speaker") AndAlso Not String.IsNullOrEmpty(selectedEvent("speaker").ToString()) Then
                                                Dim eventSpeakers = selectedEvent("speaker").ToString().Trim()
                                                LogMessage($"Using event level speakers: {eventSpeakers}")
                                                
                                                ' Extract speaker IDs using regex
                                                Dim speakerMatches = System.Text.RegularExpressions.Regex.Matches(eventSpeakers, "(\d+)")
                                                If speakerMatches.Count > 0 Then
                                                    Dim speakerNames As New List(Of String)
                                                    
                                                    For Each match In speakerMatches
                                                        Dim speakerId = match.Groups(1).Value
                                                        LogMessage($"Event speaker ID: {speakerId}")
                                                        Dim speaker = GetSpeakerById(speakerId)
                                                        
                                                        If speaker IsNot Nothing AndAlso speaker.ContainsKey("name") Then
                                                            speakerNames.Add(speaker("name").ToString())
                                                            LogMessage($"Found event speaker: {speaker("name")}")
                                                        End If
                                                    Next
                                                    
                                                    If speakerNames.Count > 0 Then
                                                        speakers = String.Join(", ", speakerNames)
                                                        LogMessage($"Set event speakers to: {speakers}")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                
                                ' If topic title is still empty, use the main title
                                If String.IsNullOrEmpty(topicTitle) Then
                                    topicTitle = title
                                End If
                                
                                LogMessage($"Adding enhanced row: {title} | Topic: {topicTitle} | {eventDate} at {startTime}-{endTime} in {room} with speakers: {speakers}")
                                
                                ' Add the enhanced row with more information
                                entriesGridView.Rows.Add(scheduleId, title, eventDate, startTime, endTime, room, speakers, topicTitle, objective)
                            Catch ex As Exception
                                LogMessage($"Error processing schedule: {ex.Message}")
                            End Try
                        Next
                        
                        LogMessage($"Added {entriesGridView.Rows.Count} rows to enhanced grid")
                        
                        ' Format the grid
                        FormatDataGridView()
                        
                        ' Update the header panel with the selected event name
                        Try
                            Dim headerPanel = FindControlRecursive(tabPageEntries, GetType(Panel), "")
                            If headerPanel IsNot Nothing Then
                                Dim roomNameLabel = FindControlByName(headerPanel, "lblRoomName")
                                If roomNameLabel IsNot Nothing AndAlso TypeOf roomNameLabel Is Label Then
                                    DirectCast(roomNameLabel, Label).Text = selectedRoom
                                End If
                            End If
                        Catch ex As Exception
                            LogMessage($"Error updating header: {ex.Message}")
                        End Try
                        
                        ' Switch to the entries tab
                        mainTabControl.SelectedTab = tabPageEntries
                        
                        ' Force update of the UI
                        Application.DoEvents()
                        LogMessage("Switched to Entries tab and updated UI")
                    Else
                        LogMessage("Warning: Could not populate grid - entriesGridView is null or no schedules found")
                    End If
                    
                    ' Update the event details panel with the selected event
                    If selectedEvent IsNot Nothing Then
                        UpdateEventDetails(selectedEvent)
                    End If
                    
                Catch ex As Exception
                    LogMessage("Error fetching schedules: " & ex.Message)
                    If ex.InnerException IsNot Nothing Then
                        LogMessage("Inner exception: " & ex.InnerException.Message)
                    End If
                    LogMessage("Stack trace: " & ex.StackTrace)
                End Try
            End Using
        Catch ex As Exception
            LogMessage("Error in FetchSchedulesData: " & ex.Message)
            If ex.InnerException IsNot Nothing Then
                LogMessage("Inner exception: " & ex.InnerException.Message)
            End If
            LogMessage("Stack trace: " & ex.StackTrace)
        End Try
    End Sub

    ' Create a simple schedule display that doesn't depend on complex UI hierarchy
    Private Sub CreateSimpleScheduleDisplay()
        Try
            LogMessage("Creating a simple schedule display...")
            
            ' Make sure tabPageEntries exists
            If tabPageEntries Is Nothing Then
                LogMessage("tabPageEntries is null, creating a new one")
                tabPageEntries = New TabPage("Events Data")
                mainTabControl.Controls.Add(tabPageEntries)
            End If
            
            ' Clear any existing controls from the tab page to start fresh
            tabPageEntries.Controls.Clear()
            
            ' Create main container panel with TableLayoutPanel
            Dim mainContainer As New TableLayoutPanel()
            mainContainer.Dock = DockStyle.Fill
            mainContainer.ColumnCount = 1
            mainContainer.RowCount = 2
            mainContainer.Padding = New Padding(10)
            
            ' Configure rows - top row for grid (55% height), bottom row for details (45% height)
            mainContainer.RowStyles.Clear()
            mainContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 55))
            mainContainer.RowStyles.Add(New RowStyle(SizeType.Percent, 45))
            mainContainer.AutoScroll = True
            
            ' Panel for the grid section
            Dim gridPanel As New Panel()
            gridPanel.Dock = DockStyle.Fill
            gridPanel.Padding = New Padding(0, 0, 0, 5)  ' Add some spacing between grid and details
            gridPanel.AutoScroll = True
            
            ' Create header for grid with reduced height
            Dim gridHeaderPanel As New Panel()
            gridHeaderPanel.Dock = DockStyle.Top
            gridHeaderPanel.Height = 45  ' Reduced from 60 to 45
            gridHeaderPanel.BackColor = ColorTranslator.FromHtml("#9c1c13")
            
            ' Grid header with room name and current time - adjusted font sizes
            Dim gridHeaderLabel As New Label()
            gridHeaderLabel.Text = "Event Schedules"
            gridHeaderLabel.Dock = DockStyle.Left
            gridHeaderLabel.Width = 200
            gridHeaderLabel.Font = New Font("Segoe UI", 11.0F, FontStyle.Bold)  ' Reduced from 12 to 11
            gridHeaderLabel.ForeColor = Color.White
            gridHeaderLabel.TextAlign = ContentAlignment.MiddleLeft
            gridHeaderLabel.Padding = New Padding(10, 0, 0, 0)
            gridHeaderLabel.Name = "lblScheduleHeader"
            
            ' Room name display in center
            Dim roomLabel As New Label()
            roomLabel.Text = If(String.IsNullOrEmpty(selectedRoom), "No Room Selected", selectedRoom)
            roomLabel.Dock = DockStyle.Fill
            roomLabel.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
            roomLabel.ForeColor = Color.White
            roomLabel.TextAlign = ContentAlignment.MiddleCenter
            roomLabel.Name = "lblRoomName"
            
            ' Current time display on right
            Dim currentTimeLabel As New Label()
            currentTimeLabel.Text = DateTime.Now.ToString("HH:mm:ss")
            currentTimeLabel.Dock = DockStyle.Right
            currentTimeLabel.Width = 150
            currentTimeLabel.Font = New Font("Segoe UI", 14.0F, FontStyle.Bold)
            currentTimeLabel.ForeColor = Color.White
            currentTimeLabel.TextAlign = ContentAlignment.MiddleRight
            currentTimeLabel.Padding = New Padding(0, 0, 10, 0)
            currentTimeLabel.Name = "lblCurrentTime"
            
            ' Timer for current time
            Dim clockTimer As New Timer()
            clockTimer.Interval = 1000
            AddHandler clockTimer.Tick, Sub(sender, e)
                If currentTimeLabel IsNot Nothing Then
                    currentTimeLabel.Text = DateTime.Now.ToString("HH:mm:ss")
                End If
            End Sub
            clockTimer.Start()
            
            ' Add labels to header
            gridHeaderPanel.Controls.Add(currentTimeLabel)
            gridHeaderPanel.Controls.Add(gridHeaderLabel)
            gridHeaderPanel.Controls.Add(roomLabel)
            
            ' Initialize and configure DataGridView
            entriesGridView = New DataGridView()
            entriesGridView.Name = "entriesGridView"
            entriesGridView.Dock = DockStyle.Fill
            entriesGridView.BackgroundColor = Color.White
            entriesGridView.BorderStyle = BorderStyle.None
            entriesGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            entriesGridView.RowHeadersVisible = False
            entriesGridView.AllowUserToAddRows = False
            entriesGridView.AllowUserToDeleteRows = False
            entriesGridView.ReadOnly = True
            entriesGridView.Font = New Font("Segoe UI", 9.0F)
            entriesGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            entriesGridView.MultiSelect = False
            entriesGridView.RowTemplate.Height = 25
            entriesGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            entriesGridView.AllowUserToResizeRows = False
            entriesGridView.AllowUserToResizeColumns = True
            
            ' Add columns with specific widths
            entriesGridView.Columns.Add("EventId", "ID")
            entriesGridView.Columns.Add("Title", "Title")
            entriesGridView.Columns.Add("EventDate", "Date")
            entriesGridView.Columns.Add("StartTime", "Start")
            entriesGridView.Columns.Add("EndTime", "End")
            entriesGridView.Columns.Add("Room", "Room")
            entriesGridView.Columns.Add("Speakers", "Speakers")
            entriesGridView.Columns.Add("TopicTitle", "Topic")
            entriesGridView.Columns.Add("Objective", "Description")
            
            ' Configure column properties
            entriesGridView.Columns("EventId").Visible = False
            entriesGridView.Columns("Title").FillWeight = 15
            entriesGridView.Columns("EventDate").FillWeight = 10
            entriesGridView.Columns("StartTime").FillWeight = 8
            entriesGridView.Columns("EndTime").FillWeight = 8
            entriesGridView.Columns("Room").FillWeight = 15
            entriesGridView.Columns("Speakers").FillWeight = 12
            entriesGridView.Columns("TopicTitle").FillWeight = 15
            entriesGridView.Columns("Objective").FillWeight = 17
            
            ' Enable column auto-sizing
            For Each col As DataGridViewColumn In entriesGridView.Columns
                col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                col.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            Next
            
            ' Add grid and header to panel
            gridPanel.Controls.Add(entriesGridView)
            gridPanel.Controls.Add(gridHeaderPanel)
            
            ' Create details panel
            Dim detailsPanel As New Panel()
            detailsPanel.Dock = DockStyle.Fill
            detailsPanel.BackColor = Color.White
            detailsPanel.Padding = New Padding(0)
            detailsPanel.AutoScroll = True
            
            ' Details header
            Dim detailsHeaderPanel As New Panel()
            detailsHeaderPanel.Dock = DockStyle.Top
            detailsHeaderPanel.Height = 40
            detailsHeaderPanel.BackColor = ColorTranslator.FromHtml("#9c1c13")
            
            Dim detailsHeaderLabel As New Label()
            detailsHeaderLabel.Text = "Event Schedule Details"
            detailsHeaderLabel.Dock = DockStyle.Fill
            detailsHeaderLabel.Font = New Font("Segoe UI", 12.0F, FontStyle.Bold)
            detailsHeaderLabel.ForeColor = Color.White
            detailsHeaderLabel.TextAlign = ContentAlignment.MiddleLeft
            detailsHeaderLabel.Padding = New Padding(10, 0, 0, 0)
            
            detailsHeaderPanel.Controls.Add(detailsHeaderLabel)
            
            ' Create details content panel
            Dim detailsContent As New TableLayoutPanel()
            CreateDetailsContentPanel(detailsContent)
            
            ' Add panels to details panel
            detailsPanel.Controls.Add(detailsContent)
            detailsPanel.Controls.Add(detailsHeaderPanel)
            
            ' Add panels to main container
            mainContainer.Controls.Add(gridPanel, 0, 0)
            mainContainer.Controls.Add(detailsPanel, 0, 1)
            
            ' Add main container to tab page
            tabPageEntries.Controls.Add(mainContainer)
            
            ' Add selection changed handler
            AddHandler entriesGridView.SelectionChanged, AddressOf EntriesGridView_SelectionChanged
            
            LogMessage("Enhanced schedule display created successfully")
        Catch ex As Exception
            LogMessage("Error creating enhanced schedule display: " & ex.Message)
        End Try
    End Sub

    ' Helper method to add detail rows
    Private Sub AddDetailRow(container As TableLayoutPanel, labelText As String, controlName As String, rowIndex As Integer)
        ' Label with simple styling
        Dim label As New Label()
        label.Text = labelText
        label.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        label.AutoSize = True
        label.Dock = DockStyle.Fill
        label.Margin = New Padding(10, 3, 5, 3)
        container.Controls.Add(label, 0, rowIndex)
        
        ' Value with simple styling
        Dim value As New Label()
        value.Name = controlName
        value.Font = New Font("Segoe UI", 10.0F)
        value.AutoSize = True
        value.MaximumSize = New Size(container.Width * 3 / 4, 0)  ' Allow text wrapping
        value.AutoEllipsis = True  ' Show ... when text is truncated
        value.Dock = DockStyle.Fill
        value.Margin = New Padding(5, 3, 10, 3)
        container.Controls.Add(value, 1, rowIndex)
    End Sub

    ' Update the details content panel creation
    Private Sub CreateDetailsContentPanel(ByRef detailsContent As TableLayoutPanel)
        detailsContent.Dock = DockStyle.Fill
        detailsContent.Padding = New Padding(10)
        detailsContent.BackColor = Color.FromArgb(248, 249, 250)
        detailsContent.ColumnCount = 2
        detailsContent.RowCount = 8
        detailsContent.AutoScroll = True
        
        ' Configure columns - labels (30%) and values (70%)
        detailsContent.ColumnStyles.Clear()
        detailsContent.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30))
        detailsContent.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 70))
        
        ' Configure rows for better spacing
        For i As Integer = 0 To detailsContent.RowCount - 1
            detailsContent.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        Next
        
        ' Add details fields with improved spacing
        AddDetailRow(detailsContent, "Title:", "lblScheduleTitle", 0)
        AddDetailRow(detailsContent, "Date:", "lblScheduleTime", 1)
        AddDetailRow(detailsContent, "Location:", "lblEventLocation", 2)
        AddDetailRow(detailsContent, "Timezone:", "lblEventTimezone", 3)
        AddDetailRow(detailsContent, "Status:", "lblEventStatus", 4)
        AddDetailRow(detailsContent, "Organizer:", "lblEventOrganizer", 5)
        AddDetailRow(detailsContent, "Event Type:", "lblEventType", 6)
        
        ' Description field with improved layout
        Dim descLabelPanel As New Panel()
        descLabelPanel.Dock = DockStyle.Fill
        descLabelPanel.Padding = New Padding(5)
        descLabelPanel.BackColor = Color.FromArgb(248, 249, 250)
        
        Dim descLabel As New Label()
        descLabel.Text = "Description:"
        descLabel.Font = New Font("Segoe UI", 10.0F, FontStyle.Bold)
        descLabel.AutoSize = True
        descLabel.Dock = DockStyle.Fill
        descLabelPanel.Controls.Add(descLabel)
        detailsContent.Controls.Add(descLabelPanel, 0, 7)
        
        ' Description text panel
        Dim descPanel As New Panel()
        descPanel.Dock = DockStyle.Fill
        descPanel.Padding = New Padding(5)
        descPanel.BackColor = Color.White
        
        Dim descText As New TextBox()
        descText.Multiline = True
        descText.ReadOnly = True
        descText.BorderStyle = BorderStyle.None
        descText.BackColor = Color.White
        descText.Dock = DockStyle.Fill
        descText.MinimumSize = New Size(0, 100)
        descText.ScrollBars = ScrollBars.Vertical
        descText.Name = "txtDescription"
        descPanel.Controls.Add(descText)
        detailsContent.Controls.Add(descPanel, 1, 7)
    End Sub

    Private Sub FormatDataGridView()
        Try
            ' Apply alternating row colors
            entriesGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245)
            entriesGridView.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#9c1c13")
            entriesGridView.DefaultCellStyle.SelectionForeColor = Color.White
            entriesGridView.ColumnHeadersDefaultCellStyle.BackColor = ColorTranslator.FromHtml("#4a4a4a")
            entriesGridView.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            entriesGridView.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0F, FontStyle.Bold)
            entriesGridView.ColumnHeadersHeight = 30
            entriesGridView.RowTemplate.Height = 24
            entriesGridView.BorderStyle = BorderStyle.None
            entriesGridView.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
            entriesGridView.EnableHeadersVisualStyles = False
            
            ' Format column widths for better appearance
            If entriesGridView.Columns.Count >= 9 Then
                entriesGridView.Columns("Title").Width = 150
                entriesGridView.Columns("EventDate").Width = 80
                entriesGridView.Columns("StartTime").Width = 70
                entriesGridView.Columns("EndTime").Width = 70
                entriesGridView.Columns("Room").Width = 80
                entriesGridView.Columns("Speakers").Width = 150
                entriesGridView.Columns("TopicTitle").Width = 150
                entriesGridView.Columns("Objective").Width = 200
                
                ' Better names for column headers
                entriesGridView.Columns("EventDate").HeaderText = "Date"
                entriesGridView.Columns("StartTime").HeaderText = "Start"
                entriesGridView.Columns("EndTime").HeaderText = "End"
                entriesGridView.Columns("TopicTitle").HeaderText = "Topic"
                entriesGridView.Columns("Objective").HeaderText = "Description"
                
                ' Set tooltip for columns
                For Each column As DataGridViewColumn In entriesGridView.Columns
                    column.ToolTipText = column.HeaderText
                Next
                
                ' Add wrapping for text in certain columns
                entriesGridView.Columns("Objective").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                entriesGridView.Columns("TopicTitle").DefaultCellStyle.WrapMode = DataGridViewTriState.True
                
                ' Highlight current day items and add countdown
                Dim today = DateTime.Today.Date
                Dim now = DateTime.Now
                
                ' Add countdown column if it doesn't exist
                If Not entriesGridView.Columns.Contains("Countdown") Then
                    entriesGridView.Columns.Add("Countdown", "Countdown")
                    entriesGridView.Columns("Countdown").Width = 100
                End If
                
                For Each row As DataGridViewRow In entriesGridView.Rows
                    If row.Cells("EventDate").Value IsNot Nothing AndAlso row.Cells("StartTime").Value IsNot Nothing Then
                        Dim eventDate As DateTime
                        Dim eventTime As DateTime
                        Dim eventDateTime As DateTime
                        
                        ' Parse event date
                        If DateTime.TryParse(row.Cells("EventDate").Value.ToString(), eventDate) Then
                            ' Parse event time
                            If DateTime.TryParse(row.Cells("StartTime").Value.ToString(), eventTime) Then
                                ' Combine date and time
                                eventDateTime = New DateTime(
                                    eventDate.Year, 
                                    eventDate.Month, 
                                    eventDate.Day, 
                                    eventTime.Hour, 
                                    eventTime.Minute, 
                                    0)
                                
                                ' Calculate time remaining
                                Dim timeRemaining As TimeSpan = eventDateTime - now
                                
                                ' Set countdown display
                                If timeRemaining.TotalMinutes < 0 Then
                                    ' Event has already started
                                    If eventDate.Date = today Then
                                        row.Cells("Countdown").Value = "In progress"
                                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235) ' Light red for current events
                                        row.DefaultCellStyle.Font = New Font(entriesGridView.Font, FontStyle.Bold)
                                    Else
                                        row.Cells("Countdown").Value = "Past"
                                    End If
                                Else
                                    ' Event is in the future
                                    If timeRemaining.TotalDays >= 1 Then
                                        ' More than a day away
                                        row.Cells("Countdown").Value = $"{Math.Floor(timeRemaining.TotalDays)}d {timeRemaining.Hours}h"
                                    ElseIf timeRemaining.TotalHours >= 1 Then
                                        ' Hours away
                                        row.Cells("Countdown").Value = $"{Math.Floor(timeRemaining.TotalHours)}h {timeRemaining.Minutes}m"
                                        
                                        ' Highlight upcoming events (next 2 hours)
                                        If timeRemaining.TotalHours <= 2 Then
                                            row.DefaultCellStyle.BackColor = Color.FromArgb(235, 255, 235) ' Light green for upcoming events
                                            row.DefaultCellStyle.Font = New Font(entriesGridView.Font, FontStyle.Bold)
                                        End If
                                    Else
                                        ' Minutes away
                                        row.Cells("Countdown").Value = $"{timeRemaining.Minutes}m {timeRemaining.Seconds}s"
                                        row.DefaultCellStyle.BackColor = Color.FromArgb(150, 255, 150) ' Brighter green for imminent events
                                        row.DefaultCellStyle.Font = New Font(entriesGridView.Font, FontStyle.Bold)
                                    End If
                                    
                                    ' Mark events today
                                    If eventDate.Date = today Then
                                        row.Cells("Countdown").Style.ForeColor = Color.FromArgb(180, 0, 0)
                                        row.Cells("Countdown").Style.Font = New Font(entriesGridView.Font, FontStyle.Bold)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            
            ' Apply sorting by date and time
            Try
                entriesGridView.Sort(entriesGridView.Columns("EventDate"), System.ComponentModel.ListSortDirection.Ascending)
                entriesGridView.Sort(entriesGridView.Columns("StartTime"), System.ComponentModel.ListSortDirection.Ascending)
            Catch ex As Exception
                ' Ignore sorting errors
                LogMessage("Sorting error (non-critical): " & ex.Message)
            End Try
            
            ' Automatically resize height of rows based on content
            entriesGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells
            
            ' Select the first row if available
            If entriesGridView.Rows.Count > 0 Then
                entriesGridView.Rows(0).Selected = True
            End If
            
            LogMessage("DataGridView formatted successfully with enhanced styling")
        Catch ex As Exception
            LogMessage("Error formatting DataGridView: " & ex.Message)
        End Try
    End Sub
    
    Private Async Function FetchSpeakersData() As Task
        Using client As New HttpClient()
            Try
                ' Add Basic Authentication
                Dim byteArray As Byte() = Encoding.ASCII.GetBytes($"{apiUsername}:{apiPassword}")
                client.DefaultRequestHeaders.Authorization = New Headers.AuthenticationHeaderValue("Basic",
                    Convert.ToBase64String(byteArray))
                
                LogMessage("Fetching speakers from: " & speakersDataApi)
                Dim response As HttpResponseMessage = Await client.GetAsync(speakersDataApi)
                response.EnsureSuccessStatusCode()
                
                Dim result As String = Await response.Content.ReadAsStringAsync()
                LogMessage("Received speaker data, length: " & result.Length)
                
                ' Parse the JSON response - the speakers API returns an array directly
                Try
                    ' First attempt to parse as a direct array
                    speakersList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(result)
                    LogMessage($"Successfully loaded {speakersList.Count} speakers from direct array")
                Catch arrayEx As Exception
                    ' If that fails, try to parse as an object with contained array
                    Try
                        Dim jsonResponse = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(result)
                        
                        If jsonResponse Is Nothing Then
                            LogMessage("Error: Speakers JSON response couldn't be parsed")
                            speakersList = New List(Of Dictionary(Of String, Object))()
                            Return
                        End If
                        
                        ' Check if the response contains the "value" key or "items" key
                        If jsonResponse.ContainsKey("value") Then
                            speakersList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(jsonResponse("value").ToString())
                            LogMessage($"Loaded {speakersList.Count} speakers from 'value' key")
                        ElseIf jsonResponse.ContainsKey("items") Then
                            ' Try "items" as an alternative key
                            speakersList = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(jsonResponse("items").ToString())
                            LogMessage($"Loaded {speakersList.Count} speakers from 'items' key")
                        Else
                            LogMessage("Error: Speakers response doesn't contain 'value' or 'items' key")
                            LogMessage($"Available keys: {String.Join(", ", jsonResponse.Keys)}")
                            speakersList = New List(Of Dictionary(Of String, Object))()
                        End If
                    Catch objEx As Exception
                        LogMessage("Error parsing speakers data: " & objEx.Message)
                        speakersList = New List(Of Dictionary(Of String, Object))()
                    End Try
                End Try
                
                ' Log the speaker IDs and names for debugging
                If speakersList IsNot Nothing AndAlso speakersList.Count > 0 Then
                    LogMessage("Available speakers:")
                    For Each speaker In speakersList
                        If speaker.ContainsKey("id") AndAlso speaker.ContainsKey("name") Then
                            LogMessage($"  Speaker ID: {speaker("id")} - Name: {speaker("name")}")
                        End If
                    Next
                End If
            Catch ex As Exception
                LogMessage("Error fetching speakers: " & ex.Message)
                ' Initialize with empty list to avoid null reference
                speakersList = New List(Of Dictionary(Of String, Object))()
            End Try
        End Using
    End Function
    
    Private Function GetSpeakerById(id As String) As Dictionary(Of String, Object)
        If speakersList Is Nothing Then
            Return Nothing
        End If
        
        Return speakersList.Find(Function(s) s.ContainsKey("id") AndAlso s("id").ToString() = id)
    End Function
    
    ' Update the event details panel with selected event information
    Private Sub UpdateEventDetails(selectedEvent As Dictionary(Of String, Object))
        Try
            If selectedEvent Is Nothing Then
                LogMessage("Cannot update details: No event selected")
                Return
            End If
            
            LogMessage("Updating event details panel with selected event information")
            
            ' Find the details panel controls
            If tabPageEntries Is Nothing OrElse tabPageEntries.Controls.Count = 0 Then
                LogMessage("Cannot update details: Tab page or controls not found")
                Return
            End If
            
            ' Get the main container
            Dim mainContainer = TryCast(tabPageEntries.Controls(0), TableLayoutPanel)
            If mainContainer Is Nothing Then
                LogMessage("Cannot update details: Main container not found")
                Return
            End If
            
            ' Get the details panel
            Dim detailsPanel = TryCast(mainContainer.Controls(1), Panel)
            If detailsPanel Is Nothing Then
                LogMessage("Cannot update details: Details panel not found")
                Return
            End If
            
            ' Find the details content panel
            Dim detailsContent = Nothing
            For Each ctrl In detailsPanel.Controls
                If TypeOf ctrl Is TableLayoutPanel Then
                    detailsContent = DirectCast(ctrl, TableLayoutPanel)
                    Exit For
                End If
            Next
            
            If detailsContent Is Nothing Then
                LogMessage("Cannot update details: Details content panel not found")
                Return
            End If
            
            ' Update title
            Dim titleLabel = FindControlByName(detailsContent, "lblScheduleTitle")
            If titleLabel IsNot Nothing Then
                DirectCast(titleLabel, Label).Text = If(selectedEvent.ContainsKey("title"), 
                    selectedEvent("title").ToString(), "Unknown Event")
            End If
            
            ' Update dates
            Dim datesLabel = FindControlByName(detailsContent, "lblScheduleTime")
            If datesLabel IsNot Nothing Then
                Dim startDate = If(selectedEvent.ContainsKey("start_date"), selectedEvent("start_date").ToString(), "Unknown")
                Dim endDate = If(selectedEvent.ContainsKey("end_date"), selectedEvent("end_date").ToString(), "Unknown")
                DirectCast(datesLabel, Label).Text = $"{startDate} to {endDate}"
            End If
            
            ' Update location
            Dim locationLabel = FindControlByName(detailsContent, "lblEventLocation")
            If locationLabel IsNot Nothing Then
                Dim locationText = "Unknown"
                If selectedEvent.ContainsKey("location") AndAlso 
                   TypeOf selectedEvent("location") Is Newtonsoft.Json.Linq.JObject Then
                    Dim locationObj = DirectCast(selectedEvent("location"), Newtonsoft.Json.Linq.JObject)
                    If locationObj.ContainsKey("address") Then
                        locationText = locationObj("address").ToString()
                    End If
                End If
                DirectCast(locationLabel, Label).Text = locationText
            End If
            
            ' Update timezone
            Dim timezoneLabel = FindControlByName(detailsContent, "lblEventTimezone")
            If timezoneLabel IsNot Nothing Then
                DirectCast(timezoneLabel, Label).Text = If(selectedEvent.ContainsKey("timezone"), 
                    selectedEvent("timezone").ToString(), "Unknown")
            End If
            
            ' Update status
            Dim statusLabel = FindControlByName(detailsContent, "lblEventStatus")
            If statusLabel IsNot Nothing Then
                DirectCast(statusLabel, Label).Text = If(selectedEvent.ContainsKey("status"), 
                    selectedEvent("status").ToString(), "Unknown")
            End If
            
            ' Update organizer
            Dim organizerLabel = FindControlByName(detailsContent, "lblEventOrganizer")
            If organizerLabel IsNot Nothing Then
                Dim organizerText = "Unknown"
                If selectedEvent.ContainsKey("organizer") AndAlso 
                   Not String.IsNullOrEmpty(selectedEvent("organizer").ToString()) Then
                    organizerText = selectedEvent("organizer").ToString()
                ElseIf selectedEvent.ContainsKey("author") Then
                    organizerText = selectedEvent("author").ToString()
                End If
                DirectCast(organizerLabel, Label).Text = organizerText
            End If
            
            ' Update event type
            Dim typeLabel = FindControlByName(detailsContent, "lblEventType")
            If typeLabel IsNot Nothing Then
                Dim typeText = If(selectedEvent.ContainsKey("event_type"), 
                    selectedEvent("event_type").ToString(), "Unknown")
                Dim isVirtual = If(selectedEvent.ContainsKey("_virtual") AndAlso 
                    selectedEvent("_virtual").ToString() = "yes", " (Virtual)", "")
                DirectCast(typeLabel, Label).Text = $"{typeText}{isVirtual}"
            End If
            
            LogMessage("Event details updated successfully")
        Catch ex As Exception
            LogMessage("Error updating event details: " & ex.Message)
        End Try
    End Sub
    
    ' Update schedule details when a row is selected in the grid
    Private Sub UpdateScheduleDetails(scheduleId As String)
        Try
            If String.IsNullOrEmpty(scheduleId) OrElse schedulesList Is Nothing Then
                LogMessage("Cannot update schedule details: Invalid schedule ID or schedules list not loaded")
                Return
            End If
            
            ' Find the schedule by ID
            Dim schedule = schedulesList.Find(Function(s) s.ContainsKey("id") AndAlso s("id").ToString() = scheduleId)
            If schedule Is Nothing Then
                LogMessage($"Schedule with ID {scheduleId} not found")
                Return
            End If
            
            LogMessage($"Updating detailed schedule information for schedule ID: {scheduleId}")
            
            ' Find the details panel controls
            If tabPageEntries Is Nothing OrElse tabPageEntries.Controls.Count = 0 Then
                LogMessage("Cannot update schedule details: Tab page or controls not found")
                Return
            End If
            
            ' Get the main container
            Dim mainContainer = TryCast(tabPageEntries.Controls(0), TableLayoutPanel)
            If mainContainer Is Nothing Then
                LogMessage("Cannot update details: Main container not found")
                Return
            End If
            
            ' Get the details panel
            Dim detailsPanel = TryCast(mainContainer.Controls(1), Panel)
            If detailsPanel Is Nothing Then
                LogMessage("Cannot update details: Details panel not found")
                Return
            End If
            
            ' Find the details content panel
            Dim detailsContent = Nothing
            For Each ctrl In detailsPanel.Controls
                If TypeOf ctrl Is TableLayoutPanel Then
                    detailsContent = DirectCast(ctrl, TableLayoutPanel)
                    Exit For
                End If
            Next
            
            If detailsContent Is Nothing Then
                LogMessage("Cannot update details: Details content panel not found")
                Return
            End If
            
            ' Extract schedule information
            Dim title = If(schedule.ContainsKey("program_title"), schedule("program_title").ToString(), "Unknown")
            Dim eventDate = If(schedule.ContainsKey("date"), schedule("date").ToString(), "Unknown date")
            Dim dayName = If(schedule.ContainsKey("day_name") AndAlso 
                Not String.IsNullOrEmpty(schedule("day_name").ToString()), schedule("day_name").ToString(), "")
            
            ' Extract additional data
            Dim startTime As String = ""
            Dim endTime As String = ""
            Dim room As String = ""
            Dim objective As String = ""
            Dim speakerInfo As New System.Text.StringBuilder()
            Dim topicTitle As String = ""
            
            ' Extract data from schedule_slot
            If schedule.ContainsKey("schedule_slot") AndAlso 
                TypeOf schedule("schedule_slot") Is Newtonsoft.Json.Linq.JArray Then
                
                Dim scheduleSlot = DirectCast(schedule("schedule_slot"), Newtonsoft.Json.Linq.JArray)
                If scheduleSlot.Count > 0 Then
                    Dim slotData = scheduleSlot(0).ToString()
                    
                    ' Extract schedule details using regex
                    Dim topicMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_schedule_topic=([^;]+)")
                    If topicMatch.Success Then topicTitle = topicMatch.Groups(1).Value.Trim()
                    
                    Dim startTimeMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_start_time=([^;]+)")
                    If startTimeMatch.Success Then startTime = startTimeMatch.Groups(1).Value.Trim()
                    
                    Dim endTimeMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_end_time=([^;]+)")
                    If endTimeMatch.Success Then endTime = endTimeMatch.Groups(1).Value.Trim()
                    
                    Dim roomMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_room=([^;]+)")
                    If roomMatch.Success Then room = roomMatch.Groups(1).Value.Trim()
                    
                    Dim objectiveMatch = System.Text.RegularExpressions.Regex.Match(slotData, "etn_shedule_objective=([^;]+)")
                    If objectiveMatch.Success Then objective = objectiveMatch.Groups(1).Value.Trim()
                    
                    ' Process speakers
                    Dim speakersMatch = System.Text.RegularExpressions.Regex.Match(slotData, "speakers=([^;]+)")
                    If speakersMatch.Success AndAlso speakersList IsNot Nothing AndAlso speakersList.Count > 0 Then
                        Dim speakersData = speakersMatch.Groups(1).Value.Trim()
                        LogMessage($"Details view - found speakers data: {speakersData}")

                        ' Try different approaches to find speakers
                        ' First, try to parse as JSON array
                        Try
                            Dim speakerIdsArray = JsonConvert.DeserializeObject(Of List(Of Integer))(speakersData)
                            If speakerIdsArray IsNot Nothing AndAlso speakerIdsArray.Count > 0 Then
                                For Each speakerId In speakerIdsArray
                                    ProcessSpeakerForDetail(speakerId.ToString(), speakerInfo)
                                Next
                            End If
                        Catch jsonEx As Exception
                            LogMessage($"Details view - not JSON array: {jsonEx.Message}")
                            
                            ' Try ID=number pattern
                            Dim speakerIds = System.Text.RegularExpressions.Regex.Matches(speakersData, "id=(\d+)")
                            If speakerIds.Count > 0 Then
                                For Each match In speakerIds
                                    Dim speakerId = match.Groups(1).Value
                                    ProcessSpeakerForDetail(speakerId, speakerInfo)
                                Next
                            Else
                                ' Try simple numbers pattern
                                Dim numberMatches = System.Text.RegularExpressions.Regex.Matches(speakersData, "(\d+)")
                                If numberMatches.Count > 0 Then
                                    For Each match In numberMatches
                                        Dim speakerId = match.Groups(1).Value
                                        ProcessSpeakerForDetail(speakerId, speakerInfo)
                                    Next
                                End If
                            End If
                        End Try
                    ElseIf Not speakersMatch.Success Then
                        ' If no speakers in slot, check event level
                        FindEventLevelSpeakers(schedule, speakerInfo)
                    End If
                End If
            End If
            
            ' Update UI with schedule information
            Dim titleLabel = FindControlByName(detailsContent, "lblScheduleTitle")
            If titleLabel IsNot Nothing Then
                DirectCast(titleLabel, Label).Text = If(String.IsNullOrEmpty(topicTitle), title, topicTitle)
            End If
            
            ' Update date and time
            Dim timeLabel = FindControlByName(detailsContent, "lblScheduleTime")
            If timeLabel IsNot Nothing Then
                Dim timeInfo = If(Not String.IsNullOrEmpty(dayName), 
                    $"{eventDate} ({dayName})", eventDate)
                
                If Not String.IsNullOrEmpty(startTime) OrElse Not String.IsNullOrEmpty(endTime) Then
                    timeInfo += $" | {startTime} - {endTime}"
                End If
                
                DirectCast(timeLabel, Label).Text = timeInfo
            End If
            
            ' Update description
            Dim descText = FindControlByName(detailsContent, "txtDescription")
                    If descText IsNot Nothing AndAlso TypeOf descText Is TextBox Then
                        Dim fullDescription As New System.Text.StringBuilder()
                        
                        If Not String.IsNullOrEmpty(objective) Then
                            fullDescription.AppendLine(objective)
                            fullDescription.AppendLine()
                        End If
                        
                If speakerInfo.Length > 0 Then
                    fullDescription.AppendLine("Speakers:")
                        fullDescription.Append(speakerInfo.ToString())
                    End If
                
                DirectCast(descText, TextBox).Text = If(fullDescription.Length > 0, 
                    fullDescription.ToString().Trim(), "No description available")
            End If
            
            LogMessage("Schedule details updated successfully")
        Catch ex As Exception
            LogMessage("Error updating schedule details: " & ex.Message)
        End Try
    End Sub

    ' Helper method to process speakers for the detailed view
    Private Sub ProcessSpeakerForDetail(speakerId As String, speakerInfo As System.Text.StringBuilder)
        Dim speaker = GetSpeakerById(speakerId)
        If speaker IsNot Nothing Then
            speakerInfo.AppendLine("� " & If(speaker.ContainsKey("name"), speaker("name").ToString(), "Unknown"))
            If speaker.ContainsKey("company_name") AndAlso Not String.IsNullOrEmpty(speaker("company_name").ToString()) Then
                speakerInfo.AppendLine("  " & speaker("company_name").ToString())
            ElseIf speaker.ContainsKey("designation") AndAlso Not String.IsNullOrEmpty(speaker("designation").ToString()) Then
                speakerInfo.AppendLine("  " & speaker("designation").ToString())
            End If
            If speaker.ContainsKey("summary") Then
                Dim bio = speaker("summary").ToString()
                bio = System.Text.RegularExpressions.Regex.Replace(bio, "<.*?>", String.Empty)
                If Not String.IsNullOrEmpty(bio) Then
                    speakerInfo.AppendLine("  " & bio)
                End If
            End If
            speakerInfo.AppendLine()
        End If
    End Sub

    ' Helper method to find event-level speakers
    Private Sub FindEventLevelSpeakers(schedule As Dictionary(Of String, Object), speakerInfo As System.Text.StringBuilder)
        ' Try to find event data from the current selected room first
        Dim currentEvent As Dictionary(Of String, Object) = Nothing
        
        ' Check if we have a selected room and events loaded
        If Not String.IsNullOrEmpty(selectedRoom) AndAlso eventsList IsNot Nothing AndAlso eventsList.Count > 0 Then
            currentEvent = eventsList.Find(Function(evt) evt("title").ToString() = selectedRoom)
        End If
        
        ' Check if we found an event and it has speaker data
        If currentEvent IsNot Nothing AndAlso currentEvent.ContainsKey("speaker") AndAlso 
           Not String.IsNullOrEmpty(currentEvent("speaker").ToString()) Then
            
            Dim eventSpeakers = currentEvent("speaker").ToString().Trim()
            LogMessage($"Using event level speakers: {eventSpeakers}")
            
            ' Extract speaker IDs using regex
            Dim speakerMatches = System.Text.RegularExpressions.Regex.Matches(eventSpeakers, "(\d+)")
            If speakerMatches.Count > 0 Then
                Dim speakerNames As New List(Of String)
                
                For Each match In speakerMatches
                    Dim speakerId = match.Groups(1).Value
                    LogMessage($"Event speaker ID: {speakerId}")
                    Dim speaker = GetSpeakerById(speakerId)
                    
                    If speaker IsNot Nothing AndAlso speaker.ContainsKey("name") Then
                        speakerNames.Add(speaker("name").ToString())
                        LogMessage($"Found event speaker: {speaker("name")}")
                    End If
                Next
                
                If speakerNames.Count > 0 Then
                    speakerInfo.AppendLine("Speakers:")
                    speakerInfo.AppendLine(String.Join(", ", speakerNames))
                    LogMessage($"Set event speakers to: {speakerInfo.ToString().Trim()}")
                End If
            End If
        Else
            LogMessage("No event-level speakers found")
        End If
    End Sub
End Class
