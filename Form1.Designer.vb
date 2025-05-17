Partial Class Form1
    Inherits System.Windows.Forms.Form

    Private WithEvents buttonConnect As System.Windows.Forms.Button
    Private WithEvents buttonSelectFolder As System.Windows.Forms.Button
    Private WithEvents buttonSyncFiles As System.Windows.Forms.Button
    Private WithEvents comboBoxRooms As System.Windows.Forms.ComboBox
    Private labelStatus As System.Windows.Forms.Label
    Private labelFolder As System.Windows.Forms.Label
    Private textBoxLogs As System.Windows.Forms.TextBox

    Private Sub InitializeComponent()
        buttonConnect = New Button()
        buttonSelectFolder = New Button()
        buttonSyncFiles = New Button()
        comboBoxRooms = New ComboBox()
        labelStatus = New Label()
        labelFolder = New Label()
        textBoxLogs = New TextBox()
        SuspendLayout()
        ' 
        ' buttonConnect
        ' 
        buttonConnect.Location = New Point(20, 20)
        buttonConnect.Name = "buttonConnect"
        buttonConnect.Size = New Size(75, 23)
        buttonConnect.TabIndex = 0
        buttonConnect.Text = "Verbinden"
        ' 
        ' buttonSelectFolder
        ' 
        buttonSelectFolder.Location = New Point(20, 60)
        buttonSelectFolder.Name = "buttonSelectFolder"
        buttonSelectFolder.Size = New Size(75, 23)
        buttonSelectFolder.TabIndex = 1
        buttonSelectFolder.Text = "Ordner wählen"
        ' 
        ' buttonSyncFiles
        ' 
        buttonSyncFiles.Location = New Point(20, 100)
        buttonSyncFiles.Name = "buttonSyncFiles"
        buttonSyncFiles.Size = New Size(75, 23)
        buttonSyncFiles.TabIndex = 2
        buttonSyncFiles.Text = "Dateien synchronisieren"
        ' 
        ' comboBoxRooms
        ' 
        comboBoxRooms.Location = New Point(20, 140)
        comboBoxRooms.Name = "comboBoxRooms"
        comboBoxRooms.Size = New Size(121, 28)
        comboBoxRooms.TabIndex = 3
        ' 
        ' labelStatus
        ' 
        labelStatus.Location = New Point(20, 180)
        labelStatus.Name = "labelStatus"
        labelStatus.Size = New Size(100, 23)
        labelStatus.TabIndex = 4
        labelStatus.Text = "Status: Nicht verbunden"
        ' 
        ' labelFolder
        ' 
        labelFolder.Location = New Point(20, 220)
        labelFolder.Name = "labelFolder"
        labelFolder.Size = New Size(100, 23)
        labelFolder.TabIndex = 5
        labelFolder.Text = "Speicherort: Nicht ausgewählt"
        labelFolder.Visible = False
        ' 
        ' textBoxLogs
        ' 
        textBoxLogs.Location = New Point(20, 260)
        textBoxLogs.Multiline = True
        textBoxLogs.Name = "textBoxLogs"
        textBoxLogs.ScrollBars = ScrollBars.Vertical
        textBoxLogs.Size = New Size(400, 200)
        textBoxLogs.TabIndex = 6
        ' 
        ' Form1
        ' 
        ClientSize = New Size(561, 480)
        Controls.Add(buttonConnect)
        Controls.Add(buttonSelectFolder)
        Controls.Add(buttonSyncFiles)
        Controls.Add(comboBoxRooms)
        Controls.Add(labelStatus)
        Controls.Add(labelFolder)
        Controls.Add(textBoxLogs)
        Name = "Form1"
        Text = "Mediatech Sync"
        ResumeLayout(False)
        PerformLayout()
    End Sub
End Class
