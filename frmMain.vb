Imports System.IO

Public Class frmMain
    Inherits System.Windows.Forms.Form

    Friend WithEvents dlgFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents windowsMediaPlayer As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents lblTime As System.Windows.Forms.Label
    Friend WithEvents cmdPause As System.Windows.Forms.Button
    Friend WithEvents cmdPlay As System.Windows.Forms.Button
    Friend WithEvents cmdStop As System.Windows.Forms.Button
    Friend WithEvents tbProgress As System.Windows.Forms.TrackBar
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents cmdPrev As System.Windows.Forms.Button
    Friend WithEvents ScrollingText1 As ScrollingText.ScrollingText

    Dim WithEvents SysVol As New Sound

    Private Playlist_pos As Integer
    Friend WithEvents chkRepeat As System.Windows.Forms.CheckBox
    Friend WithEvents numRepeat As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents chkLoop As System.Windows.Forms.CheckBox


    'we dont want anyone setting the alarm without a valid wav set
    Public wavFileLoaded As Boolean = False

    'Private Declare Function mciGetErrorString _
    'Lib "winmm.dll" Alias "mciGetErrorStringA" _
    '( _
    '  ByVal dwError As Long, _
    '  ByVal lpstrBuffer As String, _
    '  ByVal uLength As Long _
    ') As Integer

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents sfdSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuLoadPre As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSavePre As System.Windows.Forms.MenuItem
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents lblClock As System.Windows.Forms.Label
    Friend WithEvents lblAlarmSetFor As System.Windows.Forms.Label
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ntfSysTray As System.Windows.Forms.NotifyIcon
    Friend WithEvents mnuFileMinTray As System.Windows.Forms.MenuItem
    Friend WithEvents lblAlarmTimeDisp As System.Windows.Forms.Label
    Friend WithEvents lblAlarmTime As System.Windows.Forms.Label
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents radAM As System.Windows.Forms.RadioButton
    Friend WithEvents radPM As System.Windows.Forms.RadioButton
    Friend WithEvents cmbMinute As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbHour As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnTest As System.Windows.Forms.Button
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtSoundFile As System.Windows.Forms.TextBox
    Friend WithEvents ofdBrowse As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnReset As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuFileMinTray = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuFileExit = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuSavePre = New System.Windows.Forms.MenuItem
        Me.mnuLoadPre = New System.Windows.Forms.MenuItem
        Me.lblClock = New System.Windows.Forms.Label
        Me.lblAlarmSetFor = New System.Windows.Forms.Label
        Me.lblAlarmTimeDisp = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ntfSysTray = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.lblAlarmTime = New System.Windows.Forms.Label
        Me.btnApply = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkLoop = New System.Windows.Forms.CheckBox
        Me.chkRepeat = New System.Windows.Forms.CheckBox
        Me.numRepeat = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnReset = New System.Windows.Forms.Button
        Me.radAM = New System.Windows.Forms.RadioButton
        Me.radPM = New System.Windows.Forms.RadioButton
        Me.cmbMinute = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbHour = New System.Windows.Forms.ComboBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btnTest = New System.Windows.Forms.Button
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.txtSoundFile = New System.Windows.Forms.TextBox
        Me.ofdBrowse = New System.Windows.Forms.OpenFileDialog
        Me.btnStop = New System.Windows.Forms.Button
        Me.sfdSave = New System.Windows.Forms.SaveFileDialog
        Me.dlgFile = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.ScrollingText1 = New ScrollingText.ScrollingText
        Me.cmdPrev = New System.Windows.Forms.Button
        Me.btnNext = New System.Windows.Forms.Button
        Me.lblTime = New System.Windows.Forms.Label
        Me.cmdPause = New System.Windows.Forms.Button
        Me.cmdPlay = New System.Windows.Forms.Button
        Me.cmdStop = New System.Windows.Forms.Button
        Me.tbProgress = New System.Windows.Forms.TrackBar
        Me.windowsMediaPlayer = New AxWMPLib.AxWindowsMediaPlayer
        Me.GroupBox2.SuspendLayout()
        CType(Me.numRepeat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.tbProgress, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.windowsMediaPlayer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.MenuItem1})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileMinTray, Me.MenuItem3, Me.mnuFileExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuFileMinTray
        '
        Me.mnuFileMinTray.Index = 0
        Me.mnuFileMinTray.Text = "Minimize to Tray"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "-"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Index = 2
        Me.mnuFileExit.Text = "&Exit"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuSavePre, Me.mnuLoadPre})
        Me.MenuItem1.Text = "Alarm Presets"
        '
        'mnuSavePre
        '
        Me.mnuSavePre.Index = 0
        Me.mnuSavePre.Text = "Save preset"
        '
        'mnuLoadPre
        '
        Me.mnuLoadPre.Index = 1
        Me.mnuLoadPre.Text = "Load preset"
        '
        'lblClock
        '
        Me.lblClock.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.lblClock.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblClock.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblClock.Font = New System.Drawing.Font("Comic Sans MS", 36.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClock.ForeColor = System.Drawing.Color.Lime
        Me.lblClock.Location = New System.Drawing.Point(16, 8)
        Me.lblClock.Name = "lblClock"
        Me.lblClock.Size = New System.Drawing.Size(336, 64)
        Me.lblClock.TabIndex = 0
        Me.lblClock.Text = "88:88 PM"
        Me.lblClock.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAlarmSetFor
        '
        Me.lblAlarmSetFor.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblAlarmSetFor.ForeColor = System.Drawing.Color.Red
        Me.lblAlarmSetFor.Location = New System.Drawing.Point(88, 88)
        Me.lblAlarmSetFor.Name = "lblAlarmSetFor"
        Me.lblAlarmSetFor.Size = New System.Drawing.Size(104, 24)
        Me.lblAlarmSetFor.TabIndex = 1
        Me.lblAlarmSetFor.Text = "ALARM SET FOR:"
        Me.lblAlarmSetFor.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblAlarmTimeDisp
        '
        Me.lblAlarmTimeDisp.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.lblAlarmTimeDisp.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAlarmTimeDisp.ForeColor = System.Drawing.Color.Red
        Me.lblAlarmTimeDisp.Location = New System.Drawing.Point(192, 88)
        Me.lblAlarmTimeDisp.Name = "lblAlarmTimeDisp"
        Me.lblAlarmTimeDisp.Size = New System.Drawing.Size(88, 24)
        Me.lblAlarmTimeDisp.TabIndex = 2
        Me.lblAlarmTimeDisp.Text = "Not Set"
        Me.lblAlarmTimeDisp.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'ntfSysTray
        '
        Me.ntfSysTray.Text = "Alarm Not Set."
        Me.ntfSysTray.Visible = True
        '
        'lblAlarmTime
        '
        Me.lblAlarmTime.Location = New System.Drawing.Point(16, 88)
        Me.lblAlarmTime.Name = "lblAlarmTime"
        Me.lblAlarmTime.Size = New System.Drawing.Size(64, 24)
        Me.lblAlarmTime.TabIndex = 3
        Me.lblAlarmTime.Text = "Real Time (sec)"
        Me.lblAlarmTime.Visible = False
        '
        'btnApply
        '
        Me.btnApply.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnApply.Location = New System.Drawing.Point(264, 240)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(72, 24)
        Me.btnApply.TabIndex = 14
        Me.btnApply.Text = "Apply"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkLoop)
        Me.GroupBox2.Controls.Add(Me.chkRepeat)
        Me.GroupBox2.Controls.Add(Me.numRepeat)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.btnReset)
        Me.GroupBox2.Controls.Add(Me.radAM)
        Me.GroupBox2.Controls.Add(Me.radPM)
        Me.GroupBox2.Controls.Add(Me.cmbMinute)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.cmbHour)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 224)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(336, 80)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Alarm Time"
        '
        'chkLoop
        '
        Me.chkLoop.AutoSize = True
        Me.chkLoop.Location = New System.Drawing.Point(12, 62)
        Me.chkLoop.Name = "chkLoop"
        Me.chkLoop.Size = New System.Drawing.Size(50, 17)
        Me.chkLoop.TabIndex = 10
        Me.chkLoop.Text = "Loop"
        Me.chkLoop.UseVisualStyleBackColor = True
        '
        'chkRepeat
        '
        Me.chkRepeat.AutoSize = True
        Me.chkRepeat.Location = New System.Drawing.Point(12, 43)
        Me.chkRepeat.Name = "chkRepeat"
        Me.chkRepeat.Size = New System.Drawing.Size(15, 14)
        Me.chkRepeat.TabIndex = 9
        Me.chkRepeat.UseVisualStyleBackColor = True
        '
        'numRepeat
        '
        Me.numRepeat.Location = New System.Drawing.Point(103, 40)
        Me.numRepeat.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.numRepeat.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numRepeat.Name = "numRepeat"
        Me.numRepeat.Size = New System.Drawing.Size(40, 20)
        Me.numRepeat.TabIndex = 8
        Me.numRepeat.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 43)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(163, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Repeat Every:                 Minutes"
        '
        'btnReset
        '
        Me.btnReset.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnReset.Location = New System.Drawing.Point(248, 48)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(72, 24)
        Me.btnReset.TabIndex = 6
        Me.btnReset.Text = "Reset"
        '
        'radAM
        '
        Me.radAM.Checked = True
        Me.radAM.Location = New System.Drawing.Point(200, 14)
        Me.radAM.Name = "radAM"
        Me.radAM.Size = New System.Drawing.Size(42, 16)
        Me.radAM.TabIndex = 5
        Me.radAM.TabStop = True
        Me.radAM.Text = "AM"
        '
        'radPM
        '
        Me.radPM.Location = New System.Drawing.Point(200, 30)
        Me.radPM.Name = "radPM"
        Me.radPM.Size = New System.Drawing.Size(42, 16)
        Me.radPM.TabIndex = 4
        Me.radPM.Text = "PM"
        '
        'cmbMinute
        '
        Me.cmbMinute.Items.AddRange(New Object() {"00", "05", "10", "15", "20", "25", "30", "35", "40", "45", "50", "55"})
        Me.cmbMinute.Location = New System.Drawing.Point(144, 17)
        Me.cmbMinute.Name = "cmbMinute"
        Me.cmbMinute.Size = New System.Drawing.Size(40, 21)
        Me.cmbMinute.TabIndex = 3
        Me.cmbMinute.Text = "00"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(96, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Minute:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 19)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Hour:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmbHour
        '
        Me.cmbHour.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"})
        Me.cmbHour.Location = New System.Drawing.Point(48, 17)
        Me.cmbHour.Name = "cmbHour"
        Me.cmbHour.Size = New System.Drawing.Size(40, 21)
        Me.cmbHour.TabIndex = 0
        Me.cmbHour.Text = "12"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btnTest)
        Me.GroupBox1.Controls.Add(Me.btnBrowse)
        Me.GroupBox1.Controls.Add(Me.txtSoundFile)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 120)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 96)
        Me.GroupBox1.TabIndex = 12
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Choose Sound File"
        '
        'btnTest
        '
        Me.btnTest.Enabled = False
        Me.btnTest.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnTest.Location = New System.Drawing.Point(176, 56)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(144, 24)
        Me.btnTest.TabIndex = 2
        Me.btnTest.Text = "Test Sound"
        '
        'btnBrowse
        '
        Me.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnBrowse.Location = New System.Drawing.Point(16, 56)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(144, 24)
        Me.btnBrowse.TabIndex = 1
        Me.btnBrowse.Text = "Browse..."
        '
        'txtSoundFile
        '
        Me.txtSoundFile.Location = New System.Drawing.Point(16, 24)
        Me.txtSoundFile.Name = "txtSoundFile"
        Me.txtSoundFile.Size = New System.Drawing.Size(304, 20)
        Me.txtSoundFile.TabIndex = 0
        Me.txtSoundFile.Text = "Please select a Audio file"
        '
        'btnStop
        '
        Me.btnStop.Enabled = False
        Me.btnStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnStop.Location = New System.Drawing.Point(288, 88)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(64, 24)
        Me.btnStop.TabIndex = 15
        Me.btnStop.Text = "Stop"
        '
        'sfdSave
        '
        Me.sfdSave.FileName = "doc1"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.ScrollingText1)
        Me.GroupBox3.Controls.Add(Me.cmdPrev)
        Me.GroupBox3.Controls.Add(Me.btnNext)
        Me.GroupBox3.Controls.Add(Me.lblTime)
        Me.GroupBox3.Controls.Add(Me.cmdPause)
        Me.GroupBox3.Controls.Add(Me.cmdPlay)
        Me.GroupBox3.Controls.Add(Me.cmdStop)
        Me.GroupBox3.Controls.Add(Me.tbProgress)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 311)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(336, 120)
        Me.GroupBox3.TabIndex = 17
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Sound File Info"
        '
        'ScrollingText1
        '
        Me.ScrollingText1.Location = New System.Drawing.Point(7, 91)
        Me.ScrollingText1.Name = "ScrollingText1"
        Me.ScrollingText1.ScrollingTextColor1 = System.Drawing.Color.Black
        Me.ScrollingText1.ScrollingTextColor2 = System.Drawing.Color.Black
        Me.ScrollingText1.Size = New System.Drawing.Size(313, 23)
        Me.ScrollingText1.TabIndex = 15
        '
        'cmdPrev
        '
        Me.cmdPrev.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.cmdPrev.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPrev.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPrev.Location = New System.Drawing.Point(6, 60)
        Me.cmdPrev.Name = "cmdPrev"
        Me.cmdPrev.Size = New System.Drawing.Size(32, 24)
        Me.cmdPrev.TabIndex = 14
        Me.cmdPrev.Text = "<<"
        Me.cmdPrev.UseVisualStyleBackColor = False
        '
        'btnNext
        '
        Me.btnNext.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.btnNext.Location = New System.Drawing.Point(161, 61)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(32, 24)
        Me.btnNext.TabIndex = 13
        Me.btnNext.Text = ">>"
        Me.btnNext.UseVisualStyleBackColor = False
        '
        'lblTime
        '
        Me.lblTime.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTime.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.lblTime.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTime.Font = New System.Drawing.Font("Courier New", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.lblTime.ForeColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblTime.Location = New System.Drawing.Point(200, 62)
        Me.lblTime.Name = "lblTime"
        Me.lblTime.Size = New System.Drawing.Size(120, 22)
        Me.lblTime.TabIndex = 12
        Me.lblTime.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cmdPause
        '
        Me.cmdPause.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.cmdPause.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPause.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(238, Byte))
        Me.cmdPause.Location = New System.Drawing.Point(123, 61)
        Me.cmdPause.Name = "cmdPause"
        Me.cmdPause.Size = New System.Drawing.Size(32, 24)
        Me.cmdPause.TabIndex = 11
        Me.cmdPause.Text = "||"
        Me.cmdPause.UseVisualStyleBackColor = False
        '
        'cmdPlay
        '
        Me.cmdPlay.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.cmdPlay.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPlay.Font = New System.Drawing.Font("Wingdings 3", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdPlay.Location = New System.Drawing.Point(83, 61)
        Me.cmdPlay.Name = "cmdPlay"
        Me.cmdPlay.Size = New System.Drawing.Size(32, 24)
        Me.cmdPlay.TabIndex = 10
        Me.cmdPlay.Text = "u"
        Me.cmdPlay.UseVisualStyleBackColor = False
        '
        'cmdStop
        '
        Me.cmdStop.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.cmdStop.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdStop.Font = New System.Drawing.Font("Wingdings", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.cmdStop.Location = New System.Drawing.Point(43, 61)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.Size = New System.Drawing.Size(32, 24)
        Me.cmdStop.TabIndex = 9
        Me.cmdStop.Text = "n"
        Me.cmdStop.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.cmdStop.UseVisualStyleBackColor = False
        '
        'tbProgress
        '
        Me.tbProgress.Location = New System.Drawing.Point(6, 15)
        Me.tbProgress.Name = "tbProgress"
        Me.tbProgress.Size = New System.Drawing.Size(314, 45)
        Me.tbProgress.TabIndex = 0
        Me.tbProgress.TickFrequency = 5
        '
        'windowsMediaPlayer
        '
        Me.windowsMediaPlayer.Enabled = True
        Me.windowsMediaPlayer.Location = New System.Drawing.Point(0, 46)
        Me.windowsMediaPlayer.Name = "windowsMediaPlayer"
        Me.windowsMediaPlayer.OcxState = CType(resources.GetObject("windowsMediaPlayer.OcxState"), System.Windows.Forms.AxHost.State)
        Me.windowsMediaPlayer.Size = New System.Drawing.Size(10, 26)
        Me.windowsMediaPlayer.TabIndex = 16
        Me.windowsMediaPlayer.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(365, 429)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.windowsMediaPlayer)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblAlarmTime)
        Me.Controls.Add(Me.lblAlarmTimeDisp)
        Me.Controls.Add(Me.lblAlarmSetFor)
        Me.Controls.Add(Me.lblClock)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "Alarm Clock"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.numRepeat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.tbProgress, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.windowsMediaPlayer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Private MyMedia As New clsMedia

    Private Sub mnuFileExit_Click( _
                                   ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs _
                                 ) Handles mnuFileExit.Click
        'quit the program
        End
    End Sub

    Private Sub Timer1_Tick( _
                             ByVal sender As System.Object, _
                             ByVal e As System.EventArgs _
                           ) Handles Timer1.Tick
        'everything in here controls the clock and the alarm
        'i made two alarm times one without seconds and one with
        'the user doesnt know about the one with seconds tho, its there
        'for more precision, and if it isnt then the wav will play
        'over and over
        'most of this is pretty basic

        Dim datTime As Date = Microsoft.VisualBasic.DateAndTime.TimeOfDay
        Dim intHour, intMinute, intSecond As Integer
        intHour = datTime.Hour
        intMinute = datTime.Minute
        intSecond = datTime.Second
        Dim displayTime, realTime, alarmtime As String

        displayTime = FormatTime(intHour, intMinute, intSecond)

        lblClock.Text = displayTime
        realTime = displayTime
        alarmtime = lblAlarmTime.Text
        If alarmtime = realTime Then
            cmdPause.Enabled = False
            PlayMusic()
        End If

        Dim time As String = ""

        If Not IsNothing(windowsMediaPlayer.Ctlcontrols.currentItem) Then

            ScrollingText1.Text = windowsMediaPlayer.Ctlcontrols.currentItem.name

            On Error Resume Next
            tbProgress.Maximum = windowsMediaPlayer.Ctlcontrols.currentItem.duration
            tbProgress.Value = windowsMediaPlayer.Ctlcontrols.currentPosition
            On Error GoTo 0

            time = windowsMediaPlayer.Ctlcontrols.currentPositionString
            If time = "" Then time = "00:00"
            time += "/"
            time += windowsMediaPlayer.Ctlcontrols.currentItem.durationString

        Else
            time = ""
        End If

        lblTime.Text = time

        Exit Sub

    End Sub

    Private Sub mnuFileMinTray_Click( _
                                      ByVal sender As System.Object, _
                                      ByVal e As System.EventArgs _
                                    ) Handles mnuFileMinTray.Click
        'hide the form and put the icon in the tray
        Me.Hide()
        ntfSysTray.Visible = True
        ntfSysTray.Icon = Me.Icon
    End Sub

    Private Sub ntfSysTray_DoubleClick( _
                                        ByVal sender As Object, _
                                        ByVal e As System.EventArgs _
                                      ) Handles ntfSysTray.DoubleClick
        'remove the icon form the tray and show the form
        ntfSysTray.Visible = False
        Me.Show()
    End Sub

    Private Sub btnTest_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles btnTest.Click
        'simply play the file in the textbox

        PlayMusic()

    End Sub

    Private Sub btnBrowse_Click( _
                                  ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs _
                               ) Handles btnBrowse.Click

        dlgFile.ShowDialog()
        If dlgFile.FileName <> "" Then
            txtSoundFile.Text = dlgFile.FileName

            ' Set autostart to false
            windowsMediaPlayer.settings.autoStart = False

            ' Assign path to URL property.
            windowsMediaPlayer.URL = txtSoundFile.Text

            wavFileLoaded = True
            btnTest.Enabled = True

        Else
            btnTest.Enabled = False
            wavFileLoaded = False
        End If



    End Sub

    Private Sub btnApply_Click( _
                                ByVal sender As System.Object, _
                                ByVal e As System.EventArgs _
                              ) Handles btnApply.Click

        ApplySettings()

    End Sub

    Private Sub btnStop_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles btnStop.Click

        StopOrSnooze()

    End Sub

    Private Sub btnReset_Click( _
                                ByVal sender As System.Object, _
                                ByVal e As System.EventArgs _
                              ) Handles btnReset.Click

        lblAlarmTime.Text = "Not Set"
        lblAlarmTimeDisp.Text = "Not Set"
        ntfSysTray.Text = "Alarm Not Set."


    End Sub


    Private Sub mnuSavePre_Click( _
                                  ByVal sender As System.Object, _
                                  ByVal e As System.EventArgs _
                                ) Handles mnuSavePre.Click
        'write the alarm preset in the following format:
        'date created
        '
        'hour
        'minute
        ' AM/PM
        'sound file

        sfdSave.CheckPathExists = True
        sfdSave.Filter = "Preset Files (*.pre)|*.pre"
        sfdSave.Title = "Save Preset..."
        sfdSave.InitialDirectory = Application.StartupPath
        If sfdSave.ShowDialog = DialogResult.OK Then
            Dim dateCreated As Date = DateTime.Today    'store today's date

            'writing files is very simple in .net
            '1. store the save boxes filename in a fileinfo
            '2. create a streamwriter variable
            '3. use writeline to write whatever is in the brackets to the file
            '4. close the file
            Dim fileName As FileInfo = New FileInfo(sfdSave.FileName)
            Dim writeFile As StreamWriter = fileName.CreateText
            writeFile.WriteLine("Created on: " & dateCreated)
            writeFile.WriteLine("")
            writeFile.WriteLine(cmbHour.Top)
            writeFile.WriteLine(cmbMinute.Text)
            If radAM.Checked Then
                writeFile.WriteLine(" AM")
            Else
                writeFile.WriteLine(" PM")
            End If
            writeFile.WriteLine(txtSoundFile.Text)
            writeFile.Close()
            'notify the user that the save was successful
            MsgBox("Preset saved succesfully to " & sfdSave.FileName & ".", MsgBoxStyle.Information, "Save Succes")
        End If
    End Sub

    Private Sub mnuLoadPre_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLoadPre.Click
        'load the file line by line then put each value in the respective object
        'file format:
        'date created /dont read this, just for reference
        '
        'hour
        'minute
        ' AM/PM
        'sound file

        Dim ofdLoad As New OpenFileDialog()
        ofdLoad.CheckFileExists = True
        ofdLoad.CheckPathExists = True
        ofdLoad.Filter = "Preset files (*.pre)|*.pre"
        ofdLoad.Title = "Load Preset..."
        ofdLoad.InitialDirectory = Application.StartupPath
        If ofdLoad.ShowDialog = DialogResult.OK Then
            'reading files is just as easy as writing
            '1. create fileinfo variable
            '2. create a streamreader variable
            '3. store readline in a variable or object
            Dim fileName As FileInfo = New FileInfo(ofdLoad.FileName)
            Dim readFile As StreamReader = fileName.OpenText
            'the first two lines in the preset files are not needed for reading
            'so we can skip over them
            readFile.ReadLine()
            readFile.ReadLine()
            'read the next line and put it in cmbHour.text
            cmbHour.Text = readFile.ReadLine
            cmbMinute.Text = readFile.ReadLine
            If readFile.ReadLine = " AM" Then
                radAM.Checked = True
            Else
                radPM.Checked = True
            End If
            txtSoundFile.Text = readFile.ReadLine
            wavFileLoaded = True
            btnTest.Enabled = True
            'call the apply button click event, no sense in writing it all again
            btnApply_Click(mnuLoadPre, e)
        End If
    End Sub

    Private Sub Form1_Load( _
                            ByVal sender As Object, _
                            ByVal e As System.EventArgs _
                          ) Handles Me.Load


    End Sub


    Private Sub cmdStop_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles cmdStop.Click


        StopOrSnooze()

    End Sub

    Private Sub cmdPlay_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles cmdPlay.Click

        cmdPause.Enabled = True
        PlayMusic()

    End Sub

    Private Sub cmdPause_Click( _
                                ByVal sender As System.Object, _
                                ByVal e As System.EventArgs _
                              ) Handles cmdPause.Click


        windowsMediaPlayer.Ctlcontrols.pause()
    End Sub

    Private Sub tbProgress_Scroll( _
                                   ByVal sender As System.Object, _
                                   ByVal e As System.EventArgs _
                                 ) Handles tbProgress.Scroll
        windowsMediaPlayer.Ctlcontrols.currentPosition = tbProgress.Value
    End Sub

    Private Sub btnNext_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles btnNext.Click
        windowsMediaPlayer.Ctlcontrols.next()
    End Sub

    Private Sub cmdPrev_Click( _
                               ByVal sender As System.Object, _
                               ByVal e As System.EventArgs _
                             ) Handles cmdPrev.Click
        windowsMediaPlayer.Ctlcontrols.previous()
    End Sub

    Sub PlayMusic()

        Dim sound As String = txtSoundFile.Text
        If ntfSysTray.Visible = True Then
            ntfSysTray.Visible = False
            Me.Show()
        End If

        windowsMediaPlayer.Ctlcontrols.play()

        btnStop.Enabled = True
        cmdStop.Enabled = True

        ApplySettings(True)

    End Sub

    Sub SetAlarm( _
                  ByVal Hour As String, _
                  ByVal Minute As String, _
                  ByVal Morning As Boolean, _
                  Optional ByVal silent As Boolean = False _
                )

        Dim alarmTimeDisp, alarmTime As String

        If Not Morning Then
            If Hour < 12 Then
                Hour = Hour + 12
                'Else
                '    Hour = 0
            End If
        End If

        If wavFileLoaded = True Then
            'this sets the alarm
            'i think its pretty easy to understand

            alarmTimeDisp = FormatTime(Hour, Minute, ShowSecond:=False)
            lblAlarmTimeDisp.Text = alarmTimeDisp
            alarmTime = FormatTime(Hour, Minute, "00")
            lblAlarmTime.Text = alarmTime
            ntfSysTray.Text = "Wake at " & alarmTimeDisp & "." & vbCrLf & _
                              "Play " & txtSoundFile.Text

            If Not silent Then
                'notify the user that the operation was a success
                MsgBox( _
                        "Alarm succesfully set for " & alarmTimeDisp, _
                        MsgBoxStyle.Information, _
                        "Alarm Success" _
                      )
            End If
        Else
            If Not silent Then
                MsgBox( _
                        "You must enter a valid audio file.", _
                        MsgBoxStyle.Exclamation, _
                        "Invalid Audio File" _
                      )
            End If
        End If

    End Sub

    Sub SetRepeatAlarm()

        Dim alarmTime As String
        Dim clock() As String
        Dim hour, minute, second As Integer
        Dim morning As Boolean

        Dim temp As DateTime = DateTime.Now

        temp = temp.AddMinutes(numRepeat.Value)

        hour = temp.Hour
        minute = temp.Minute
        second = temp.Second

        alarmTime = FormatTime(hour, minute, second)
        clock = Split(alarmTime, " ")

        Select Case clock(1)
            Case "AM"
                morning = True
            Case "PM"
                morning = False
        End Select

        SetAlarm(hour, minute, morning, True)

    End Sub

    Function FormatTime( _
                         ByVal intHour As Integer, _
                         ByVal intMinute As Integer, _
                         Optional ByVal intSecond As Integer = 0, _
                         Optional ByVal ShowSecond As Boolean = True _
                       ) As String


        Dim strHour, strMinute, strSecond, strAmPm As String
        Dim displayTime As String

        If intMinute < 10 Then
            strMinute = 0 & intMinute
        Else
            strMinute = intMinute
        End If

        If intHour > 12 Then
            intHour -= 12
            strAmPm = " PM"
        ElseIf intHour = 0 Then
            intHour = 12
            strAmPm = " AM"
        ElseIf intHour = 12 Then
            strAmPm = " PM"
        Else
            strAmPm = " AM"
        End If

        If intSecond < 10 Then
            strSecond = 0 & intSecond
        Else
            strSecond = intSecond
        End If

        strHour = intHour

        displayTime = strHour & ":" & strMinute

        If ShowSecond Then
            displayTime += ":" & strSecond
        End If

        displayTime += strAmPm

        Return displayTime

    End Function

    Private Sub chkRepeat_CheckedChanged( _
                                          ByVal sender As System.Object, _
                                          ByVal e As System.EventArgs _
                                        ) Handles chkRepeat.CheckedChanged

        If CType(sender, CheckBox).Checked Then
            btnStop.Text = "Snooze" '& " " & "(" & numRepeat.Value & ")"
        Else
            btnStop.Text = "Stop"
        End If

    End Sub

    Private Sub chkLoop_CheckedChanged( _
                                        ByVal sender As System.Object, _
                                        ByVal e As System.EventArgs _
                                      ) Handles chkLoop.CheckedChanged

        If chkLoop.Checked Then
            windowsMediaPlayer.settings.setMode("loop", True)
        Else
            windowsMediaPlayer.settings.setMode("loop", False)
        End If

    End Sub

    Sub StopOrSnooze()


        'play a non-existing wav file to stop the previous playing wav
        windowsMediaPlayer.Ctlcontrols.stop()

        If chkRepeat.Checked Then
            SetRepeatAlarm()
        End If

        btnStop.Enabled = False
        cmdStop.Enabled = False
        cmdPause.Enabled = True

    End Sub

    Sub ApplySettings(Optional ByVal silent As Boolean = False)


        Dim hour, minute As String
        Dim morning As Boolean

        hour = cmbHour.Text
        minute = cmbMinute.Text

        If radAM.Checked = True Then
            morning = True
        Else
            morning = False
        End If

        SetAlarm(hour, minute, morning, silent)


    End Sub

End Class
