Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
Imports DLL_SetupAndCalibrate
Imports System.Threading
Imports Microsoft.win32
Imports Microsoft.Office.Interop
Imports System.Messaging

Public Class FormProduction

    Inherits System.Windows.Forms.Form

    'this is to set the potlife timer on or off
    Public m_PotLifeOn As Boolean = False

    'potlife time tracking flags
    Public ElapsedTimeSincePotLifeStart As Integer = 0
    Public RemainingTimeOfPotLifeStart As Date

    'autopurging time tracking flags
    Public CurrentTime As Date
    Public StartTimeOfMachineIdle As Date = Now
    Public ElapsedTimeSinceMachineIdle As Integer = 0 'mins

    'check whether the machine has been running in continuous mode 
    Public HasBeenRunning As Boolean = False
    Private Production_info(50) As String

    'constants
    Public Const tickToMinute As Double = 0.0000000016666666  '0.0000001/60.0

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'TimerMonitor.Enabled = True
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        TimerMonitor.Enabled = False
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
    Friend WithEvents TBBStart As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBPause As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBStop As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageListOper As System.Windows.Forms.ImageList
    Friend WithEvents ImageListPotEtc As System.Windows.Forms.ImageList
    Public WithEvents MainMenuProduction As System.Windows.Forms.MainMenu
    Friend WithEvents ImageListGeneralTools As System.Windows.Forms.ImageList
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBoxRobotPos As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PanelProDownTimeInfor As System.Windows.Forms.Panel
    Friend WithEvents RichTextBoxNote As System.Windows.Forms.RichTextBox
    Friend WithEvents CheckBoxPotOn As System.Windows.Forms.CheckBox
    Friend WithEvents DoorLock As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonPotReset As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonOpenFile As System.Windows.Forms.Button
    Friend WithEvents TextBoxFilename As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButtonNdlCalib As System.Windows.Forms.Button
    Friend WithEvents ButtonVolCalib As System.Windows.Forms.Button
    Friend WithEvents ButtonClean As System.Windows.Forms.Button
    Friend WithEvents ButtonPurge As System.Windows.Forms.Button
    Friend WithEvents ButtonChgSyringe As System.Windows.Forms.Button
    Friend WithEvents ButtonHome As System.Windows.Forms.Button
    Friend WithEvents LabelMessege As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents ButtonCV_Prod_Release As System.Windows.Forms.Button
    Friend WithEvents ButtonCV_Prod_Retrieve As System.Windows.Forms.Button
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents TBOperation As System.Windows.Forms.ToolBar
    Friend WithEvents PBYellow As System.Windows.Forms.PictureBox
    Friend WithEvents PBRed As System.Windows.Forms.PictureBox
    Friend WithEvents PBGreen As System.Windows.Forms.PictureBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ImageListOperation As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TimerMonitor As System.Windows.Forms.Timer
    Friend WithEvents ButtonStartFirstStage As System.Windows.Forms.Button
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Needle As System.Windows.Forms.Label
    Friend WithEvents Syringe As System.Windows.Forms.Label
    Friend WithEvents Station1 As System.Windows.Forms.Label
    Friend WithEvents Station2 As System.Windows.Forms.Label
    Friend WithEvents Station3 As System.Windows.Forms.Label
    Friend WithEvents ContinuousMode As System.Windows.Forms.CheckBox
    Friend WithEvents Station1Picture As System.Windows.Forms.PictureBox
    Friend WithEvents Station2Picture As System.Windows.Forms.PictureBox
    Friend WithEvents Station3Picture As System.Windows.Forms.PictureBox
    Friend WithEvents SyringePicture As System.Windows.Forms.PictureBox
    Friend WithEvents NeedlePicture As System.Windows.Forms.PictureBox
    Friend WithEvents Station3Label As System.Windows.Forms.Label
    Friend WithEvents Station2Label As System.Windows.Forms.Label
    Friend WithEvents Station1Label As System.Windows.Forms.Label
    Friend WithEvents NeedleLabel As System.Windows.Forms.Label
    Friend WithEvents SyringeLabel As System.Windows.Forms.Label
    Friend WithEvents ConveyorBox As System.Windows.Forms.GroupBox
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents HeaterBox As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents ResetPLCLogic As System.Windows.Forms.Button
    Friend WithEvents btExit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormProduction))
        Me.TBBStart = New System.Windows.Forms.ToolBarButton
        Me.TBBPause = New System.Windows.Forms.ToolBarButton
        Me.TBBStop = New System.Windows.Forms.ToolBarButton
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListPotEtc = New System.Windows.Forms.ImageList(Me.components)
        Me.MainMenuProduction = New System.Windows.Forms.MainMenu
        Me.ImageListGeneralTools = New System.Windows.Forms.ImageList(Me.components)
        Me.Label7 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.TextBoxRobotPos = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.ConveyorBox = New System.Windows.Forms.GroupBox
        Me.ResetPLCLogic = New System.Windows.Forms.Button
        Me.ButtonStartFirstStage = New System.Windows.Forms.Button
        Me.ContinuousMode = New System.Windows.Forms.CheckBox
        Me.ButtonCV_Prod_Release = New System.Windows.Forms.Button
        Me.ButtonCV_Prod_Retrieve = New System.Windows.Forms.Button
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.TBOperation = New System.Windows.Forms.ToolBar
        Me.PBRed = New System.Windows.Forms.PictureBox
        Me.PBGreen = New System.Windows.Forms.PictureBox
        Me.PBYellow = New System.Windows.Forms.PictureBox
        Me.HeaterBox = New System.Windows.Forms.GroupBox
        Me.Station3Picture = New System.Windows.Forms.PictureBox
        Me.Syringe = New System.Windows.Forms.Label
        Me.Station1 = New System.Windows.Forms.Label
        Me.Station2 = New System.Windows.Forms.Label
        Me.Station3 = New System.Windows.Forms.Label
        Me.SyringePicture = New System.Windows.Forms.PictureBox
        Me.NeedlePicture = New System.Windows.Forms.PictureBox
        Me.Station3Label = New System.Windows.Forms.Label
        Me.Station2Label = New System.Windows.Forms.Label
        Me.Needle = New System.Windows.Forms.Label
        Me.Station1Picture = New System.Windows.Forms.PictureBox
        Me.Station1Label = New System.Windows.Forms.Label
        Me.Station2Picture = New System.Windows.Forms.PictureBox
        Me.NeedleLabel = New System.Windows.Forms.Label
        Me.SyringeLabel = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.PanelProDownTimeInfor = New System.Windows.Forms.Panel
        Me.RichTextBoxNote = New System.Windows.Forms.RichTextBox
        Me.CheckBoxPotOn = New System.Windows.Forms.CheckBox
        Me.DoorLock = New System.Windows.Forms.CheckBox
        Me.ButtonPotReset = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonOpenFile = New System.Windows.Forms.Button
        Me.TextBoxFilename = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonNdlCalib = New System.Windows.Forms.Button
        Me.ButtonClean = New System.Windows.Forms.Button
        Me.ButtonPurge = New System.Windows.Forms.Button
        Me.ButtonChgSyringe = New System.Windows.Forms.Button
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.ButtonVolCalib = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ImageListOperation = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TimerMonitor = New System.Windows.Forms.Timer(Me.components)
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.btExit = New System.Windows.Forms.Button
        Me.Panel2.SuspendLayout()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel5.SuspendLayout()
        Me.ConveyorBox.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.HeaterBox.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.PanelProDownTimeInfor.SuspendLayout()
        Me.SuspendLayout()
        '
        'TBBStart
        '
        Me.TBBStart.Enabled = False
        Me.TBBStart.ImageIndex = 0
        '
        'TBBPause
        '
        Me.TBBPause.Enabled = False
        Me.TBBPause.ImageIndex = 1
        '
        'TBBStop
        '
        Me.TBBStop.Enabled = False
        Me.TBBStop.ImageIndex = 2
        '
        'ImageListOper
        '
        Me.ImageListOper.ImageSize = New System.Drawing.Size(48, 48)
        Me.ImageListOper.ImageStream = CType(resources.GetObject("ImageListOper.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOper.TransparentColor = System.Drawing.Color.White
        '
        'ImageListPotEtc
        '
        Me.ImageListPotEtc.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListPotEtc.ImageStream = CType(resources.GetObject("ImageListPotEtc.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListPotEtc.TransparentColor = System.Drawing.Color.Transparent
        '
        'ImageListGeneralTools
        '
        Me.ImageListGeneralTools.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListGeneralTools.ImageStream = CType(resources.GetObject("ImageListGeneralTools.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListGeneralTools.TransparentColor = System.Drawing.Color.White
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(304, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 23)
        Me.Label7.TabIndex = 134
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.ValueBrightness)
        Me.Panel2.Controls.Add(Me.TextBoxRobotPos)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 960)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(768, 32)
        Me.Panel2.TabIndex = 4
        '
        'ValueBrightness
        '
        Me.ValueBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ValueBrightness.Location = New System.Drawing.Point(80, 5)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.Size = New System.Drawing.Size(45, 21)
        Me.ValueBrightness.TabIndex = 77
        Me.ValueBrightness.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'TextBoxRobotPos
        '
        Me.TextBoxRobotPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotPos.Location = New System.Drawing.Point(504, 3)
        Me.TextBoxRobotPos.Name = "TextBoxRobotPos"
        Me.TextBoxRobotPos.ReadOnly = True
        Me.TextBoxRobotPos.Size = New System.Drawing.Size(256, 21)
        Me.TextBoxRobotPos.TabIndex = 8
        Me.TextBoxRobotPos.Text = "X: -100.000,   Y: -100.000,  Z: -100.000"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(450, 5)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Brightness"
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.Panel5.Controls.Add(Me.btExit)
        Me.Panel5.Controls.Add(Me.ConveyorBox)
        Me.Panel5.Controls.Add(Me.Panel6)
        Me.Panel5.Controls.Add(Me.HeaterBox)
        Me.Panel5.Controls.Add(Me.TextBox1)
        Me.Panel5.Controls.Add(Me.TextBox2)
        Me.Panel5.Controls.Add(Me.GroupBox1)
        Me.Panel5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel5.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Panel5.Location = New System.Drawing.Point(768, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(504, 992)
        Me.Panel5.TabIndex = 0
        '
        'ConveyorBox
        '
        Me.ConveyorBox.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.ConveyorBox.Controls.Add(Me.ResetPLCLogic)
        Me.ConveyorBox.Controls.Add(Me.ButtonStartFirstStage)
        Me.ConveyorBox.Controls.Add(Me.ContinuousMode)
        Me.ConveyorBox.Controls.Add(Me.ButtonCV_Prod_Release)
        Me.ConveyorBox.Controls.Add(Me.ButtonCV_Prod_Retrieve)
        Me.ConveyorBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ConveyorBox.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.ConveyorBox.Location = New System.Drawing.Point(16, 200)
        Me.ConveyorBox.Name = "ConveyorBox"
        Me.ConveyorBox.Size = New System.Drawing.Size(480, 184)
        Me.ConveyorBox.TabIndex = 130
        Me.ConveyorBox.TabStop = False
        Me.ConveyorBox.Text = "Conveyor Operation"
        '
        'ResetPLCLogic
        '
        Me.ResetPLCLogic.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResetPLCLogic.Location = New System.Drawing.Point(140, 120)
        Me.ResetPLCLogic.Name = "ResetPLCLogic"
        Me.ResetPLCLogic.Size = New System.Drawing.Size(200, 48)
        Me.ResetPLCLogic.TabIndex = 143
        Me.ResetPLCLogic.Text = "Reset PLC Logic"
        '
        'ButtonStartFirstStage
        '
        Me.ButtonStartFirstStage.Location = New System.Drawing.Point(56, 80)
        Me.ButtonStartFirstStage.Name = "ButtonStartFirstStage"
        Me.ButtonStartFirstStage.Size = New System.Drawing.Size(120, 23)
        Me.ButtonStartFirstStage.TabIndex = 142
        Me.ButtonStartFirstStage.Text = "Start"
        '
        'ContinuousMode
        '
        Me.ContinuousMode.AutoCheck = False
        Me.ContinuousMode.Location = New System.Drawing.Point(40, 48)
        Me.ContinuousMode.Name = "ContinuousMode"
        Me.ContinuousMode.Size = New System.Drawing.Size(160, 24)
        Me.ContinuousMode.TabIndex = 141
        Me.ContinuousMode.Text = "Continuous Run"
        '
        'ButtonCV_Prod_Release
        '
        Me.ButtonCV_Prod_Release.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Release.Image = CType(resources.GetObject("ButtonCV_Prod_Release.Image"), System.Drawing.Image)
        Me.ButtonCV_Prod_Release.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Release.Location = New System.Drawing.Point(344, 48)
        Me.ButtonCV_Prod_Release.Name = "ButtonCV_Prod_Release"
        Me.ButtonCV_Prod_Release.Size = New System.Drawing.Size(88, 56)
        Me.ButtonCV_Prod_Release.TabIndex = 140
        Me.ButtonCV_Prod_Release.Text = "Release"
        Me.ButtonCV_Prod_Release.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonCV_Prod_Retrieve
        '
        Me.ButtonCV_Prod_Retrieve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Retrieve.Image = CType(resources.GetObject("ButtonCV_Prod_Retrieve.Image"), System.Drawing.Image)
        Me.ButtonCV_Prod_Retrieve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Retrieve.Location = New System.Drawing.Point(224, 48)
        Me.ButtonCV_Prod_Retrieve.Name = "ButtonCV_Prod_Retrieve"
        Me.ButtonCV_Prod_Retrieve.Size = New System.Drawing.Size(88, 56)
        Me.ButtonCV_Prod_Retrieve.TabIndex = 139
        Me.ButtonCV_Prod_Retrieve.Text = "Retrieve"
        Me.ButtonCV_Prod_Retrieve.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Panel6.Controls.Add(Me.TBOperation)
        Me.Panel6.Controls.Add(Me.PBRed)
        Me.Panel6.Controls.Add(Me.PBGreen)
        Me.Panel6.Controls.Add(Me.PBYellow)
        Me.Panel6.Location = New System.Drawing.Point(104, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(304, 184)
        Me.Panel6.TabIndex = 0
        '
        'TBOperation
        '
        Me.TBOperation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TBOperation.BackColor = System.Drawing.SystemColors.Menu
        Me.TBOperation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TBOperation.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBStart, Me.TBBPause, Me.TBBStop})
        Me.TBOperation.ButtonSize = New System.Drawing.Size(90, 60)
        Me.TBOperation.Dock = System.Windows.Forms.DockStyle.None
        Me.TBOperation.DropDownArrows = True
        Me.TBOperation.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBOperation.ForeColor = System.Drawing.Color.Black
        Me.TBOperation.ImageList = Me.ImageListOper
        Me.TBOperation.Location = New System.Drawing.Point(16, 96)
        Me.TBOperation.Name = "TBOperation"
        Me.TBOperation.ShowToolTips = True
        Me.TBOperation.Size = New System.Drawing.Size(274, 68)
        Me.TBOperation.TabIndex = 128
        '
        'PBRed
        '
        Me.PBRed.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBRed.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBRed.Image = CType(resources.GetObject("PBRed.Image"), System.Drawing.Image)
        Me.PBRed.Location = New System.Drawing.Point(13, 11)
        Me.PBRed.Name = "PBRed"
        Me.PBRed.Size = New System.Drawing.Size(279, 75)
        Me.PBRed.TabIndex = 129
        Me.PBRed.TabStop = False
        '
        'PBGreen
        '
        Me.PBGreen.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBGreen.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBGreen.Image = CType(resources.GetObject("PBGreen.Image"), System.Drawing.Image)
        Me.PBGreen.Location = New System.Drawing.Point(13, 11)
        Me.PBGreen.Name = "PBGreen"
        Me.PBGreen.Size = New System.Drawing.Size(279, 75)
        Me.PBGreen.TabIndex = 130
        Me.PBGreen.TabStop = False
        '
        'PBYellow
        '
        Me.PBYellow.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBYellow.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBYellow.Image = CType(resources.GetObject("PBYellow.Image"), System.Drawing.Image)
        Me.PBYellow.Location = New System.Drawing.Point(13, 11)
        Me.PBYellow.Name = "PBYellow"
        Me.PBYellow.Size = New System.Drawing.Size(279, 75)
        Me.PBYellow.TabIndex = 131
        Me.PBYellow.TabStop = False
        '
        'HeaterBox
        '
        Me.HeaterBox.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.HeaterBox.Controls.Add(Me.Station3Picture)
        Me.HeaterBox.Controls.Add(Me.Syringe)
        Me.HeaterBox.Controls.Add(Me.Station1)
        Me.HeaterBox.Controls.Add(Me.Station2)
        Me.HeaterBox.Controls.Add(Me.Station3)
        Me.HeaterBox.Controls.Add(Me.SyringePicture)
        Me.HeaterBox.Controls.Add(Me.NeedlePicture)
        Me.HeaterBox.Controls.Add(Me.Station3Label)
        Me.HeaterBox.Controls.Add(Me.Station2Label)
        Me.HeaterBox.Controls.Add(Me.Needle)
        Me.HeaterBox.Controls.Add(Me.Station1Picture)
        Me.HeaterBox.Controls.Add(Me.Station1Label)
        Me.HeaterBox.Controls.Add(Me.Station2Picture)
        Me.HeaterBox.Controls.Add(Me.NeedleLabel)
        Me.HeaterBox.Controls.Add(Me.SyringeLabel)
        Me.HeaterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.HeaterBox.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.HeaterBox.Location = New System.Drawing.Point(16, 456)
        Me.HeaterBox.Name = "HeaterBox"
        Me.HeaterBox.Size = New System.Drawing.Size(480, 224)
        Me.HeaterBox.TabIndex = 130
        Me.HeaterBox.TabStop = False
        Me.HeaterBox.Text = "Thermal Readings"
        '
        'Station3Picture
        '
        Me.Station3Picture.Image = CType(resources.GetObject("Station3Picture.Image"), System.Drawing.Image)
        Me.Station3Picture.Location = New System.Drawing.Point(320, 92)
        Me.Station3Picture.Name = "Station3Picture"
        Me.Station3Picture.Size = New System.Drawing.Size(112, 64)
        Me.Station3Picture.TabIndex = 138
        Me.Station3Picture.TabStop = False
        Me.Station3Picture.Visible = False
        '
        'Syringe
        '
        Me.Syringe.ForeColor = System.Drawing.Color.Red
        Me.Syringe.Location = New System.Drawing.Point(324, 32)
        Me.Syringe.Name = "Syringe"
        Me.Syringe.Size = New System.Drawing.Size(56, 24)
        Me.Syringe.TabIndex = 143
        Me.Syringe.Text = "50 oC"
        Me.Syringe.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Syringe.Visible = False
        '
        'Station1
        '
        Me.Station1.ForeColor = System.Drawing.Color.Red
        Me.Station1.Location = New System.Drawing.Point(72, 180)
        Me.Station1.Name = "Station1"
        Me.Station1.Size = New System.Drawing.Size(56, 24)
        Me.Station1.TabIndex = 143
        Me.Station1.Text = "50 oC"
        Me.Station1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Station1.Visible = False
        '
        'Station2
        '
        Me.Station2.ForeColor = System.Drawing.Color.Red
        Me.Station2.Location = New System.Drawing.Point(208, 180)
        Me.Station2.Name = "Station2"
        Me.Station2.Size = New System.Drawing.Size(56, 24)
        Me.Station2.TabIndex = 143
        Me.Station2.Text = "50 oC"
        Me.Station2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Station2.Visible = False
        '
        'Station3
        '
        Me.Station3.ForeColor = System.Drawing.Color.Red
        Me.Station3.Location = New System.Drawing.Point(352, 180)
        Me.Station3.Name = "Station3"
        Me.Station3.Size = New System.Drawing.Size(56, 24)
        Me.Station3.TabIndex = 143
        Me.Station3.Text = "50 oC"
        Me.Station3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Station3.Visible = False
        '
        'SyringePicture
        '
        Me.SyringePicture.Image = CType(resources.GetObject("SyringePicture.Image"), System.Drawing.Image)
        Me.SyringePicture.Location = New System.Drawing.Point(276, 40)
        Me.SyringePicture.Name = "SyringePicture"
        Me.SyringePicture.Size = New System.Drawing.Size(40, 16)
        Me.SyringePicture.TabIndex = 40
        Me.SyringePicture.TabStop = False
        Me.SyringePicture.Visible = False
        '
        'NeedlePicture
        '
        Me.NeedlePicture.Image = CType(resources.GetObject("NeedlePicture.Image"), System.Drawing.Image)
        Me.NeedlePicture.Location = New System.Drawing.Point(92, 40)
        Me.NeedlePicture.Name = "NeedlePicture"
        Me.NeedlePicture.Size = New System.Drawing.Size(40, 16)
        Me.NeedlePicture.TabIndex = 39
        Me.NeedlePicture.TabStop = False
        Me.NeedlePicture.Visible = False
        '
        'Station3Label
        '
        Me.Station3Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station3Label.Location = New System.Drawing.Point(320, 156)
        Me.Station3Label.Name = "Station3Label"
        Me.Station3Label.Size = New System.Drawing.Size(112, 16)
        Me.Station3Label.TabIndex = 26
        Me.Station3Label.Text = "Post Heater"
        Me.Station3Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Station3Label.Visible = False
        '
        'Station2Label
        '
        Me.Station2Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station2Label.Location = New System.Drawing.Point(184, 156)
        Me.Station2Label.Name = "Station2Label"
        Me.Station2Label.Size = New System.Drawing.Size(112, 16)
        Me.Station2Label.TabIndex = 25
        Me.Station2Label.Text = "Disp. Heater"
        Me.Station2Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Station2Label.Visible = False
        '
        'Needle
        '
        Me.Needle.ForeColor = System.Drawing.Color.Red
        Me.Needle.Location = New System.Drawing.Point(148, 32)
        Me.Needle.Name = "Needle"
        Me.Needle.Size = New System.Drawing.Size(56, 24)
        Me.Needle.TabIndex = 143
        Me.Needle.Text = "50 oC"
        Me.Needle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Needle.Visible = False
        '
        'Station1Picture
        '
        Me.Station1Picture.Image = CType(resources.GetObject("Station1Picture.Image"), System.Drawing.Image)
        Me.Station1Picture.Location = New System.Drawing.Point(48, 92)
        Me.Station1Picture.Name = "Station1Picture"
        Me.Station1Picture.Size = New System.Drawing.Size(112, 64)
        Me.Station1Picture.TabIndex = 136
        Me.Station1Picture.TabStop = False
        Me.Station1Picture.Visible = False
        '
        'Station1Label
        '
        Me.Station1Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station1Label.Location = New System.Drawing.Point(48, 156)
        Me.Station1Label.Name = "Station1Label"
        Me.Station1Label.Size = New System.Drawing.Size(112, 16)
        Me.Station1Label.TabIndex = 24
        Me.Station1Label.Text = "Pre Heater"
        Me.Station1Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Station1Label.Visible = False
        '
        'Station2Picture
        '
        Me.Station2Picture.Image = CType(resources.GetObject("Station2Picture.Image"), System.Drawing.Image)
        Me.Station2Picture.Location = New System.Drawing.Point(184, 92)
        Me.Station2Picture.Name = "Station2Picture"
        Me.Station2Picture.Size = New System.Drawing.Size(112, 64)
        Me.Station2Picture.TabIndex = 137
        Me.Station2Picture.TabStop = False
        Me.Station2Picture.Visible = False
        '
        'NeedleLabel
        '
        Me.NeedleLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.NeedleLabel.Location = New System.Drawing.Point(84, 64)
        Me.NeedleLabel.Name = "NeedleLabel"
        Me.NeedleLabel.Size = New System.Drawing.Size(128, 16)
        Me.NeedleLabel.TabIndex = 23
        Me.NeedleLabel.Text = "Needle Heater"
        Me.NeedleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.NeedleLabel.Visible = False
        '
        'SyringeLabel
        '
        Me.SyringeLabel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SyringeLabel.Location = New System.Drawing.Point(268, 64)
        Me.SyringeLabel.Name = "SyringeLabel"
        Me.SyringeLabel.Size = New System.Drawing.Size(128, 16)
        Me.SyringeLabel.TabIndex = 22
        Me.SyringeLabel.Text = "Syringe Heater"
        Me.SyringeLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.SyringeLabel.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(0, 8)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(104, 26)
        Me.TextBox1.TabIndex = 132
        Me.TextBox1.Text = "TextBox1"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(0, 32)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(104, 26)
        Me.TextBox2.TabIndex = 132
        Me.TextBox2.Text = "TextBox1"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.ComboBox3)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBox1.Location = New System.Drawing.Point(16, 704)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 256)
        Me.GroupBox1.TabIndex = 130
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Thermal Readings"
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"Production Start", "Board Comes", "Board Goes", "Pause", "Resume", "Board Partial Failure", "Board Total Failure", "Production End", "Purge And Clean", "Dot Size Check Passed", "Volume Calibration Passed", "Volume Calibration Failed", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox1.TabIndex = 140
        Me.ComboBox1.Text = "ComboBox1"
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.Location = New System.Drawing.Point(16, 80)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(176, 40)
        Me.Button1.TabIndex = 139
        Me.Button1.Text = "enter 1 SPC event"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button2.Location = New System.Drawing.Point(16, 192)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(208, 56)
        Me.Button2.TabIndex = 139
        Me.Button2.Text = "enter all SPC events"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button3.Location = New System.Drawing.Point(16, 136)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(192, 40)
        Me.Button3.TabIndex = 139
        Me.Button3.Text = "generate SPC report"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ComboBox2
        '
        Me.ComboBox2.Items.AddRange(New Object() {"Production Start", "Board Comes", "Board Goes", "Pause", "Resume", "Board Partial Failure", "Board Total Failure", "Production End", "Purge And Clean", "Dot Size Check Passed", "Volume Calibration Passed", "Volume Calibration Failed", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed"})
        Me.ComboBox2.Location = New System.Drawing.Point(16, 32)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox2.TabIndex = 140
        Me.ComboBox2.Text = "ComboBox1"
        '
        'ComboBox3
        '
        Me.ComboBox3.Items.AddRange(New Object() {"Width Adjustment Failed", "Lifter 1 Blocked", "Lifter 2 Blocked", "Lifter 3 Blocked", "Retrieve Timeout", "Stage 1 Travel Timeout", "Release Travel Timeout", "Stage 3 Travel Timeout", "Board Not Aligned", "PLC Communication Error", "PLC No Signal", "Low Incoming Air Pressure", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed", "Volume Calibration Failed"})
        Me.ComboBox3.Location = New System.Drawing.Point(288, 32)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox3.TabIndex = 140
        Me.ComboBox3.Text = "ComboBox1"
        '
        'Button4
        '
        Me.Button4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.Location = New System.Drawing.Point(288, 80)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(176, 40)
        Me.Button4.TabIndex = 139
        Me.Button4.Text = "generate popup"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PanelProDownTimeInfor
        '
        Me.PanelProDownTimeInfor.BackColor = System.Drawing.SystemColors.Info
        Me.PanelProDownTimeInfor.Controls.Add(Me.RichTextBoxNote)
        Me.PanelProDownTimeInfor.Controls.Add(Me.CheckBoxPotOn)
        Me.PanelProDownTimeInfor.Controls.Add(Me.DoorLock)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPotReset)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label6)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonOpenFile)
        Me.PanelProDownTimeInfor.Controls.Add(Me.TextBoxFilename)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label5)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonNdlCalib)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonClean)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPurge)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonChgSyringe)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonHome)
        Me.PanelProDownTimeInfor.Controls.Add(Me.LabelMessege)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonVolCalib)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label7)
        Me.PanelProDownTimeInfor.Location = New System.Drawing.Point(0, 0)
        Me.PanelProDownTimeInfor.Name = "PanelProDownTimeInfor"
        Me.PanelProDownTimeInfor.Size = New System.Drawing.Size(768, 384)
        Me.PanelProDownTimeInfor.TabIndex = 6
        '
        'RichTextBoxNote
        '
        Me.RichTextBoxNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBoxNote.Location = New System.Drawing.Point(56, 128)
        Me.RichTextBoxNote.Name = "RichTextBoxNote"
        Me.RichTextBoxNote.ReadOnly = True
        Me.RichTextBoxNote.Size = New System.Drawing.Size(688, 208)
        Me.RichTextBoxNote.TabIndex = 98
        Me.RichTextBoxNote.Text = ""
        '
        'CheckBoxPotOn
        '
        Me.CheckBoxPotOn.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxPotOn.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.CheckBoxPotOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBoxPotOn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CheckBoxPotOn.ImageIndex = 6
        Me.CheckBoxPotOn.ImageList = Me.ImageListPotEtc
        Me.CheckBoxPotOn.Location = New System.Drawing.Point(475, 0)
        Me.CheckBoxPotOn.Name = "CheckBoxPotOn"
        Me.CheckBoxPotOn.Size = New System.Drawing.Size(75, 56)
        Me.CheckBoxPotOn.TabIndex = 119
        Me.CheckBoxPotOn.Text = "Pot Life On"
        Me.CheckBoxPotOn.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'DoorLock
        '
        Me.DoorLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.DoorLock.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.DoorLock.Checked = True
        Me.DoorLock.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.DoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.DoorLock.ImageIndex = 3
        Me.DoorLock.ImageList = Me.ImageListPotEtc
        Me.DoorLock.Location = New System.Drawing.Point(650, 0)
        Me.DoorLock.Name = "DoorLock"
        Me.DoorLock.Size = New System.Drawing.Size(75, 56)
        Me.DoorLock.TabIndex = 116
        Me.DoorLock.Text = "Lock Door"
        Me.DoorLock.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonPotReset
        '
        Me.ButtonPotReset.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonPotReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPotReset.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPotReset.ImageIndex = 1
        Me.ButtonPotReset.ImageList = Me.ImageListPotEtc
        Me.ButtonPotReset.Location = New System.Drawing.Point(550, 0)
        Me.ButtonPotReset.Name = "ButtonPotReset"
        Me.ButtonPotReset.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPotReset.TabIndex = 105
        Me.ButtonPotReset.Text = "Reset Pot"
        Me.ButtonPotReset.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 128)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 23)
        Me.Label6.TabIndex = 97
        Me.Label6.Text = "Note"
        '
        'ButtonOpenFile
        '
        Me.ButtonOpenFile.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ButtonOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonOpenFile.Location = New System.Drawing.Point(672, 80)
        Me.ButtonOpenFile.Name = "ButtonOpenFile"
        Me.ButtonOpenFile.Size = New System.Drawing.Size(72, 26)
        Me.ButtonOpenFile.TabIndex = 96
        Me.ButtonOpenFile.Text = "Open"
        '
        'TextBoxFilename
        '
        Me.TextBoxFilename.BackColor = System.Drawing.Color.White
        Me.TextBoxFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBoxFilename.Location = New System.Drawing.Point(56, 80)
        Me.TextBoxFilename.Name = "TextBoxFilename"
        Me.TextBoxFilename.ReadOnly = True
        Me.TextBoxFilename.Size = New System.Drawing.Size(600, 27)
        Me.TextBoxFilename.TabIndex = 95
        Me.TextBoxFilename.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 94
        Me.Label5.Text = "File"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonNdlCalib
        '
        Me.ButtonNdlCalib.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonNdlCalib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonNdlCalib.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonNdlCalib.ImageIndex = 3
        Me.ButtonNdlCalib.ImageList = Me.ImageListGeneralTools
        Me.ButtonNdlCalib.Location = New System.Drawing.Point(300, 0)
        Me.ButtonNdlCalib.Name = "ButtonNdlCalib"
        Me.ButtonNdlCalib.Size = New System.Drawing.Size(75, 56)
        Me.ButtonNdlCalib.TabIndex = 91
        Me.ButtonNdlCalib.Text = "Ndl. Cal."
        Me.ButtonNdlCalib.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonClean
        '
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClean.ImageIndex = 5
        Me.ButtonClean.ImageList = Me.ImageListGeneralTools
        Me.ButtonClean.Location = New System.Drawing.Point(150, 0)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.Size = New System.Drawing.Size(75, 56)
        Me.ButtonClean.TabIndex = 89
        Me.ButtonClean.Text = "Clean On"
        Me.ButtonClean.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonPurge
        '
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPurge.ImageIndex = 4
        Me.ButtonPurge.ImageList = Me.ImageListGeneralTools
        Me.ButtonPurge.Location = New System.Drawing.Point(75, 0)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPurge.TabIndex = 88
        Me.ButtonPurge.Text = "Purge On"
        Me.ButtonPurge.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonChgSyringe
        '
        Me.ButtonChgSyringe.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonChgSyringe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonChgSyringe.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonChgSyringe.ImageIndex = 1
        Me.ButtonChgSyringe.ImageList = Me.ImageListGeneralTools
        Me.ButtonChgSyringe.Location = New System.Drawing.Point(376, 0)
        Me.ButtonChgSyringe.Name = "ButtonChgSyringe"
        Me.ButtonChgSyringe.Size = New System.Drawing.Size(75, 56)
        Me.ButtonChgSyringe.TabIndex = 87
        Me.ButtonChgSyringe.Text = "Chg. Syr."
        Me.ButtonChgSyringe.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonHome
        '
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonHome.ImageIndex = 0
        Me.ButtonHome.ImageList = Me.ImageListGeneralTools
        Me.ButtonHome.Location = New System.Drawing.Point(0, 0)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(75, 56)
        Me.ButtonHome.TabIndex = 86
        Me.ButtonHome.Text = "Home"
        Me.ButtonHome.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Menu
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Blue
        Me.LabelMessege.Location = New System.Drawing.Point(0, 352)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(760, 32)
        Me.LabelMessege.TabIndex = 85
        Me.LabelMessege.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonVolCalib
        '
        Me.ButtonVolCalib.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonVolCalib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonVolCalib.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonVolCalib.ImageIndex = 2
        Me.ButtonVolCalib.ImageList = Me.ImageListGeneralTools
        Me.ButtonVolCalib.Location = New System.Drawing.Point(225, 0)
        Me.ButtonVolCalib.Name = "ButtonVolCalib"
        Me.ButtonVolCalib.Size = New System.Drawing.Size(75, 56)
        Me.ButtonVolCalib.TabIndex = 90
        Me.ButtonVolCalib.Text = "Vol. Cal."
        Me.ButtonVolCalib.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ImageListOperation
        '
        Me.ImageListOperation.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageListOperation.ImageStream = CType(resources.GetObject("ImageListOperation.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOperation.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimerMonitor
        '
        Me.TimerMonitor.Interval = 50
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.Location = New System.Drawing.Point(0, 384)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(768, 576)
        Me.PanelVision.TabIndex = 7
        '
        'btExit
        '
        Me.btExit.ForeColor = System.Drawing.Color.Black
        Me.btExit.Location = New System.Drawing.Point(416, 8)
        Me.btExit.Name = "btExit"
        Me.btExit.Size = New System.Drawing.Size(80, 48)
        Me.btExit.TabIndex = 133
        Me.btExit.Text = "Exit"
        '
        'FormProduction
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1276, 990)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.PanelProDownTimeInfor)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Menu = Me.MainMenuProduction
        Me.Name = "FormProduction"
        Me.Text = "Production"
        Me.Panel2.ResumeLayout(False)
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel5.ResumeLayout(False)
        Me.ConveyorBox.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.HeaterBox.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.PanelProDownTimeInfor.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'reset!
        ResetToIdle()

        'gui visibility
        HeaterBox.Visible = IDS.Data.Hardware.Thermal.HeaterFeatureEnabled

        'initialize private flags
        HasBeenRunning = False
        m_PotLifeOn = False
        TextBoxFilename.Text = " "
        ProductionInfoDispClear()

        m_Execution.m_Pattern.SubCallSheetInitialization(200)     '200 subsheets maximum without duplicated name
        Programming.ErrorSubSheetStructIni(200, 500)    '500 subsheets maximum with duplicated name

        ' get default value from the default pat file
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        SystemSetupDataRetrieve()
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()

        'if jetting valve or auger valve turn on material air
        With IDS.Data.Hardware.Dispenser.Left
            If .HeadType = "Jetting Valve" Or .HeadType = "Auger Valve" Then
                m_Tri.TurnOn("Material Air")
            End If
        End With

        'error handling message
        Form_Service.ResetEventCode()

        'hardware
        Laser.OpenPort()
        Weighting_Scale.OpenPort()
        'conveyor
        Conveyor.OpenPort()

        'vision
        PanelVision.Controls.Add(Vision.FrmVision.PanelVision) 'lsgoh
        SwitchToTeachCamera()
        ValueBrightness.Value = IDS.Data.Hardware.Camera.Brightness

        'motion controller
        m_Tri.Connect_Controller()
        SetState("Homing")

        'timers start
        IDS.StartErrorCheck()
        TimerMonitor.Start()
        Programming.IOCheck.Start()
        Conveyor.PositionTimer.Start()

        ' background threads help to update UI without slowing them down.
        ThreadMonitor = New Threading.Thread(AddressOf StateMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.Normal
        'ThreadMonitor.IsBackground = True
        ThreadMonitor.Start()

        ThreadExecutor = New Threading.Thread(AddressOf StateChanger)
        ThreadExecutor.Priority = Threading.ThreadPriority.Normal
        ThreadExecutor.Start()

    End Sub

    Private Sub FormProduction_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

    End Sub

    Public Sub ProductionInfoDispClear()
        Dim i As Integer
        For i = 0 To 49
            Production_info(i) = ""
        Next

        RichTextBoxNote.Lines = Production_info
        RichTextBoxNote.Show()
    End Sub


    Public Sub ProductionInfoDisp()
        '  IDS.Data.OpenData()
        Production_info(0) = "Author: " + IDS.Data.Hardware.SPC.ProgAuthorName
        Production_info(1) = "Contact: " + IDS.Data.Hardware.SPC.ProgAuthorContact
        Production_info(2) = "Notes: " + IDS.Data.Hardware.SPC.ProductionNote
        Production_info(3) = "Needle: Left"
        Production_info(4) = "        Material: " & IDS.Data.Hardware.Dispenser.Left.MaterialInfo

        Production_info(7) = "        Left Needle Color "

        RichTextBoxNote.Lines = Production_info

        RichTextBoxNote.Show()
        If Not GetInputState = 0 Then Application.DoEvents()
    End Sub

    ' open file sub
    Private Sub OpenFile()

        ButtonOpenFile.Enabled = False

        If Programming.ProductionFileOpen() < 0 Then
            ButtonOpenFile.Enabled = True
            Exit Sub
        End If

        IDS.Data.OpenData()
        If Nothing = TextBoxFilename.Text Then
            ButtonOpenFile.Enabled = True
            Return
        End If

        Dim filename As String = m_Execution.m_File.FolderWithNameFromFileName(TextBoxFilename.Text)
        Programming.Disp_Dispenser_Unit_info()

        LabelMessage("Please wait, system is uploading..")
        TextBoxFilename.Refresh()
        RichTextBoxNote.Refresh()

        ProductionInfoDisp()

        HeaterSettings()
        OnHeaters()

        PressureSettings()
        OnPressure()

        LabelMessage("Finish..")
        ButtonOpenFile.Enabled = True
    End Sub

    Private Sub ButtonOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenFile.Click
        OpenFile()
    End Sub

    Public Sub ProdPurgeClean() 'for now, do not discriminate between left and right

        Dim cmdStr As String
        Dim needleIO, mode, purgeduration, cleanduration As Integer

        LockMovementButtons()
        LabelMessage("Purge and clean.")
        If ShouldLog() Then Form_Service.LogEventInSPCReport("Purge And Clean")
        needleIO = gIOLeftNeedle
        purgeduration = IDS.Data.Hardware.Dispenser.Left.AutoPurgingDuration
        cleanduration = IDS.Data.Hardware.Dispenser.Left.AutoCleaningDuration

        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.PurgePosition.Z
        position(0) = IDS.Data.Hardware.Gantry.PurgePosition.X - offset(0)
        position(1) = IDS.Data.Hardware.Gantry.PurgePosition.Y - offset(1)

        m_Tri.SetMachineRun()
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ElementZSpeed)
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration 'z axis move to purge pos
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration 'x and y axis move to purge pos
        If Not m_Tri.Move_Z(offset(2)) Then GoTo StopCalibration 'z axis move to purge pos
        cmdStr = "WA(" + purgeduration.ToString + ")"
        m_Tri.TurnOn("Left Needle IO")
        m_Tri.m_TriCtrl.Execute(cmdStr) 'wait after turning on valve
        m_Tri.TurnOff("Left Needle IO")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo StopCalibration 'z axis move to 0 before move to clean position

        offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.CleanPosition.Z
        position(0) = IDS.Data.Hardware.Gantry.CleanPosition.X - offset(0)
        position(1) = IDS.Data.Hardware.Gantry.CleanPosition.Y - offset(1)

        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration 'x and y axis move to clean pos
        If Not m_Tri.Move_Z(offset(2)) Then GoTo StopCalibration 'z axis move to clean pos
        cmdStr = "WA(" + cleanduration.ToString + ")"
        m_Tri.TurnOn("Clean")
        m_Tri.m_TriCtrl.Execute(cmdStr) 'wait after turning on valve
        m_Tri.TurnOff("Clean")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo StopCalibration 'z axis move to 0

        'commented out since kr changed to FSM to handle state
        'If IsRunning() Then 'if in the middle of a continuous run
        '    m_Tri.SetMachineStop()
        '    Exit Sub
        'End If

StopCalibration:
        ResetToIdle()
    End Sub

    Public Function ProdChangeSyringeMaterial(ByVal needle As String) As Integer

    End Function

    ' Start is a continuation of tmrrunpro_tick. it is called when continuous    '
    ' run is not checked (conveyor box), and skips the PLC signal check.             '
    ' it will check the current mode, enable/disable the appropriate timers,         '
    ' check the door/key status, set the status message.                             '
    ' it will then continue the dispensing process through runproduction()           '

    Public Sub Start() 'kr dec24 made some changes - better check how they affect runtime.

        Programming.SetRunMode(4) 'Wet mode

        If DoorCheck() = False Then
            MyMsgBox("Door is open")
            ResetToIdle()
            Exit Sub
        End If

        If (ContinuousMode.Checked = True) Then
            ResetTimer("Reset Volume Calibration Timers")
            ResetTimer("Reset Autopurging Timers")
        End If

        GeneratePatternAndRunProgram()

    End Sub

    Public Sub TBOperation_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick


        If e.Button Is TBOperation.Buttons(0) Then 'run

            SetState("Start")

        ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause

            PauseDispensing()

        ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop

            StopDispensing()
        End If

    End Sub

    Private Sub ButtonHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHome.Click
        SetState("Homing")
    End Sub

    Private Sub ButtonChgSyringe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonChgSyringe.Click
        SetState("Change Syringe")
    End Sub

    Private Sub ButtonPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPurge.Click
        If SetState("Purge") Then DoPurge()
    End Sub

    Private Sub ButtonClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClean.Click
        If SetState("Clean") Then DoClean()
    End Sub

    Private Sub ButtonVolCalib_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolCalib.Click
        SetState("Volume Calibration")
    End Sub

    Private Sub ButtonNdlCalib_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNdlCalib.Click
        SetState("Needle Calibration")
    End Sub

    Private Sub DoorLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoorLock.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't unlock the door when machine running!")
            DoorLock.Text = "Unlock Door"
            DoorLock.ImageIndex = 2
            Exit Sub
        End If

        If DoorLock.Checked = False Then
            DoorLock.Text = "Unlock Door"
            DoorLock.ImageIndex = 2
            LockDoor()
        Else
            DoorLock.Text = "Lock Door"
            DoorLock.ImageIndex = 3
            UnlockDoor()
        End If
    End Sub

    Private Sub CheckBoxPotOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPotOn.CheckedChanged

        If CheckBoxPotOn.Checked = True Then
            CheckBoxPotOn.Text = "Turn Pot Life Off"
            CheckBoxPotOn.ImageIndex = 0
            m_PotLifeOn = True
            ResetTimer("Reset Potlife Start Timer")
        Else
            CheckBoxPotOn.Text = "Turn Pot Life On"
            CheckBoxPotOn.ImageIndex = 6
            m_PotLifeOn = False
        End If

    End Sub

    Private Sub ButtonPotReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPotReset.Click

        ResetTimer("Reset Potlife Start Timer")

    End Sub

    '''''''''''''''
    '                                       kr                                       '
    ' timermonitor is to monitor for potlife and machine state changes               '
    ' if the run state changes to idle, change back to teaching camera vision        '
    ' changes in run state also updates the tower lights                             '
    ' timermonitor updates continuously, does not disable itself                     '
    ' calls prodpurgeclean, volumcalib and prodchangesyringe                         '

    Public Delegate Sub DispensingFinish()
    Public DispensingFinishCallback As DispensingFinish = New DispensingFinish(AddressOf DispensingCallback)

    Public Sub DispensingSwitchCallback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Public Sub DispensingCallback()
        If ContinuousMode.Checked = True Then
            Finish_Dispensing()
        Else
            ResetToIdle()
        End If
    End Sub

    Private Sub ButtonCV_Prod_Retrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCV_Prod_Retrieve.Click
        Conveyor.Command("Retrieve")
    End Sub

    Private Sub ButtonCV_Prod_Release_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCV_Prod_Release.Click
        Conveyor.Command("Release")
    End Sub

    Private Sub ButtonStartFirstStage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStartFirstStage.Click
        DIO_Service.TriggerUpstream()
    End Sub

    Public Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContinuousMode.Click

        '                                       kr                                       '
        ' this option allows user to choose between continuous operation or to operate   '
        ' on a single board using the buttons start, retrieve and release to control the '
        ' conveyor manually. otherwise the conveyor will be controlled by signals        '

        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag
        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        If door_interlock And door_close = False Then
            LabelMessage("Door is open. Please close to turn on continuous mode.")
            Exit Sub
        ElseIf ContinuousMode.Checked Then
            Conveyor.Command("Manual Mode")
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned off.")
            End If
        ElseIf ContinuousMode.Checked = False Then
            Conveyor.Command("Auto Mode")
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned on.")
            End If
        End If

        If ContinuousMode.Checked Then
            ButtonStartFirstStage.Enabled = True
            ButtonCV_Prod_Retrieve.Enabled = True
            ButtonCV_Prod_Release.Enabled = True
        Else
            ButtonStartFirstStage.Enabled = False
            ButtonCV_Prod_Retrieve.Enabled = False
            ButtonCV_Prod_Release.Enabled = False
        End If

        ContinuousMode.Checked = Not ContinuousMode.Checked

    End Sub

    Private Sub ValueBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness.ValueChanged
        If 0 <= ValueBrightness.Value < 255 Then
            Vision.IDSV_SetBrightness(ValueBrightness.Value)
        Else
            ValueBrightness.Value = 128
            Vision.IDSV_SetBrightness(ValueBrightness.Value)
        End If
    End Sub

    Public Sub WriteSPCReport()

        'Dim checkID As Integer
        'checkID = CInt(IDS.__ID) Mod 100
        'If (checkID <> 0) Then
        '    '   Cannot write the event while device error is occured.   '
        '    Exit Sub
        'End If


        If ShouldLog() And HasBeenRunning Then Form_Service.LogEventInSPCReport("Production End")

        'production end
        If IDS.Data.Hardware.SPC.EnableSPCReport And HasBeenRunning Then
            Form_Service.GenerateSPCReport()
            HasBeenRunning = False
        End If


    End Sub


    '                                       kr                                       '
    ' tmrRunPro_tick ensures continuous operation in production (if conveyor ticked) '
    ' enabled by tmrfinish, prodpurgeclean, prodvolumcalib and prodchangesyringe     '
    ' also enabled when starting production run                                      '
    ' this function will check to see if starting is valid (m_ButtonState, mode)      '
    ' or else it will exit. if all checks are passed, the last check is for a signal '
    ' from the PLC to tell when the board is ready for dispensing.                   '
    ' then it will call Start                                                    '

    Public Sub Conveyor_Check()

        Dim onoff As Integer = 1
        Dim conveyorErrorID, ErrorID As Integer

        ChangeButtonState("Only Stop Displayed")

        While onoff = 1 And Not m_Tri.MotionError

            m_Tri.m_TriCtrl.In(gIOloadReady, gIOloadReady, onoff) 'when ready set onoff = 0
            LabelMessage("Waiting for board..")
            TraceDoEvents()

            If IsIdle() Or m_Tri.MotionError Then
                Form_Service.LogEventInSPCReport("Motion controller error")
                ResetToIdle()
                Exit Sub
            End If

            '3202,3203,3204 are errors from lifter
            '3205,3206,3208 are conveyor retrieve or stage 1/3 jamming errors.
            conveyorErrorID = CInt(IDS.__ID) Mod 10000
            If conveyorErrorID = 3202 Or conveyorErrorID = 3203 Or conveyorErrorID = 3204 Or conveyorErrorID = 3205 Or conveyorErrorID = 3206 Or conveyorErrorID = 3208 Then
                LabelMessage("No board is coming or conveyor jammed.")
                ResetToIdle()
                Exit Sub
            End If

        End While

        'ErrorID = CInt(IDS.__ID) Mod 100
        'While ErrorID <> 0
        '    ErrorID = CInt(IDS.__ID) Mod 100
        '    LabelMessage("Clear pending error messages first.")
        '    TraceDoEvents()
        'End While

        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Comes")
        Start()

    End Sub

    Public Sub ResetTimer(ByVal command As String)

        Select Case command
            Case "Reset Autopurging Timers"
                StartTimeOfMachineIdle = Now
            Case "Reset Potlife Start Timer"
                RemainingTimeOfPotLifeStart = Now
        End Select

    End Sub

    '''''''''''''''
    '                                       kr                                       '
    ' show and set tower light according to machine state                            '
    '''''''''''''''

    Public Sub Finish_Dispensing()

        ' for both single and continuous run
        m_Tri.m_TriCtrl.Op(gIODispensingDone, 1) 'release board
        TravelToParkPosition()
        m_Tri.m_TriCtrl.Op(gIODispensingDone, 0) 'release board

        Form_Service.LogEventInSPCReport("Board Goes")

        If IsRunning() Then
            LabelMessage("Dispensing cycle re-starting.")
            HasBeenRunning = True
            Conveyor_Check() 'after timer finish automatically call back Start()
            Exit Sub
        End If

    End Sub


    Public Function DoorCheck()

        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag

        If door_interlock = False Then
            Return True
        ElseIf door_interlock = True Then
            If door_close Then
                LockDoor()
            End If
            Return door_close
        End If

    End Function

    Private Sub OnHeaters()
        'turn on heaters
        'With IDS.Devices.Thermal.FrmHeater
        '    If ContinuousMode.Checked = True Then
        '        If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '            .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.OperationTemp)
        '            .On_Off_thermal(1, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '            .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.OperationTemp)
        '            .On_Off_thermal(2, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '            .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.OperationTemp)
        '            .On_Off_thermal(3, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '            .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.OperationTemp)
        '            .On_Off_thermal(4, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '            .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.OperationTemp)
        '            .On_Off_thermal(5, True)
        '        End If
        '    ElseIf ContinuousMode.Checked = False Then
        '        If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '            .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.StandbyTemp)
        '            .On_Off_thermal(1, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '            .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.StandbyTemp)
        '            .On_Off_thermal(2, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '            .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.StandbyTemp)
        '            .On_Off_thermal(3, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '            .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.StandbyTemp)
        '            .On_Off_thermal(4, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '            .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.StandbyTemp)
        '            .On_Off_thermal(5, False)
        '        End If
        '    End If
        'End With
    End Sub

    Private Sub HeaterSettings()
        'With IDS.Devices.Thermal.FrmHeater
        '    If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '        Needle.Visible = True
        '        NeedlePicture.Visible = True
        '        NeedleLabel.Visible = True
        '        MyHeaterSettings.Heater_Enabled(0) = True
        '        .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.StandbyTemp)
        '        .Set_Alarm(1, IDS.Data.Hardware.Thermal.Needle.AlarmThreshold)
        '    Else
        '        Needle.Visible = False
        '        NeedlePicture.Visible = False
        '        NeedleLabel.Visible = False
        '        MyHeaterSettings.Heater_Enabled(0) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '        Syringe.Visible = True
        '        SyringePicture.Visible = True
        '        SyringeLabel.Visible = True
        '        MyHeaterSettings.Heater_Enabled(1) = True
        '        .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.StandbyTemp)
        '        .Set_Alarm(2, IDS.Data.Hardware.Thermal.Syringe.AlarmThreshold)
        '    Else
        '        Syringe.Visible = False
        '        SyringePicture.Visible = False
        '        SyringeLabel.Visible = False
        '        MyHeaterSettings.Heater_Enabled(1) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '        Station1.Visible = True
        '        Station1Picture.Visible = True
        '        Station1Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(2) = True
        '        .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.StandbyTemp)
        '        .Set_Alarm(3, IDS.Data.Hardware.Thermal.Station1.AlarmThreshold)
        '    Else
        '        Station1.Visible = False
        '        Station1Picture.Visible = False
        '        Station1Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(2) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '        Station2.Visible = True
        '        Station2Picture.Visible = True
        '        Station2Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(3) = True
        '        .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.StandbyTemp)
        '        .Set_Alarm(4, IDS.Data.Hardware.Thermal.Station2.AlarmThreshold)
        '    Else
        '        Station2.Visible = False
        '        Station2Picture.Visible = False
        '        Station2Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(3) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '        Station3.Visible = True
        '        Station3Picture.Visible = True
        '        Station3Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(4) = True
        '        .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.StandbyTemp)
        '        .Set_Alarm(5, IDS.Data.Hardware.Thermal.Station3.AlarmThreshold)
        '    Else
        '        Station3.Visible = False
        '        Station3Picture.Visible = False
        '        Station3Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(4) = False
        '    End If
        'End With
    End Sub

    Public Sub PressureSettings()

    End Sub

    Public Sub OnPressure()
        IDS.Data.OpenData()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form_Service.LogEventInSPCReport(ComboBox1.SelectedItem.ToString)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        For i As Integer = 0 To ComboBox1.Items.Count - 1
            Form_Service.LogEventInSPCReport(ComboBox1.Items(i))
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form_Service.GenerateSPCReport()
    End Sub

    Private Sub TimerMonitor_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMonitor.Tick

        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag
        Dim LowMaterialSensorLeft As Boolean = IDS.Devices.DIO.DIO.LevelSensor(True)

        'reset the timers if 
        '1) door is not closed when door interlocked is true
        '2) program is currently paused or running
        '3) conveyor mode is not ticked
        If (door_close = False And door_interlock = True) Or ContinuousMode.Checked = False And IsIdle() Then
            ResetTimer("Reset Autopurging Timers")
            ResetTimer("Reset Potlife Start Timer")
            Exit Sub
        End If

        Dim auto_purging_enabled As Boolean = IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption And ContinuousMode.Checked And IsIdle() And Form_Service.Visible = False
        CurrentTime = Now

        If m_PotLifeOn Then
            ElapsedTimeSincePotLifeStart = (CurrentTime.Ticks - RemainingTimeOfPotLifeStart.Ticks) * tickToMinute
            If ElapsedTimeSincePotLifeStart >= IDS.Data.Hardware.Dispenser.Left.PotLifeDuration Then
                'what do we want to do when potlife is up?
            End If
        End If

        'auto purging and auto cleaning is only carried out if conveyor mode is ticked and auto purging option is enabled
        If auto_purging_enabled Then
            ElapsedTimeSinceMachineIdle = (CurrentTime.Ticks - StartTimeOfMachineIdle.Ticks) * tickToMinute    'elapsed time
            If ElapsedTimeSinceMachineIdle >= IDS.Data.Hardware.Dispenser.Left.AutoPurgingInterval Then  'when elapsed time > time
                ResetTimer("Reset Autopurging Timers")
                SetState("AutoPurge")
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Form_Service.DisplayErrorMessage(ComboBox3.SelectedItem.ToString)
    End Sub

    Private Sub ResetPLCLogic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPLCLogic.Click
        Conveyor.Command("Reset PLC Logic")
    End Sub

    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExit.Click

        If ContinuousMode.Checked = True Then
            ContinuousMode.Checked = False
        End If

        'error handling
        Form_Service.ResetEventCode()

        'vision
        SwitchToTeachCamera()
        PanelVision.Controls.Remove(Vision.FrmVision.PanelVision) 'lim

        'timers stop
        IDS.StopErrorCheck()
        TimerMonitor.Stop()
        Programming.IOCheck.Stop()
        ThreadMonitor.Abort()
        ThreadExecutor.Abort()
        Conveyor.PositionTimer.Stop()

        'motion controller
        m_Tri.TurnOff("Material Air")
        m_Tri.Disconnect_Controller()

        'hardware
        MyConveyorSettings.CloseConveyorSetup()
        Conveyor.ClosePort()
        Weighting_Scale.ClosePort()
        Laser.ClosePort()
        OffLaser()
        OffTowerLamp()
        UnlockDoor()

        Close()
        IDS.FrmExecution.Hide()

    End Sub
End Class
