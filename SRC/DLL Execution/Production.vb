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
    Private productCnt As Integer = 0
    'constants
    Public Const tickToMinute As Double = 0.0000000016666666  '0.0000001/60.0
    Public m_PotLifeExpire As Boolean = False

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
    Friend WithEvents Station3Label As System.Windows.Forms.Label
    Friend WithEvents Station2Label As System.Windows.Forms.Label
    Friend WithEvents Station1Label As System.Windows.Forms.Label
    Friend WithEvents NeedleLabel As System.Windows.Forms.Label
    Friend WithEvents SyringeLabel As System.Windows.Forms.Label
    Friend WithEvents ConveyorBox As System.Windows.Forms.GroupBox
    Friend WithEvents HeaterBox As System.Windows.Forms.GroupBox
    Friend WithEvents ResetPLCLogic As System.Windows.Forms.Button
    Friend WithEvents btExit As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RedLight As System.Windows.Forms.PictureBox
    Friend WithEvents GreenLight As System.Windows.Forms.PictureBox
    Friend WithEvents TowerLightImageList As System.Windows.Forms.ImageList
    Friend WithEvents panelVision As System.Windows.Forms.Panel
    Friend WithEvents AmberLight As System.Windows.Forms.PictureBox
    Friend WithEvents btPlay As System.Windows.Forms.Button
    Friend WithEvents btPause As System.Windows.Forms.Button
    Friend WithEvents btStop As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents tbEquipmentID As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents tbCurrentDispense As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents tbDispensedUnit As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents tbTime As System.Windows.Forms.TextBox
    Friend WithEvents TimeDisplayTimer As System.Windows.Forms.Timer
    Friend WithEvents tbDispensePressure As System.Windows.Forms.TextBox
    Friend WithEvents tbPortLifeRemain As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormProduction))
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.tbTime = New System.Windows.Forms.TextBox
        Me.tbDispensedUnit = New System.Windows.Forms.TextBox
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.tbCurrentDispense = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.tbPortLifeRemain = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.tbDispensePressure = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.tbEquipmentID = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GreenLight = New System.Windows.Forms.PictureBox
        Me.AmberLight = New System.Windows.Forms.PictureBox
        Me.RedLight = New System.Windows.Forms.PictureBox
        Me.btStop = New System.Windows.Forms.Button
        Me.btPause = New System.Windows.Forms.Button
        Me.btPlay = New System.Windows.Forms.Button
        Me.ContinuousMode = New System.Windows.Forms.CheckBox
        Me.btExit = New System.Windows.Forms.Button
        Me.ConveyorBox = New System.Windows.Forms.GroupBox
        Me.ResetPLCLogic = New System.Windows.Forms.Button
        Me.ButtonStartFirstStage = New System.Windows.Forms.Button
        Me.ButtonCV_Prod_Release = New System.Windows.Forms.Button
        Me.ButtonCV_Prod_Retrieve = New System.Windows.Forms.Button
        Me.HeaterBox = New System.Windows.Forms.GroupBox
        Me.Syringe = New System.Windows.Forms.Label
        Me.Station1 = New System.Windows.Forms.Label
        Me.Station2 = New System.Windows.Forms.Label
        Me.Station3 = New System.Windows.Forms.Label
        Me.Station3Label = New System.Windows.Forms.Label
        Me.Station2Label = New System.Windows.Forms.Label
        Me.Needle = New System.Windows.Forms.Label
        Me.Station1Label = New System.Windows.Forms.Label
        Me.NeedleLabel = New System.Windows.Forms.Label
        Me.SyringeLabel = New System.Windows.Forms.Label
        Me.PanelProDownTimeInfor = New System.Windows.Forms.Panel
        Me.RichTextBoxNote = New System.Windows.Forms.RichTextBox
        Me.CheckBoxPotOn = New System.Windows.Forms.CheckBox
        Me.DoorLock = New System.Windows.Forms.CheckBox
        Me.ButtonPotReset = New System.Windows.Forms.Button
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
        Me.panelVision = New System.Windows.Forms.Panel
        Me.TowerLightImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.TimeDisplayTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Panel2.SuspendLayout()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.ConveyorBox.SuspendLayout()
        Me.HeaterBox.SuspendLayout()
        Me.PanelProDownTimeInfor.SuspendLayout()
        Me.SuspendLayout()
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
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(312, 24)
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
        Me.Panel5.BackColor = System.Drawing.SystemColors.Control
        Me.Panel5.Controls.Add(Me.GroupBox2)
        Me.Panel5.Controls.Add(Me.GroupBox1)
        Me.Panel5.Controls.Add(Me.btExit)
        Me.Panel5.Controls.Add(Me.ConveyorBox)
        Me.Panel5.Controls.Add(Me.HeaterBox)
        Me.Panel5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel5.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Panel5.Location = New System.Drawing.Point(768, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(504, 992)
        Me.Panel5.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.tbTime)
        Me.GroupBox2.Controls.Add(Me.tbDispensedUnit)
        Me.GroupBox2.Controls.Add(Me.TextBox7)
        Me.GroupBox2.Controls.Add(Me.tbCurrentDispense)
        Me.GroupBox2.Controls.Add(Me.TextBox6)
        Me.GroupBox2.Controls.Add(Me.tbPortLifeRemain)
        Me.GroupBox2.Controls.Add(Me.TextBox5)
        Me.GroupBox2.Controls.Add(Me.tbDispensePressure)
        Me.GroupBox2.Controls.Add(Me.TextBox3)
        Me.GroupBox2.Controls.Add(Me.tbEquipmentID)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 216)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(480, 256)
        Me.GroupBox2.TabIndex = 143
        Me.GroupBox2.TabStop = False
        '
        'tbTime
        '
        Me.tbTime.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbTime.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbTime.Location = New System.Drawing.Point(16, 32)
        Me.tbTime.Name = "tbTime"
        Me.tbTime.ReadOnly = True
        Me.tbTime.Size = New System.Drawing.Size(448, 38)
        Me.tbTime.TabIndex = 11
        Me.tbTime.Text = ""
        Me.tbTime.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'tbDispensedUnit
        '
        Me.tbDispensedUnit.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbDispensedUnit.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDispensedUnit.Location = New System.Drawing.Point(196, 104)
        Me.tbDispensedUnit.Name = "tbDispensedUnit"
        Me.tbDispensedUnit.ReadOnly = True
        Me.tbDispensedUnit.Size = New System.Drawing.Size(272, 31)
        Me.tbDispensedUnit.TabIndex = 9
        Me.tbDispensedUnit.Text = "0"
        Me.tbDispensedUnit.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox7
        '
        Me.TextBox7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox7.Location = New System.Drawing.Point(12, 104)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = True
        Me.TextBox7.Size = New System.Drawing.Size(184, 31)
        Me.TextBox7.TabIndex = 8
        Me.TextBox7.Text = "Dispensed Unit :"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbCurrentDispense
        '
        Me.tbCurrentDispense.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbCurrentDispense.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCurrentDispense.Location = New System.Drawing.Point(196, 216)
        Me.tbCurrentDispense.Name = "tbCurrentDispense"
        Me.tbCurrentDispense.ReadOnly = True
        Me.tbCurrentDispense.Size = New System.Drawing.Size(272, 31)
        Me.tbCurrentDispense.TabIndex = 7
        Me.tbCurrentDispense.Text = "0"
        Me.tbCurrentDispense.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.tbCurrentDispense.Visible = False
        '
        'TextBox6
        '
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox6.Location = New System.Drawing.Point(12, 216)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = True
        Me.TextBox6.Size = New System.Drawing.Size(184, 31)
        Me.TextBox6.TabIndex = 6
        Me.TextBox6.Text = "Current Dispense :"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox6.Visible = False
        '
        'tbPortLifeRemain
        '
        Me.tbPortLifeRemain.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbPortLifeRemain.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPortLifeRemain.Location = New System.Drawing.Point(196, 168)
        Me.tbPortLifeRemain.Name = "tbPortLifeRemain"
        Me.tbPortLifeRemain.ReadOnly = True
        Me.tbPortLifeRemain.Size = New System.Drawing.Size(272, 31)
        Me.tbPortLifeRemain.TabIndex = 5
        Me.tbPortLifeRemain.Text = "NA"
        Me.tbPortLifeRemain.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox5
        '
        Me.TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox5.Location = New System.Drawing.Point(12, 168)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ReadOnly = True
        Me.TextBox5.Size = New System.Drawing.Size(184, 31)
        Me.TextBox5.TabIndex = 4
        Me.TextBox5.Text = "Port Life :"
        Me.TextBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbDispensePressure
        '
        Me.tbDispensePressure.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbDispensePressure.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbDispensePressure.Location = New System.Drawing.Point(196, 136)
        Me.tbDispensePressure.Name = "tbDispensePressure"
        Me.tbDispensePressure.ReadOnly = True
        Me.tbDispensePressure.Size = New System.Drawing.Size(272, 31)
        Me.tbDispensePressure.TabIndex = 3
        Me.tbDispensePressure.Text = "NA"
        Me.tbDispensePressure.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox3
        '
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(12, 136)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(184, 31)
        Me.TextBox3.TabIndex = 2
        Me.TextBox3.Text = "Pressure :"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbEquipmentID
        '
        Me.tbEquipmentID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.tbEquipmentID.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbEquipmentID.Location = New System.Drawing.Point(196, 72)
        Me.tbEquipmentID.Name = "tbEquipmentID"
        Me.tbEquipmentID.ReadOnly = True
        Me.tbEquipmentID.Size = New System.Drawing.Size(272, 31)
        Me.tbEquipmentID.TabIndex = 1
        Me.tbEquipmentID.Text = "ID12345"
        Me.tbEquipmentID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(12, 72)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(184, 31)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "Equipment ID :"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Panel1)
        Me.GroupBox1.Controls.Add(Me.btStop)
        Me.GroupBox1.Controls.Add(Me.btPause)
        Me.GroupBox1.Controls.Add(Me.btPlay)
        Me.GroupBox1.Controls.Add(Me.ContinuousMode)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Black
        Me.GroupBox1.Location = New System.Drawing.Point(16, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 216)
        Me.GroupBox1.TabIndex = 142
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Process"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GreenLight)
        Me.Panel1.Controls.Add(Me.AmberLight)
        Me.Panel1.Controls.Add(Me.RedLight)
        Me.Panel1.Location = New System.Drawing.Point(400, 16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(64, 192)
        Me.Panel1.TabIndex = 134
        '
        'GreenLight
        '
        Me.GreenLight.BackgroundImage = CType(resources.GetObject("GreenLight.BackgroundImage"), System.Drawing.Image)
        Me.GreenLight.Location = New System.Drawing.Point(0, 120)
        Me.GreenLight.Name = "GreenLight"
        Me.GreenLight.Size = New System.Drawing.Size(64, 64)
        Me.GreenLight.TabIndex = 2
        Me.GreenLight.TabStop = False
        '
        'AmberLight
        '
        Me.AmberLight.BackgroundImage = CType(resources.GetObject("AmberLight.BackgroundImage"), System.Drawing.Image)
        Me.AmberLight.Location = New System.Drawing.Point(0, 56)
        Me.AmberLight.Name = "AmberLight"
        Me.AmberLight.Size = New System.Drawing.Size(64, 64)
        Me.AmberLight.TabIndex = 1
        Me.AmberLight.TabStop = False
        '
        'RedLight
        '
        Me.RedLight.BackgroundImage = CType(resources.GetObject("RedLight.BackgroundImage"), System.Drawing.Image)
        Me.RedLight.Location = New System.Drawing.Point(0, 0)
        Me.RedLight.Name = "RedLight"
        Me.RedLight.Size = New System.Drawing.Size(64, 64)
        Me.RedLight.TabIndex = 0
        Me.RedLight.TabStop = False
        '
        'btStop
        '
        Me.btStop.ForeColor = System.Drawing.Color.Black
        Me.btStop.Image = CType(resources.GetObject("btStop.Image"), System.Drawing.Image)
        Me.btStop.Location = New System.Drawing.Point(256, 48)
        Me.btStop.Name = "btStop"
        Me.btStop.Size = New System.Drawing.Size(90, 90)
        Me.btStop.TabIndex = 137
        '
        'btPause
        '
        Me.btPause.ForeColor = System.Drawing.Color.Black
        Me.btPause.Image = CType(resources.GetObject("btPause.Image"), System.Drawing.Image)
        Me.btPause.Location = New System.Drawing.Point(160, 48)
        Me.btPause.Name = "btPause"
        Me.btPause.Size = New System.Drawing.Size(90, 90)
        Me.btPause.TabIndex = 136
        '
        'btPlay
        '
        Me.btPlay.ForeColor = System.Drawing.Color.Black
        Me.btPlay.Image = CType(resources.GetObject("btPlay.Image"), System.Drawing.Image)
        Me.btPlay.Location = New System.Drawing.Point(64, 48)
        Me.btPlay.Name = "btPlay"
        Me.btPlay.Size = New System.Drawing.Size(90, 90)
        Me.btPlay.TabIndex = 135
        '
        'ContinuousMode
        '
        Me.ContinuousMode.AutoCheck = False
        Me.ContinuousMode.ForeColor = System.Drawing.Color.Black
        Me.ContinuousMode.Location = New System.Drawing.Point(64, 160)
        Me.ContinuousMode.Name = "ContinuousMode"
        Me.ContinuousMode.Size = New System.Drawing.Size(160, 24)
        Me.ContinuousMode.TabIndex = 141
        Me.ContinuousMode.Text = "Continuous Run"
        '
        'btExit
        '
        Me.btExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btExit.ForeColor = System.Drawing.Color.Black
        Me.btExit.Location = New System.Drawing.Point(360, 904)
        Me.btExit.Name = "btExit"
        Me.btExit.Size = New System.Drawing.Size(120, 64)
        Me.btExit.TabIndex = 133
        Me.btExit.Text = "Exit"
        '
        'ConveyorBox
        '
        Me.ConveyorBox.BackColor = System.Drawing.SystemColors.Control
        Me.ConveyorBox.Controls.Add(Me.ResetPLCLogic)
        Me.ConveyorBox.Controls.Add(Me.ButtonStartFirstStage)
        Me.ConveyorBox.Controls.Add(Me.ButtonCV_Prod_Release)
        Me.ConveyorBox.Controls.Add(Me.ButtonCV_Prod_Retrieve)
        Me.ConveyorBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ConveyorBox.ForeColor = System.Drawing.Color.Black
        Me.ConveyorBox.Location = New System.Drawing.Point(16, 480)
        Me.ConveyorBox.Name = "ConveyorBox"
        Me.ConveyorBox.Size = New System.Drawing.Size(480, 96)
        Me.ConveyorBox.TabIndex = 130
        Me.ConveyorBox.TabStop = False
        Me.ConveyorBox.Text = "Conveyor Operation"
        '
        'ResetPLCLogic
        '
        Me.ResetPLCLogic.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResetPLCLogic.Location = New System.Drawing.Point(359, 24)
        Me.ResetPLCLogic.Name = "ResetPLCLogic"
        Me.ResetPLCLogic.Size = New System.Drawing.Size(104, 56)
        Me.ResetPLCLogic.TabIndex = 143
        Me.ResetPLCLogic.Text = "Reset PLC Logic"
        '
        'ButtonStartFirstStage
        '
        Me.ButtonStartFirstStage.Location = New System.Drawing.Point(17, 24)
        Me.ButtonStartFirstStage.Name = "ButtonStartFirstStage"
        Me.ButtonStartFirstStage.Size = New System.Drawing.Size(104, 56)
        Me.ButtonStartFirstStage.TabIndex = 142
        Me.ButtonStartFirstStage.Text = "Receive"
        '
        'ButtonCV_Prod_Release
        '
        Me.ButtonCV_Prod_Release.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Release.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Release.Location = New System.Drawing.Point(245, 24)
        Me.ButtonCV_Prod_Release.Name = "ButtonCV_Prod_Release"
        Me.ButtonCV_Prod_Release.Size = New System.Drawing.Size(104, 56)
        Me.ButtonCV_Prod_Release.TabIndex = 140
        Me.ButtonCV_Prod_Release.Text = "Release"
        '
        'ButtonCV_Prod_Retrieve
        '
        Me.ButtonCV_Prod_Retrieve.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonCV_Prod_Retrieve.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCV_Prod_Retrieve.Location = New System.Drawing.Point(131, 24)
        Me.ButtonCV_Prod_Retrieve.Name = "ButtonCV_Prod_Retrieve"
        Me.ButtonCV_Prod_Retrieve.Size = New System.Drawing.Size(104, 56)
        Me.ButtonCV_Prod_Retrieve.TabIndex = 139
        Me.ButtonCV_Prod_Retrieve.Text = "Retrieve"
        '
        'HeaterBox
        '
        Me.HeaterBox.BackColor = System.Drawing.SystemColors.Control
        Me.HeaterBox.Controls.Add(Me.Syringe)
        Me.HeaterBox.Controls.Add(Me.Station1)
        Me.HeaterBox.Controls.Add(Me.Station2)
        Me.HeaterBox.Controls.Add(Me.Station3)
        Me.HeaterBox.Controls.Add(Me.Station3Label)
        Me.HeaterBox.Controls.Add(Me.Station2Label)
        Me.HeaterBox.Controls.Add(Me.Needle)
        Me.HeaterBox.Controls.Add(Me.Station1Label)
        Me.HeaterBox.Controls.Add(Me.NeedleLabel)
        Me.HeaterBox.Controls.Add(Me.SyringeLabel)
        Me.HeaterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.HeaterBox.ForeColor = System.Drawing.Color.Black
        Me.HeaterBox.Location = New System.Drawing.Point(16, 592)
        Me.HeaterBox.Name = "HeaterBox"
        Me.HeaterBox.Size = New System.Drawing.Size(480, 160)
        Me.HeaterBox.TabIndex = 130
        Me.HeaterBox.TabStop = False
        Me.HeaterBox.Text = "Thermal Readings"
        Me.HeaterBox.Visible = False
        '
        'Syringe
        '
        Me.Syringe.ForeColor = System.Drawing.Color.Red
        Me.Syringe.Location = New System.Drawing.Point(304, 32)
        Me.Syringe.Name = "Syringe"
        Me.Syringe.Size = New System.Drawing.Size(56, 24)
        Me.Syringe.TabIndex = 143
        Me.Syringe.Text = "50 oC"
        Me.Syringe.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Station1
        '
        Me.Station1.ForeColor = System.Drawing.Color.Red
        Me.Station1.Location = New System.Drawing.Point(72, 128)
        Me.Station1.Name = "Station1"
        Me.Station1.Size = New System.Drawing.Size(56, 24)
        Me.Station1.TabIndex = 143
        Me.Station1.Text = "50 oC"
        Me.Station1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Station2
        '
        Me.Station2.ForeColor = System.Drawing.Color.Red
        Me.Station2.Location = New System.Drawing.Point(208, 128)
        Me.Station2.Name = "Station2"
        Me.Station2.Size = New System.Drawing.Size(56, 24)
        Me.Station2.TabIndex = 143
        Me.Station2.Text = "50 oC"
        Me.Station2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Station3
        '
        Me.Station3.ForeColor = System.Drawing.Color.Red
        Me.Station3.Location = New System.Drawing.Point(352, 128)
        Me.Station3.Name = "Station3"
        Me.Station3.Size = New System.Drawing.Size(56, 24)
        Me.Station3.TabIndex = 143
        Me.Station3.Text = "50 oC"
        Me.Station3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Station3Label
        '
        Me.Station3Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station3Label.Location = New System.Drawing.Point(320, 104)
        Me.Station3Label.Name = "Station3Label"
        Me.Station3Label.Size = New System.Drawing.Size(112, 16)
        Me.Station3Label.TabIndex = 26
        Me.Station3Label.Text = "Post Heater"
        Me.Station3Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Station2Label
        '
        Me.Station2Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station2Label.Location = New System.Drawing.Point(184, 104)
        Me.Station2Label.Name = "Station2Label"
        Me.Station2Label.Size = New System.Drawing.Size(112, 16)
        Me.Station2Label.TabIndex = 25
        Me.Station2Label.Text = "Disp. Heater"
        Me.Station2Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Needle
        '
        Me.Needle.ForeColor = System.Drawing.Color.Red
        Me.Needle.Location = New System.Drawing.Point(112, 32)
        Me.Needle.Name = "Needle"
        Me.Needle.Size = New System.Drawing.Size(56, 24)
        Me.Needle.TabIndex = 143
        Me.Needle.Text = "50 oC"
        Me.Needle.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Station1Label
        '
        Me.Station1Label.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Station1Label.Location = New System.Drawing.Point(48, 104)
        Me.Station1Label.Name = "Station1Label"
        Me.Station1Label.Size = New System.Drawing.Size(112, 16)
        Me.Station1Label.TabIndex = 24
        Me.Station1Label.Text = "Pre Heater"
        Me.Station1Label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        '
        'PanelProDownTimeInfor
        '
        Me.PanelProDownTimeInfor.BackColor = System.Drawing.SystemColors.Control
        Me.PanelProDownTimeInfor.Controls.Add(Me.RichTextBoxNote)
        Me.PanelProDownTimeInfor.Controls.Add(Me.CheckBoxPotOn)
        Me.PanelProDownTimeInfor.Controls.Add(Me.DoorLock)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPotReset)
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
        Me.RichTextBoxNote.Location = New System.Drawing.Point(8, 128)
        Me.RichTextBoxNote.Name = "RichTextBoxNote"
        Me.RichTextBoxNote.ReadOnly = True
        Me.RichTextBoxNote.Size = New System.Drawing.Size(752, 208)
        Me.RichTextBoxNote.TabIndex = 98
        Me.RichTextBoxNote.Text = ""
        '
        'CheckBoxPotOn
        '
        Me.CheckBoxPotOn.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxPotOn.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxPotOn.Enabled = False
        Me.CheckBoxPotOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxPotOn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CheckBoxPotOn.ImageList = Me.ImageListPotEtc
        Me.CheckBoxPotOn.Location = New System.Drawing.Point(480, 8)
        Me.CheckBoxPotOn.Name = "CheckBoxPotOn"
        Me.CheckBoxPotOn.Size = New System.Drawing.Size(75, 56)
        Me.CheckBoxPotOn.TabIndex = 119
        Me.CheckBoxPotOn.Text = "Pot Life On"
        Me.CheckBoxPotOn.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'DoorLock
        '
        Me.DoorLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.DoorLock.BackColor = System.Drawing.SystemColors.Control
        Me.DoorLock.Checked = True
        Me.DoorLock.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.DoorLock.ImageList = Me.ImageListPotEtc
        Me.DoorLock.Location = New System.Drawing.Point(672, 8)
        Me.DoorLock.Name = "DoorLock"
        Me.DoorLock.Size = New System.Drawing.Size(75, 56)
        Me.DoorLock.TabIndex = 116
        Me.DoorLock.Text = "Lock Door"
        Me.DoorLock.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonPotReset
        '
        Me.ButtonPotReset.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonPotReset.Enabled = False
        Me.ButtonPotReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonPotReset.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPotReset.ImageList = Me.ImageListPotEtc
        Me.ButtonPotReset.Location = New System.Drawing.Point(560, 8)
        Me.ButtonPotReset.Name = "ButtonPotReset"
        Me.ButtonPotReset.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPotReset.TabIndex = 105
        Me.ButtonPotReset.Text = "Reset Pot"
        '
        'ButtonOpenFile
        '
        Me.ButtonOpenFile.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonOpenFile.Enabled = False
        Me.ButtonOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonOpenFile.Location = New System.Drawing.Point(680, 80)
        Me.ButtonOpenFile.Name = "ButtonOpenFile"
        Me.ButtonOpenFile.Size = New System.Drawing.Size(64, 27)
        Me.ButtonOpenFile.TabIndex = 96
        Me.ButtonOpenFile.Text = "..."
        '
        'TextBoxFilename
        '
        Me.TextBoxFilename.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxFilename.Enabled = False
        Me.TextBoxFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBoxFilename.Location = New System.Drawing.Point(56, 80)
        Me.TextBoxFilename.Name = "TextBoxFilename"
        Me.TextBoxFilename.ReadOnly = True
        Me.TextBoxFilename.Size = New System.Drawing.Size(608, 27)
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
        Me.Label5.Text = "File :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonNdlCalib
        '
        Me.ButtonNdlCalib.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonNdlCalib.Enabled = False
        Me.ButtonNdlCalib.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNdlCalib.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonNdlCalib.ImageList = Me.ImageListGeneralTools
        Me.ButtonNdlCalib.Location = New System.Drawing.Point(308, 8)
        Me.ButtonNdlCalib.Name = "ButtonNdlCalib"
        Me.ButtonNdlCalib.Size = New System.Drawing.Size(75, 56)
        Me.ButtonNdlCalib.TabIndex = 91
        Me.ButtonNdlCalib.Text = "Need. Cal."
        '
        'ButtonClean
        '
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonClean.Enabled = False
        Me.ButtonClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClean.ImageList = Me.ImageListGeneralTools
        Me.ButtonClean.Location = New System.Drawing.Point(158, 8)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.Size = New System.Drawing.Size(75, 56)
        Me.ButtonClean.TabIndex = 89
        Me.ButtonClean.Text = "Clean On"
        '
        'ButtonPurge
        '
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonPurge.Enabled = False
        Me.ButtonPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPurge.ImageList = Me.ImageListGeneralTools
        Me.ButtonPurge.Location = New System.Drawing.Point(83, 8)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPurge.TabIndex = 88
        Me.ButtonPurge.Text = "Purge On"
        '
        'ButtonChgSyringe
        '
        Me.ButtonChgSyringe.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonChgSyringe.Enabled = False
        Me.ButtonChgSyringe.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonChgSyringe.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonChgSyringe.ImageList = Me.ImageListGeneralTools
        Me.ButtonChgSyringe.Location = New System.Drawing.Point(383, 8)
        Me.ButtonChgSyringe.Name = "ButtonChgSyringe"
        Me.ButtonChgSyringe.Size = New System.Drawing.Size(75, 56)
        Me.ButtonChgSyringe.TabIndex = 87
        Me.ButtonChgSyringe.Text = "Change Syringe"
        '
        'ButtonHome
        '
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonHome.Enabled = False
        Me.ButtonHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonHome.ImageList = Me.ImageListGeneralTools
        Me.ButtonHome.Location = New System.Drawing.Point(8, 8)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(75, 56)
        Me.ButtonHome.TabIndex = 86
        Me.ButtonHome.Text = "Home"
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Control
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Black
        Me.LabelMessege.Location = New System.Drawing.Point(8, 344)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(752, 32)
        Me.LabelMessege.TabIndex = 85
        Me.LabelMessege.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonVolCalib
        '
        Me.ButtonVolCalib.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonVolCalib.Enabled = False
        Me.ButtonVolCalib.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolCalib.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonVolCalib.ImageList = Me.ImageListGeneralTools
        Me.ButtonVolCalib.Location = New System.Drawing.Point(233, 8)
        Me.ButtonVolCalib.Name = "ButtonVolCalib"
        Me.ButtonVolCalib.Size = New System.Drawing.Size(75, 56)
        Me.ButtonVolCalib.TabIndex = 90
        Me.ButtonVolCalib.Text = "Vol. Cal."
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
        'panelVision
        '
        Me.panelVision.BackColor = System.Drawing.Color.SlateGray
        Me.panelVision.Location = New System.Drawing.Point(0, 384)
        Me.panelVision.Name = "panelVision"
        Me.panelVision.Size = New System.Drawing.Size(768, 576)
        Me.panelVision.TabIndex = 7
        '
        'TowerLightImageList
        '
        Me.TowerLightImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth24Bit
        Me.TowerLightImageList.ImageSize = New System.Drawing.Size(64, 64)
        Me.TowerLightImageList.ImageStream = CType(resources.GetObject("TowerLightImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.TowerLightImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimeDisplayTimer
        '
        Me.TimeDisplayTimer.Interval = 500
        '
        'FormProduction
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1276, 990)
        Me.Controls.Add(Me.panelVision)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.PanelProDownTimeInfor)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Menu = Me.MainMenuProduction
        Me.Name = "FormProduction"
        Me.Text = "Production"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel2.ResumeLayout(False)
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ConveyorBox.ResumeLayout(False)
        Me.HeaterBox.ResumeLayout(False)
        Me.PanelProDownTimeInfor.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TimeDisplayTimer.Start()
        Init()
        While isInited = False
            Application.DoEvents()
        End While
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
        panelVision.Controls.Add(Vision.FrmVision.PanelVision) 'lsgoh
        SwitchToTeachCamera()
        Try
            ValueBrightness.Value = IDS.Data.Hardware.Camera.Brightness
        Catch ex As Exception
            ValueBrightness.Value = ValueBrightness.Minimum
        End Try


        'motion controller
        'm_Tri.Connect_Controller()
        SetState("Homing")

        'timers start
        IDS.StartErrorCheck()
        TimerMonitor.Enabled = True
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
        isInited = False
        InitThread.Abort()
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
        tbDispensePressure.Text = IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure.ToString("0.00") + " Bar"
        LabelMessage("File loaded")

        If IDS.Data.Hardware.Dispenser.Left.PotLifeOption Then
            ButtonPotReset.Enabled = True
            CheckBoxPotOn.Enabled = True
        End If
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

    Private Sub ButtonHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHome.Click
        forceHome = True
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
        If Me.ButtonVolCalib.Text = "Vol. Cal." Then
            SetState("Volume Calibration")
            Me.ButtonVolCalib.Text = "Stop Vol. Cal."
        Else
            Me.ButtonVolCalib.Enabled = False
            Me.ButtonVolCalib.Text = "Vol. Cal."
            StopDispensing()
        End If

    End Sub

    Private Sub ButtonNdlCalib_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNdlCalib.Click
        If Me.ButtonNdlCalib.Text = "Need. Cal." Then
            SetState("Needle Calibration")
            Me.ButtonNdlCalib.Text = "Stop Need. Cal."
        Else
            Me.ButtonNdlCalib.Enabled = False
            Me.ButtonNdlCalib.Text = "Need. Cal."
            StopDispensing()
        End If
    End Sub

    Private Sub DoorLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoorLock.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't unlock the door when machine running!")
            DoorLock.Text = "Unlock Door"
            'DoorLock.ImageIndex = 2
            Exit Sub
        End If

        If DoorLock.Checked = False Then
            DoorLock.Text = "Unlock Door"
            'DoorLock.ImageIndex = 2
            LockDoor()
        Else
            DoorLock.Text = "Lock Door"
            'DoorLock.ImageIndex = 3
            UnlockDoor()
        End If
    End Sub

    Private Sub CheckBoxPotOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPotOn.CheckedChanged

        If CheckBoxPotOn.Checked = True Then
            CheckBoxPotOn.Text = "Pot Life Off"
            m_PotLifeExpire = False
            m_PotLifeOn = True
            If RemainingTimeOfPotLifeStart = Nothing Then
                ResetTimer("Reset Potlife Start Timer")
            End If
        Else
            CheckBoxPotOn.Text = "Pot Life On"
            'CheckBoxPotOn.ImageIndex = 6
            m_PotLifeOn = False
            m_PotLifeExpire = False
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
        End If
        ContinuousMode.Checked = Not ContinuousMode.Checked
        If ContinuousMode.Checked = False Then
            Conveyor.Command("Manual Mode")
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned off.")
            End If
        Else
            Conveyor.Command("Auto Mode")
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned on.")
            End If
        End If
        If ContinuousMode.Checked Then
            ButtonStartFirstStage.Enabled = True
            ButtonCV_Prod_Retrieve.Enabled = True
            ButtonCV_Prod_Release.Enabled = True
            ResetPLCLogic.Enabled = True
        Else
            ButtonStartFirstStage.Enabled = False
            ButtonCV_Prod_Retrieve.Enabled = False
            ButtonCV_Prod_Release.Enabled = False
            ResetPLCLogic.Enabled = False
        End If
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
        productCnt += 1
        tbDispensedUnit.Text = productCnt
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

    Private Sub TimerMonitor_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMonitor.Tick

        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag
        Dim LowMaterialSensorLeft As Boolean = IDS.Devices.DIO.DIO.LevelSensor(True)

        'reset the timers if 
        '1) door is not closed when door interlocked is true
        '2) program is currently paused or running
        '3) conveyor mode is not ticked
        'If (door_close = False And door_interlock = True) Or ContinuousMode.Checked = False And IsIdle() Then
        '    ResetTimer("Reset Autopurging Timers")
        '    ResetTimer("Reset Potlife Start Timer")
        '    Exit Sub
        'End If

        Dim auto_purging_enabled As Boolean = IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption And ContinuousMode.Checked And IsIdle() And Form_Service.Visible = False
        CurrentTime = Now

        If m_PotLifeOn Then
            ElapsedTimeSincePotLifeStart = (CurrentTime.Ticks - RemainingTimeOfPotLifeStart.Ticks) * tickToMinute

            If ElapsedTimeSincePotLifeStart >= IDS.Data.Hardware.Dispenser.Left.PotLifeDuration Then
                'what do we want to do when potlife is up?
                tbPortLifeRemain.Text = "Expired"
                m_PotLifeExpire = True

            Else
                tbPortLifeRemain.Text = (IDS.Data.Hardware.Dispenser.Left.PotLifeDuration - ElapsedTimeSincePotLifeStart).ToString() + " Minutes"
            End If
        Else
            tbPortLifeRemain.Text = "NA"
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
        panelVision.Controls.Remove(Vision.FrmVision.PanelVision) 'lim

        'timers stop
        IDS.StopErrorCheck()
        TimerMonitor.Stop()
        Programming.IOCheck.Stop()
        If Not (ThreadMonitor) Is Nothing Then
            ThreadMonitor.Abort()
        End If
        If Not (ThreadExecutor) Is Nothing Then
            ThreadExecutor.Abort()
        End If
        Conveyor.PositionTimer.Stop()

        'm_Tri.Move_Z(0)
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

    Private Sub btPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPlay.Click
        SetState("Start")
    End Sub

    Private Sub btPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPause.Click
        PauseDispensing()
    End Sub

    Private Sub btStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStop.Click
        StopDispensing()
    End Sub

    Private Sub RichTextBoxNote_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBoxNote.TextChanged

    End Sub

    Private Sub TimeDisplayTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeDisplayTimer.Tick
        tbTime.Text = DateTime.Today.ToShortDateString() + TimeOfDay.ToString(" h:mm:ss tt")
    End Sub
End Class
