Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
'Imports DLL_SetupAndCalibrate
Imports System.Threading
'Imports Microsoft.win32
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

    Private isJogON As Boolean = False
    Shared mouse_pos As New Point
    Shared cursor_hide As Boolean = False
    Shared isPress As Boolean
    Dim deadzone As Integer = 3
    Dim jogspeed As Integer = 0
    Const maxSpeed As Integer = 100
    Const maxMouseRangeX = 600.0 '6000.0 '4500.0
    Const maxMouseRangeY = 300.0 '2000.0
    Const ratioLB = 0.4
    Const ratioUB = 1.05
    Dim ratio As Double
    Dim countMouseTimer As Integer = 0
    Private m_TrackBall As New DLL_Export_Device_Motor.Mouse(Me)
    'Private m_keyBoard As New DLL_Export_Device_Motor.Keyboard(Me)

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
        HardwareInitTimer.Enabled = False
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
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBoxRobotPos As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PanelProDownTimeInfor As System.Windows.Forms.Panel
    Friend WithEvents RichTextBoxNote As System.Windows.Forms.RichTextBox
    Friend WithEvents CheckBoxPotOn As System.Windows.Forms.CheckBox
    Friend WithEvents DoorLock As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonPotReset As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonOpenFile As System.Windows.Forms.Button
    Friend WithEvents TextBoxFilename As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButtonClean As System.Windows.Forms.Button
    Friend WithEvents ButtonPurge As System.Windows.Forms.Button
    Friend WithEvents ButtonChgSyringe As System.Windows.Forms.Button
    Friend WithEvents ButtonHome As System.Windows.Forms.Button
    Friend WithEvents LabelMessege As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ImageListOperation As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TimerMonitor As System.Windows.Forms.Timer
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
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
    Friend WithEvents ButtonCalibrate As System.Windows.Forms.Button
    Friend WithEvents ContinuousMode As System.Windows.Forms.CheckBox
    Public WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents btExit As System.Windows.Forms.Button
    Friend WithEvents HardwareInitTimer As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents pbRedLight As System.Windows.Forms.PictureBox
    Friend WithEvents pbAmberLight As System.Windows.Forms.PictureBox
    Friend WithEvents pbGreenLight As System.Windows.Forms.PictureBox
    Friend WithEvents TowerLightImageList As System.Windows.Forms.ImageList
    Friend WithEvents btStop As System.Windows.Forms.Button
    Friend WithEvents btPause As System.Windows.Forms.Button
    Friend WithEvents btStart As System.Windows.Forms.Button
    Friend WithEvents imageListProcessBt As System.Windows.Forms.ImageList
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents tbEquipmentID As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents tbOperatorID As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents tbTotalDispense As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents tbCurrDispense As System.Windows.Forms.TextBox
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents tbQuantity As System.Windows.Forms.TextBox
    Friend WithEvents ImageListButtons As System.Windows.Forms.ImageList
    Friend WithEvents gbSeparator As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormProduction))
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListPotEtc = New System.Windows.Forms.ImageList(Me.components)
        Me.MainMenuProduction = New System.Windows.Forms.MainMenu
        Me.ImageListGeneralTools = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TextBoxRobotPos = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.gbSeparator = New System.Windows.Forms.GroupBox
        Me.TextBox10 = New System.Windows.Forms.TextBox
        Me.tbQuantity = New System.Windows.Forms.TextBox
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.tbCurrDispense = New System.Windows.Forms.TextBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.tbTotalDispense = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.tbOperatorID = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.tbEquipmentID = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.btStart = New System.Windows.Forms.Button
        Me.imageListProcessBt = New System.Windows.Forms.ImageList(Me.components)
        Me.btStop = New System.Windows.Forms.Button
        Me.btPause = New System.Windows.Forms.Button
        Me.ContinuousMode = New System.Windows.Forms.CheckBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.pbGreenLight = New System.Windows.Forms.PictureBox
        Me.pbAmberLight = New System.Windows.Forms.PictureBox
        Me.pbRedLight = New System.Windows.Forms.PictureBox
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.btExit = New System.Windows.Forms.Button
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.PanelProDownTimeInfor = New System.Windows.Forms.Panel
        Me.RichTextBoxNote = New System.Windows.Forms.RichTextBox
        Me.CheckBoxPotOn = New System.Windows.Forms.CheckBox
        Me.DoorLock = New System.Windows.Forms.CheckBox
        Me.ButtonPotReset = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonOpenFile = New System.Windows.Forms.Button
        Me.TextBoxFilename = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonClean = New System.Windows.Forms.Button
        Me.ButtonPurge = New System.Windows.Forms.Button
        Me.ButtonChgSyringe = New System.Windows.Forms.Button
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ImageListOperation = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TimerMonitor = New System.Windows.Forms.Timer(Me.components)
        Me.HardwareInitTimer = New System.Windows.Forms.Timer(Me.components)
        Me.TowerLightImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListButtons = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
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
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.TextBoxRobotPos)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 960)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(768, 32)
        Me.Panel2.TabIndex = 4
        '
        'TextBoxRobotPos
        '
        Me.TextBoxRobotPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotPos.Location = New System.Drawing.Point(496, 3)
        Me.TextBoxRobotPos.Name = "TextBoxRobotPos"
        Me.TextBoxRobotPos.ReadOnly = True
        Me.TextBoxRobotPos.Size = New System.Drawing.Size(256, 21)
        Me.TextBoxRobotPos.TabIndex = 8
        Me.TextBoxRobotPos.Text = "X: -100.000,   Y: -100.000,  Z: -100.000"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(440, 5)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'Panel5
        '
        Me.Panel5.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel5.BackColor = System.Drawing.SystemColors.Control
        Me.Panel5.Controls.Add(Me.gbSeparator)
        Me.Panel5.Controls.Add(Me.TextBox10)
        Me.Panel5.Controls.Add(Me.tbQuantity)
        Me.Panel5.Controls.Add(Me.TextBox8)
        Me.Panel5.Controls.Add(Me.tbCurrDispense)
        Me.Panel5.Controls.Add(Me.TextBox5)
        Me.Panel5.Controls.Add(Me.tbTotalDispense)
        Me.Panel5.Controls.Add(Me.TextBox6)
        Me.Panel5.Controls.Add(Me.tbOperatorID)
        Me.Panel5.Controls.Add(Me.TextBox4)
        Me.Panel5.Controls.Add(Me.tbEquipmentID)
        Me.Panel5.Controls.Add(Me.TextBox3)
        Me.Panel5.Controls.Add(Me.GroupBox2)
        Me.Panel5.Controls.Add(Me.Panel1)
        Me.Panel5.Controls.Add(Me.PanelToBeAdded)
        Me.Panel5.Controls.Add(Me.TextBox2)
        Me.Panel5.Controls.Add(Me.GroupBox1)
        Me.Panel5.Controls.Add(Me.btExit)
        Me.Panel5.Controls.Add(Me.PanelVision)
        Me.Panel5.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Panel5.Location = New System.Drawing.Point(768, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(504, 960)
        Me.Panel5.TabIndex = 0
        '
        'gbSeparator
        '
        Me.gbSeparator.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.gbSeparator.Location = New System.Drawing.Point(0, -8)
        Me.gbSeparator.Name = "gbSeparator"
        Me.gbSeparator.Size = New System.Drawing.Size(8, 1016)
        Me.gbSeparator.TabIndex = 153
        Me.gbSeparator.TabStop = False
        '
        'TextBox10
        '
        Me.TextBox10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox10.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox10.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox10.Location = New System.Drawing.Point(40, 432)
        Me.TextBox10.Name = "TextBox10"
        Me.TextBox10.ReadOnly = True
        Me.TextBox10.Size = New System.Drawing.Size(176, 31)
        Me.TextBox10.TabIndex = 152
        Me.TextBox10.Text = "Quantity:"
        Me.TextBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbQuantity
        '
        Me.tbQuantity.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbQuantity.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbQuantity.Location = New System.Drawing.Point(216, 432)
        Me.tbQuantity.Name = "tbQuantity"
        Me.tbQuantity.ReadOnly = True
        Me.tbQuantity.Size = New System.Drawing.Size(216, 31)
        Me.tbQuantity.TabIndex = 151
        Me.tbQuantity.Text = "0"
        Me.tbQuantity.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox8
        '
        Me.TextBox8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox8.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox8.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox8.Location = New System.Drawing.Point(40, 392)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.ReadOnly = True
        Me.TextBox8.Size = New System.Drawing.Size(176, 31)
        Me.TextBox8.TabIndex = 150
        Me.TextBox8.Text = "Current Dispense:"
        Me.TextBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbCurrDispense
        '
        Me.tbCurrDispense.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbCurrDispense.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCurrDispense.Location = New System.Drawing.Point(216, 392)
        Me.tbCurrDispense.Name = "tbCurrDispense"
        Me.tbCurrDispense.ReadOnly = True
        Me.tbCurrDispense.Size = New System.Drawing.Size(216, 31)
        Me.tbCurrDispense.TabIndex = 149
        Me.tbCurrDispense.Text = "0"
        Me.tbCurrDispense.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox5
        '
        Me.TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox5.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox5.Location = New System.Drawing.Point(40, 352)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ReadOnly = True
        Me.TextBox5.Size = New System.Drawing.Size(176, 31)
        Me.TextBox5.TabIndex = 148
        Me.TextBox5.Text = "Total Dispense:"
        Me.TextBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbTotalDispense
        '
        Me.tbTotalDispense.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbTotalDispense.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbTotalDispense.Location = New System.Drawing.Point(216, 352)
        Me.tbTotalDispense.Name = "tbTotalDispense"
        Me.tbTotalDispense.ReadOnly = True
        Me.tbTotalDispense.Size = New System.Drawing.Size(216, 31)
        Me.tbTotalDispense.TabIndex = 147
        Me.tbTotalDispense.Text = "10000"
        Me.tbTotalDispense.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox6
        '
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox6.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox6.Location = New System.Drawing.Point(40, 312)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = True
        Me.TextBox6.Size = New System.Drawing.Size(176, 31)
        Me.TextBox6.TabIndex = 146
        Me.TextBox6.Text = "Operator ID:"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbOperatorID
        '
        Me.tbOperatorID.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbOperatorID.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbOperatorID.Location = New System.Drawing.Point(216, 312)
        Me.tbOperatorID.Name = "tbOperatorID"
        Me.tbOperatorID.ReadOnly = True
        Me.tbOperatorID.Size = New System.Drawing.Size(216, 31)
        Me.tbOperatorID.TabIndex = 145
        Me.tbOperatorID.Text = "OpID1122"
        Me.tbOperatorID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox4
        '
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox4.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(40, 272)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(176, 31)
        Me.TextBox4.TabIndex = 144
        Me.TextBox4.Text = "Equipment ID:"
        Me.TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'tbEquipmentID
        '
        Me.tbEquipmentID.Cursor = System.Windows.Forms.Cursors.Default
        Me.tbEquipmentID.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbEquipmentID.Location = New System.Drawing.Point(216, 272)
        Me.tbEquipmentID.Name = "tbEquipmentID"
        Me.tbEquipmentID.ReadOnly = True
        Me.tbEquipmentID.Size = New System.Drawing.Size(216, 31)
        Me.tbEquipmentID.TabIndex = 143
        Me.tbEquipmentID.Text = "ID1288"
        Me.tbEquipmentID.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox3
        '
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(40, 232)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(176, 31)
        Me.TextBox3.TabIndex = 142
        Me.TextBox3.Text = "Machine State:"
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btStart)
        Me.GroupBox2.Controls.Add(Me.btStop)
        Me.GroupBox2.Controls.Add(Me.btPause)
        Me.GroupBox2.Controls.Add(Me.ContinuousMode)
        Me.GroupBox2.ForeColor = System.Drawing.Color.Black
        Me.GroupBox2.Location = New System.Drawing.Point(40, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(384, 176)
        Me.GroupBox2.TabIndex = 141
        Me.GroupBox2.TabStop = False
        '
        'btStart
        '
        Me.btStart.BackColor = System.Drawing.SystemColors.Control
        Me.btStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btStart.ImageIndex = 0
        Me.btStart.ImageList = Me.imageListProcessBt
        Me.btStart.Location = New System.Drawing.Point(46, 32)
        Me.btStart.Name = "btStart"
        Me.btStart.Size = New System.Drawing.Size(84, 80)
        Me.btStart.TabIndex = 138
        Me.btStart.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'imageListProcessBt
        '
        Me.imageListProcessBt.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.imageListProcessBt.ImageSize = New System.Drawing.Size(64, 64)
        Me.imageListProcessBt.ImageStream = CType(resources.GetObject("imageListProcessBt.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListProcessBt.TransparentColor = System.Drawing.Color.Transparent
        '
        'btStop
        '
        Me.btStop.BackColor = System.Drawing.SystemColors.Control
        Me.btStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btStop.ImageIndex = 2
        Me.btStop.ImageList = Me.imageListProcessBt
        Me.btStop.Location = New System.Drawing.Point(254, 32)
        Me.btStop.Name = "btStop"
        Me.btStop.Size = New System.Drawing.Size(84, 80)
        Me.btStop.TabIndex = 140
        Me.btStop.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'btPause
        '
        Me.btPause.BackColor = System.Drawing.SystemColors.Control
        Me.btPause.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btPause.ImageIndex = 1
        Me.btPause.ImageList = Me.imageListProcessBt
        Me.btPause.Location = New System.Drawing.Point(150, 32)
        Me.btPause.Name = "btPause"
        Me.btPause.Size = New System.Drawing.Size(84, 80)
        Me.btPause.TabIndex = 139
        Me.btPause.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ContinuousMode
        '
        Me.ContinuousMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContinuousMode.ForeColor = System.Drawing.Color.Black
        Me.ContinuousMode.Location = New System.Drawing.Point(48, 120)
        Me.ContinuousMode.Name = "ContinuousMode"
        Me.ContinuousMode.Size = New System.Drawing.Size(176, 40)
        Me.ContinuousMode.TabIndex = 133
        Me.ContinuousMode.Text = "Continuous Mode"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.pbGreenLight)
        Me.Panel1.Controls.Add(Me.pbAmberLight)
        Me.Panel1.Controls.Add(Me.pbRedLight)
        Me.Panel1.Location = New System.Drawing.Point(432, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(64, 205)
        Me.Panel1.TabIndex = 137
        '
        'pbGreenLight
        '
        Me.pbGreenLight.Image = CType(resources.GetObject("pbGreenLight.Image"), System.Drawing.Image)
        Me.pbGreenLight.Location = New System.Drawing.Point(0, 128)
        Me.pbGreenLight.Name = "pbGreenLight"
        Me.pbGreenLight.Size = New System.Drawing.Size(64, 64)
        Me.pbGreenLight.TabIndex = 2
        Me.pbGreenLight.TabStop = False
        '
        'pbAmberLight
        '
        Me.pbAmberLight.Image = CType(resources.GetObject("pbAmberLight.Image"), System.Drawing.Image)
        Me.pbAmberLight.Location = New System.Drawing.Point(0, 64)
        Me.pbAmberLight.Name = "pbAmberLight"
        Me.pbAmberLight.Size = New System.Drawing.Size(64, 64)
        Me.pbAmberLight.TabIndex = 1
        Me.pbAmberLight.TabStop = False
        '
        'pbRedLight
        '
        Me.pbRedLight.Image = CType(resources.GetObject("pbRedLight.Image"), System.Drawing.Image)
        Me.pbRedLight.Location = New System.Drawing.Point(0, 0)
        Me.pbRedLight.Name = "pbRedLight"
        Me.pbRedLight.Size = New System.Drawing.Size(64, 64)
        Me.pbRedLight.TabIndex = 0
        Me.pbRedLight.TabStop = False
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.Color.LightSteelBlue
        Me.PanelToBeAdded.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(84, 528)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(336, 280)
        Me.PanelToBeAdded.TabIndex = 134
        Me.PanelToBeAdded.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Cursor = System.Windows.Forms.Cursors.Default
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(216, 232)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(216, 31)
        Me.TextBox2.TabIndex = 132
        Me.TextBox2.Text = "Idle"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.ComboBox3)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBox1.Location = New System.Drawing.Point(16, 680)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 216)
        Me.GroupBox1.TabIndex = 130
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Test"
        Me.GroupBox1.Visible = False
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
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(280, 136)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(184, 27)
        Me.TextBox1.TabIndex = 132
        Me.TextBox1.Text = "prev state"
        Me.TextBox1.Visible = False
        '
        'btExit
        '
        Me.btExit.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.btExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btExit.ForeColor = System.Drawing.Color.Black
        Me.btExit.Location = New System.Drawing.Point(392, 880)
        Me.btExit.Name = "btExit"
        Me.btExit.Size = New System.Drawing.Size(96, 64)
        Me.btExit.TabIndex = 136
        Me.btExit.Text = "Exit"
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.Location = New System.Drawing.Point(24, 624)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(48, 48)
        Me.PanelVision.TabIndex = 7
        Me.PanelVision.Visible = False
        '
        'PanelProDownTimeInfor
        '
        Me.PanelProDownTimeInfor.BackColor = System.Drawing.SystemColors.Control
        Me.PanelProDownTimeInfor.Controls.Add(Me.RichTextBoxNote)
        Me.PanelProDownTimeInfor.Controls.Add(Me.CheckBoxPotOn)
        Me.PanelProDownTimeInfor.Controls.Add(Me.DoorLock)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPotReset)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label6)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonOpenFile)
        Me.PanelProDownTimeInfor.Controls.Add(Me.TextBoxFilename)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label5)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonClean)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPurge)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonChgSyringe)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonHome)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonCalibrate)
        Me.PanelProDownTimeInfor.Controls.Add(Me.LabelMessege)
        Me.PanelProDownTimeInfor.Location = New System.Drawing.Point(0, 0)
        Me.PanelProDownTimeInfor.Name = "PanelProDownTimeInfor"
        Me.PanelProDownTimeInfor.Size = New System.Drawing.Size(768, 960)
        Me.PanelProDownTimeInfor.TabIndex = 6
        '
        'RichTextBoxNote
        '
        Me.RichTextBoxNote.Cursor = System.Windows.Forms.Cursors.Arrow
        Me.RichTextBoxNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBoxNote.Location = New System.Drawing.Point(16, 184)
        Me.RichTextBoxNote.Name = "RichTextBoxNote"
        Me.RichTextBoxNote.ReadOnly = True
        Me.RichTextBoxNote.Size = New System.Drawing.Size(728, 720)
        Me.RichTextBoxNote.TabIndex = 98
        Me.RichTextBoxNote.Text = ""
        '
        'CheckBoxPotOn
        '
        Me.CheckBoxPotOn.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxPotOn.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBoxPotOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBoxPotOn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CheckBoxPotOn.ImageIndex = 6
        Me.CheckBoxPotOn.Location = New System.Drawing.Point(526, 8)
        Me.CheckBoxPotOn.Name = "CheckBoxPotOn"
        Me.CheckBoxPotOn.Size = New System.Drawing.Size(80, 80)
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
        Me.DoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.DoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.DoorLock.ImageIndex = 3
        Me.DoorLock.Location = New System.Drawing.Point(656, 8)
        Me.DoorLock.Name = "DoorLock"
        Me.DoorLock.Size = New System.Drawing.Size(80, 80)
        Me.DoorLock.TabIndex = 116
        Me.DoorLock.Text = "Lock Door"
        Me.DoorLock.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonPotReset
        '
        Me.ButtonPotReset.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonPotReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPotReset.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPotReset.ImageIndex = 1
        Me.ButtonPotReset.Location = New System.Drawing.Point(441, 8)
        Me.ButtonPotReset.Name = "ButtonPotReset"
        Me.ButtonPotReset.Size = New System.Drawing.Size(80, 80)
        Me.ButtonPotReset.TabIndex = 105
        Me.ButtonPotReset.Text = "Reset Pot"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 97
        Me.Label6.Text = "Log:"
        '
        'ButtonOpenFile
        '
        Me.ButtonOpenFile.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonOpenFile.Location = New System.Drawing.Point(672, 112)
        Me.ButtonOpenFile.Name = "ButtonOpenFile"
        Me.ButtonOpenFile.Size = New System.Drawing.Size(72, 26)
        Me.ButtonOpenFile.TabIndex = 96
        Me.ButtonOpenFile.Text = "....."
        '
        'TextBoxFilename
        '
        Me.TextBoxFilename.BackColor = System.Drawing.SystemColors.Control
        Me.TextBoxFilename.Enabled = False
        Me.TextBoxFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBoxFilename.Location = New System.Drawing.Point(56, 112)
        Me.TextBoxFilename.Name = "TextBoxFilename"
        Me.TextBoxFilename.ReadOnly = True
        Me.TextBoxFilename.Size = New System.Drawing.Size(600, 27)
        Me.TextBoxFilename.TabIndex = 95
        Me.TextBoxFilename.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 94
        Me.Label5.Text = "File:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonClean
        '
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClean.ImageIndex = 5
        Me.ButtonClean.Location = New System.Drawing.Point(186, 8)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.Size = New System.Drawing.Size(80, 80)
        Me.ButtonClean.TabIndex = 89
        Me.ButtonClean.Text = "Clean On"
        '
        'ButtonPurge
        '
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPurge.ImageIndex = 4
        Me.ButtonPurge.Location = New System.Drawing.Point(101, 8)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(80, 80)
        Me.ButtonPurge.TabIndex = 88
        Me.ButtonPurge.Text = "Purge On"
        '
        'ButtonChgSyringe
        '
        Me.ButtonChgSyringe.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonChgSyringe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonChgSyringe.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonChgSyringe.ImageIndex = 1
        Me.ButtonChgSyringe.Location = New System.Drawing.Point(271, 8)
        Me.ButtonChgSyringe.Name = "ButtonChgSyringe"
        Me.ButtonChgSyringe.Size = New System.Drawing.Size(80, 80)
        Me.ButtonChgSyringe.TabIndex = 87
        Me.ButtonChgSyringe.Text = "Change Syringe"
        '
        'ButtonHome
        '
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonHome.ImageIndex = 0
        Me.ButtonHome.Location = New System.Drawing.Point(16, 8)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(80, 80)
        Me.ButtonHome.TabIndex = 86
        Me.ButtonHome.Text = "Home"
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonCalibrate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonCalibrate.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCalibrate.ImageIndex = 5
        Me.ButtonCalibrate.Location = New System.Drawing.Point(356, 8)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(80, 80)
        Me.ButtonCalibrate.TabIndex = 88
        Me.ButtonCalibrate.Text = "Move Calibrate"
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Menu
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Black
        Me.LabelMessege.Location = New System.Drawing.Point(8, 920)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(750, 32)
        Me.LabelMessege.TabIndex = 85
        Me.LabelMessege.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
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
        'HardwareInitTimer
        '
        '
        'TowerLightImageList
        '
        Me.TowerLightImageList.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.TowerLightImageList.ImageSize = New System.Drawing.Size(64, 64)
        Me.TowerLightImageList.ImageStream = CType(resources.GetObject("TowerLightImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.TowerLightImageList.TransparentColor = System.Drawing.Color.Transparent
        '
        'ImageListButtons
        '
        Me.ImageListButtons.ImageSize = New System.Drawing.Size(40, 40)
        Me.ImageListButtons.ImageStream = CType(resources.GetObject("ImageListButtons.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListButtons.TransparentColor = System.Drawing.Color.Transparent
        '
        'FormProduction
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1276, 990)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.PanelProDownTimeInfor)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Menu = Me.MainMenuProduction
        Me.Name = "FormProduction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel2.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.PanelProDownTimeInfor.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'reset!
        KeyboardControl.GainControls()
        Init()
        While isInited = False
            Application.DoEvents()
        End While
        Console.WriteLine(DateTime.Now & "#1")
        ResetToIdle()
        'gui visibility
        Panel5.Controls.Add(m_Tri.SteppingButtons)
        m_Tri.SteppingButtons.Location = PanelToBeAdded.Location() 'New Point(84, 192)
        m_Tri.SteppingButtons.BringToFront()
        m_Tri.SteppingButtons.Show()
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
        SystemSetupDataRetrieve(SettingsMode.GlobalSettings)
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
        'vision
        'motion controller
        'm_Tri.Connect_Controller()
        SetState("Homing")
        HardwareInitTimer.Start()
        'timers start
        IDS.StartErrorCheck()
        TimerMonitor.Start()
        Programming.IOCheck.Start()
        ' background threads help to update UI without slowing them down.
        ThreadMonitor = New Threading.Thread(AddressOf StateMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.Normal
        'ThreadMonitor.IsBackground = True
        ThreadMonitor.Start()

        ThreadExecutor = New Threading.Thread(AddressOf StateChanger)
        ThreadExecutor.Priority = Threading.ThreadPriority.Normal
        ThreadExecutor.Start()

        MouseTimer = New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 200)

    End Sub

    Private Sub MouseJogging(ByVal state As Object)

        If Production.ButtonCalibrate.Text = "Set Calibrate" Then
        Else
            If IsBusy() And Not IsJogging() Then Exit Sub
            If m_Tri.MachineHoming Or m_Tri.MachineRunning Or m_Tri.Stepping Then Exit Sub
        End If

        'm_keyBoard.Poll()
        m_TrackBall.Poll()
        'isPress = m_keyBoard.State.Item(Key.LeftControl)
        isPress = KeyboardControl.ControlKeyPressed

        Dim x As Integer
        Dim y As Integer
        Dim VrData(3) As Single

        If isPress Then

            SetState("Jogging")

            VrData(0) = 0
            VrData(1) = 0.0
            VrData(2) = 0.0

            'Dim isPressAlt As Boolean = m_keyBoard.State.Item(Key.LeftAlt)
            'If isPressAlt Then
            '    Exit Sub
            'End If
            x = m_TrackBall.MouseX()
            y = m_TrackBall.MouseY()

            Dim ratio As Double

            If Math.Abs(x) >= Math.Abs(y) Then
                If x > deadzone Then
                    jogspeed = CInt(Math.Abs(x) / maxMouseRangeX * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(y) / x
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X+ Y-
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = -jogspeed
                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = 0

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                ElseIf x < -deadzone Then
                    jogspeed = CInt(Math.Abs(x) / maxMouseRangeX * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(y) / x
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X- Y+
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = 0

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_Tri.SetTrioMotionValues("Jogging", VrData)
                    isJogON = True 'False
                End If
            Else
                If y < -deadzone Then
                    jogspeed = CInt(Math.Abs(y) / maxMouseRangeY * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If

                    ratio = CDbl(x) / y
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X- Y+
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = jogspeed

                        SetState("Jogging")
                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y+
                        VrData(0) = 1
                        VrData(1) = 0
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If

                ElseIf y > deadzone Then
                    jogspeed = CInt(Math.Abs(y) / maxMouseRangeY * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(x) / y
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X+ Y-
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y-
                        VrData(0) = 1
                        VrData(1) = 0
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_Tri.SetTrioMotionValues("Jogging", VrData)
                    isJogON = True 'False
                End If
            End If
        Else
            If isJogON Then
                VrData(0) = 2
                VrData(1) = 0.0
                VrData(2) = 0.0

                m_Tri.SetTrioMotionValues("Jogging", VrData)
                isJogON = False
                If m_EditStateFlag Then
                    'reset to idle without the camera thing
                    SetState("Idle")
                    m_Tri.SetMachineStop()
                    SetLampsToReadyMode()
                    UnlockMovementButtons()
                    ChangeButtonState("Idle")
                Else
                    If Production.ButtonCalibrate.Text = "Set Calibrate" Then
                        Production.ButtonCalibrate.Enabled = True
                        m_Tri.SteppingButtons.Enabled = True
                    Else
                        ResetToIdle()
                    End If
                End If
            End If

            countMouseTimer += 1
            If (countMouseTimer >= 5) Then
                TraceGCCollect()
                countMouseTimer = 0
            End If
        End If

    End Sub

    Private Sub FormProduction_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'If ContinuousMode.Checked = True Then
        '    ContinuousMode.Checked = False
        'End If
        'isPress = False
        ''error handling
        'Form_Service.ResetEventCode()
        ''vision
        ''timers stop
        'IDS.StopErrorCheck()
        'TimerMonitor.Stop()
        'Programming.IOCheck.Stop()
        'ThreadMonitor.Abort()
        'ThreadExecutor.Abort()
        'MouseTimer.Dispose()
        ''motion controller
        'm_Tri.TurnOff("Material Air")
        'm_Tri.Disconnect_Controller()
        ''hardware
        'OffLaser()
        'OffTowerLamp()
        'UnlockDoor()
        ''Close()
        'IDS.FrmExecution.Hide()
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
        'IDS.Data.ParameterID.RecordID = ""
        'IDSData.Admin.Folder.FileExtension = "Pat"
        'IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        If Nothing = TextBoxFilename.Text Then
            ButtonOpenFile.Enabled = True
            Return
        End If

        LabelMessage("Please wait, system is uploading..")
        Dim filename As String = m_Execution.m_File.FolderWithNameFromFileName(TextBoxFilename.Text)
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()
        Programming.Disp_Dispenser_Unit_info()

        TextBoxFilename.Refresh()
        RichTextBoxNote.Refresh()

        ProductionInfoDisp()

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

    Private Sub DoorLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoorLock.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't unlock the door when machine running!")
            DoorLock.Text = "Unlock Door"
            DoorLock.ImageIndex = 2
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

    Public Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

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
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned off.")
            End If
        ElseIf ContinuousMode.Checked = False Then
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned on.")
            End If
        End If

        ContinuousMode.Checked = Not ContinuousMode.Checked

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

        'Dim onoff As Integer = 1
        'Dim conveyorErrorID, ErrorID As Integer

        'ChangeButtonState("Only Stop Displayed")

        'While onoff = 1 And Not m_Tri.MotionError

        '    m_Tri.m_TriCtrl.In(gIOloadReady, gIOloadReady, onoff) 'when ready set onoff = 0
        '    LabelMessage("Waiting for board..")
        '    TraceDoEvents()

        '    If IsIdle() Or m_Tri.MotionError Then
        '        Form_Service.LogEventInSPCReport("Motion controller error")
        '        ResetToIdle()
        '        Exit Sub
        '    End If

        '    '3202,3203,3204 are errors from lifter
        '    '3205,3206,3208 are conveyor retrieve or stage 1/3 jamming errors.
        '    conveyorErrorID = CInt(IDS.__ID) Mod 10000
        '    If conveyorErrorID = 3202 Or conveyorErrorID = 3203 Or conveyorErrorID = 3204 Or conveyorErrorID = 3205 Or conveyorErrorID = 3206 Or conveyorErrorID = 3208 Then
        '        LabelMessage("No board is coming or conveyor jammed.")
        '        ResetToIdle()
        '        Exit Sub
        '    End If

        'End While

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
        If (door_close = False And door_interlock = True) And IsIdle() Then
            ResetTimer("Reset Autopurging Timers")
            ResetTimer("Reset Potlife Start Timer")
            Exit Sub
        End If

        Dim auto_purging_enabled As Boolean = IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption And IsIdle() And Form_Service.Visible = False
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

    Private Sub ButtonCalibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click
        If SetState("Needle Calibration") Then DoCalibrate()
    End Sub

    Private Sub RichTextBoxNote_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBoxNote.TextChanged

    End Sub

    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExit.Click
        If ContinuousMode.Checked = True Then
            ContinuousMode.Checked = False
        End If
        isPress = False
        KeyboardControl.ReleaseControls()
        'error handling
        Form_Service.ResetEventCode()
        'vision
        'timers stop
        IDS.StopErrorCheck()
        TimerMonitor.Stop()
        HardwareInitTimer.Stop()
        HardwareInitTimer.Dispose()
        Programming.IOCheck.Stop()
        ThreadMonitor.Abort()
        ThreadExecutor.Abort()
        MouseTimer.Dispose()
        'motion controller
        m_Tri.TurnOff("Material Air")
        m_Tri.Disconnect_Controller()
        'hardware
        OffLaser()
        OffTowerLamp()
        UnlockDoor()
        'Close()
        IDS.FrmExecution.Hide()
        Close()
    End Sub

    Private Sub HardwareInitTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles HardwareInitTimer.Tick
        HardwareInitTimer.Stop()
        HardwareInitTimer.Enabled = False
        m_Tri.Connect_Controller()
        SetState("Homing")
        Console.WriteLine("Hardware Init Timer Called")
    End Sub

    Private Sub btStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStart.Click
        SetState("Start")
    End Sub

    Private Sub btPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPause.Click
        PauseDispensing()
    End Sub

    Private Sub btStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStop.Click
        StopDispensing()
    End Sub
End Class
