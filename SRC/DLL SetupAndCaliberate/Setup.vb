Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
Imports System.Threading

Public Class Setup
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents PanelRight As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBoxRobotZ As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotY As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotX As System.Windows.Forms.TextBox
    Friend WithEvents ButtonSystemIO As System.Windows.Forms.Button
    Friend WithEvents GpB_Configurations As System.Windows.Forms.GroupBox
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents ButtonGantrySetup As System.Windows.Forms.Button
    Friend WithEvents ButtonCameraSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonLaserSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenserSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonStationPositions As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonSPCLogging As System.Windows.Forms.Button
    Friend WithEvents ButtonThermalSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonEventHandling As System.Windows.Forms.Button
    Friend WithEvents ButtonHardwareCommunicationSetup As System.Windows.Forms.Button
    Public WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents LabelMessage As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents btLaserInterface As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents HardwareTimer As System.Timers.Timer
    Friend WithEvents HomingCheckTimer As System.Windows.Forms.Timer
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents rbOneHead As System.Windows.Forms.RadioButton
    Friend WithEvents rbTwoHead As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents tbEquipmentID As System.Windows.Forms.TextBox
    Friend WithEvents btExit As System.Windows.Forms.Button
    Friend WithEvents gbSystemSetup As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Setup))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GpB_Configurations = New System.Windows.Forms.GroupBox
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.tbEquipmentID = New System.Windows.Forms.TextBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.rbTwoHead = New System.Windows.Forms.RadioButton
        Me.rbOneHead = New System.Windows.Forms.RadioButton
        Me.ButtonHardwareCommunicationSetup = New System.Windows.Forms.Button
        Me.PanelRight = New System.Windows.Forms.Panel
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.ButtonGantrySetup = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSetup = New System.Windows.Forms.Button
        Me.ButtonSystemIO = New System.Windows.Forms.Button
        Me.ButtonLaserSetup = New System.Windows.Forms.Button
        Me.ButtonStationPositions = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSettings = New System.Windows.Forms.Button
        Me.ButtonEventHandling = New System.Windows.Forms.Button
        Me.ButtonThermalSettings = New System.Windows.Forms.Button
        Me.ButtonSPCLogging = New System.Windows.Forms.Button
        Me.ButtonDispenserSettings = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotZ = New System.Windows.Forms.TextBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotY = New System.Windows.Forms.TextBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.TextBoxRobotX = New System.Windows.Forms.TextBox
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.HardwareTimer = New System.Timers.Timer
        Me.LabelMessage = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.gbSystemSetup = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.btLaserInterface = New System.Windows.Forms.Button
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.HomingCheckTimer = New System.Windows.Forms.Timer(Me.components)
        Me.btExit = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GpB_Configurations.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.HardwareTimer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.gbSystemSetup.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.GpB_Configurations)
        Me.Panel1.Controls.Add(Me.ButtonHardwareCommunicationSetup)
        Me.Panel1.Location = New System.Drawing.Point(0, 528)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(744, 352)
        Me.Panel1.TabIndex = 5
        '
        'GpB_Configurations
        '
        Me.GpB_Configurations.Controls.Add(Me.GroupBox8)
        Me.GpB_Configurations.Controls.Add(Me.GroupBox7)
        Me.GpB_Configurations.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GpB_Configurations.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GpB_Configurations.Location = New System.Drawing.Point(28, 8)
        Me.GpB_Configurations.Name = "GpB_Configurations"
        Me.GpB_Configurations.Size = New System.Drawing.Size(308, 296)
        Me.GpB_Configurations.TabIndex = 62
        Me.GpB_Configurations.TabStop = False
        Me.GpB_Configurations.Text = "Configurations"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.tbEquipmentID)
        Me.GroupBox8.Location = New System.Drawing.Point(16, 104)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(280, 72)
        Me.GroupBox8.TabIndex = 64
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Equipment ID"
        '
        'tbEquipmentID
        '
        Me.tbEquipmentID.Location = New System.Drawing.Point(16, 32)
        Me.tbEquipmentID.Name = "tbEquipmentID"
        Me.tbEquipmentID.Size = New System.Drawing.Size(248, 27)
        Me.tbEquipmentID.TabIndex = 0
        Me.tbEquipmentID.Text = ""
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.rbTwoHead)
        Me.GroupBox7.Controls.Add(Me.rbOneHead)
        Me.GroupBox7.Location = New System.Drawing.Point(16, 24)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(280, 64)
        Me.GroupBox7.TabIndex = 63
        Me.GroupBox7.TabStop = False
        '
        'rbTwoHead
        '
        Me.rbTwoHead.Location = New System.Drawing.Point(144, 24)
        Me.rbTwoHead.Name = "rbTwoHead"
        Me.rbTwoHead.Size = New System.Drawing.Size(128, 24)
        Me.rbTwoHead.TabIndex = 67
        Me.rbTwoHead.Text = "Two Heads"
        Me.rbTwoHead.Visible = False
        '
        'rbOneHead
        '
        Me.rbOneHead.Location = New System.Drawing.Point(8, 24)
        Me.rbOneHead.Name = "rbOneHead"
        Me.rbOneHead.Size = New System.Drawing.Size(128, 24)
        Me.rbOneHead.TabIndex = 66
        Me.rbOneHead.Text = "Single Head"
        '
        'ButtonHardwareCommunicationSetup
        '
        Me.ButtonHardwareCommunicationSetup.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonHardwareCommunicationSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonHardwareCommunicationSetup.Location = New System.Drawing.Point(504, 280)
        Me.ButtonHardwareCommunicationSetup.Name = "ButtonHardwareCommunicationSetup"
        Me.ButtonHardwareCommunicationSetup.Size = New System.Drawing.Size(232, 32)
        Me.ButtonHardwareCommunicationSetup.TabIndex = 46
        Me.ButtonHardwareCommunicationSetup.Text = "Hardware Communication"
        Me.ButtonHardwareCommunicationSetup.Visible = False
        '
        'PanelRight
        '
        Me.PanelRight.Location = New System.Drawing.Point(8, 16)
        Me.PanelRight.Name = "PanelRight"
        Me.PanelRight.Size = New System.Drawing.Size(504, 864)
        Me.PanelRight.TabIndex = 6
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.PanelRight)
        Me.GroupBox4.Location = New System.Drawing.Point(744, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(520, 888)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        '
        'ButtonGantrySetup
        '
        Me.ButtonGantrySetup.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.ButtonGantrySetup.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonGantrySetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonGantrySetup.Location = New System.Drawing.Point(96, 32)
        Me.ButtonGantrySetup.Name = "ButtonGantrySetup"
        Me.ButtonGantrySetup.Size = New System.Drawing.Size(224, 56)
        Me.ButtonGantrySetup.TabIndex = 43
        Me.ButtonGantrySetup.Text = "Gantry Parameters"
        '
        'ButtonNeedleCalibSetup
        '
        Me.ButtonNeedleCalibSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSetup.Location = New System.Drawing.Point(96, 104)
        Me.ButtonNeedleCalibSetup.Name = "ButtonNeedleCalibSetup"
        Me.ButtonNeedleCalibSetup.Size = New System.Drawing.Size(224, 56)
        Me.ButtonNeedleCalibSetup.TabIndex = 44
        Me.ButtonNeedleCalibSetup.Text = "Needle Setup"
        '
        'ButtonSystemIO
        '
        Me.ButtonSystemIO.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSystemIO.Location = New System.Drawing.Point(360, 32)
        Me.ButtonSystemIO.Name = "ButtonSystemIO"
        Me.ButtonSystemIO.Size = New System.Drawing.Size(224, 56)
        Me.ButtonSystemIO.TabIndex = 37
        Me.ButtonSystemIO.Text = "System IO"
        '
        'ButtonLaserSetup
        '
        Me.ButtonLaserSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonLaserSetup.Location = New System.Drawing.Point(360, 104)
        Me.ButtonLaserSetup.Name = "ButtonLaserSetup"
        Me.ButtonLaserSetup.Size = New System.Drawing.Size(224, 56)
        Me.ButtonLaserSetup.TabIndex = 45
        Me.ButtonLaserSetup.Text = "Laser Setup"
        Me.ButtonLaserSetup.Visible = False
        '
        'ButtonStationPositions
        '
        Me.ButtonStationPositions.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStationPositions.Location = New System.Drawing.Point(360, 56)
        Me.ButtonStationPositions.Name = "ButtonStationPositions"
        Me.ButtonStationPositions.Size = New System.Drawing.Size(224, 56)
        Me.ButtonStationPositions.TabIndex = 43
        Me.ButtonStationPositions.Text = "Gantry Settings"
        '
        'ButtonNeedleCalibSettings
        '
        Me.ButtonNeedleCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSettings.Location = New System.Drawing.Point(96, 120)
        Me.ButtonNeedleCalibSettings.Name = "ButtonNeedleCalibSettings"
        Me.ButtonNeedleCalibSettings.Size = New System.Drawing.Size(224, 56)
        Me.ButtonNeedleCalibSettings.TabIndex = 44
        Me.ButtonNeedleCalibSettings.Text = "Needle Calibration Settings"
        '
        'ButtonEventHandling
        '
        Me.ButtonEventHandling.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEventHandling.Location = New System.Drawing.Point(96, 184)
        Me.ButtonEventHandling.Name = "ButtonEventHandling"
        Me.ButtonEventHandling.Size = New System.Drawing.Size(224, 56)
        Me.ButtonEventHandling.TabIndex = 47
        Me.ButtonEventHandling.Text = "Event Handling Settings"
        '
        'ButtonThermalSettings
        '
        Me.ButtonThermalSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonThermalSettings.Location = New System.Drawing.Point(360, 184)
        Me.ButtonThermalSettings.Name = "ButtonThermalSettings"
        Me.ButtonThermalSettings.Size = New System.Drawing.Size(224, 56)
        Me.ButtonThermalSettings.TabIndex = 46
        Me.ButtonThermalSettings.Text = "Heater Settings"
        Me.ButtonThermalSettings.Visible = False
        '
        'ButtonSPCLogging
        '
        Me.ButtonSPCLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCLogging.Location = New System.Drawing.Point(360, 120)
        Me.ButtonSPCLogging.Name = "ButtonSPCLogging"
        Me.ButtonSPCLogging.Size = New System.Drawing.Size(224, 56)
        Me.ButtonSPCLogging.TabIndex = 51
        Me.ButtonSPCLogging.Text = "SPC Logging Settings"
        '
        'ButtonDispenserSettings
        '
        Me.ButtonDispenserSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenserSettings.Location = New System.Drawing.Point(96, 56)
        Me.ButtonDispenserSettings.Name = "ButtonDispenserSettings"
        Me.ButtonDispenserSettings.Size = New System.Drawing.Size(224, 56)
        Me.ButtonDispenserSettings.TabIndex = 42
        Me.ButtonDispenserSettings.Text = "Dispenser Settings"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.CheckBox3)
        Me.Panel2.Controls.Add(Me.TextBoxRobotZ)
        Me.Panel2.Controls.Add(Me.CheckBox4)
        Me.Panel2.Controls.Add(Me.TextBoxRobotY)
        Me.Panel2.Controls.Add(Me.CheckBox5)
        Me.Panel2.Controls.Add(Me.TextBoxRobotX)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel2.Location = New System.Drawing.Point(872, 952)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(392, 28)
        Me.Panel2.TabIndex = 5
        '
        'CheckBox3
        '
        Me.CheckBox3.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox3.Image = CType(resources.GetObject("CheckBox3.Image"), System.Drawing.Image)
        Me.CheckBox3.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.CheckBox3.Location = New System.Drawing.Point(248, 8)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox3.TabIndex = 75
        '
        'TextBoxRobotZ
        '
        Me.TextBoxRobotZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotZ.Location = New System.Drawing.Point(288, 4)
        Me.TextBoxRobotZ.Name = "TextBoxRobotZ"
        Me.TextBoxRobotZ.ReadOnly = True
        Me.TextBoxRobotZ.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotZ.TabIndex = 74
        Me.TextBoxRobotZ.Text = "Z 100.000"
        '
        'CheckBox4
        '
        Me.CheckBox4.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox4.Image = CType(resources.GetObject("CheckBox4.Image"), System.Drawing.Image)
        Me.CheckBox4.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.CheckBox4.Location = New System.Drawing.Point(128, 8)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox4.TabIndex = 73
        '
        'TextBoxRobotY
        '
        Me.TextBoxRobotY.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotY.Location = New System.Drawing.Point(168, 4)
        Me.TextBoxRobotY.Name = "TextBoxRobotY"
        Me.TextBoxRobotY.ReadOnly = True
        Me.TextBoxRobotY.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotY.TabIndex = 72
        Me.TextBoxRobotY.Text = "Y: 100.000"
        '
        'CheckBox5
        '
        Me.CheckBox5.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox5.Image = CType(resources.GetObject("CheckBox5.Image"), System.Drawing.Image)
        Me.CheckBox5.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.CheckBox5.Location = New System.Drawing.Point(8, 8)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox5.TabIndex = 71
        '
        'TextBoxRobotX
        '
        Me.TextBoxRobotX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotX.Location = New System.Drawing.Point(48, 4)
        Me.TextBoxRobotX.Name = "TextBoxRobotX"
        Me.TextBoxRobotX.ReadOnly = True
        Me.TextBoxRobotX.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotX.TabIndex = 6
        Me.TextBoxRobotX.Text = "X: 100.000"
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""C:\IDS\BIN\EEtver1.mdb"";Mode=Share Deny None;Jet OLED" & _
        "B:Engine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Je" & _
        "t OLEDB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Com" & _
        "pact Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Cre" & _
        "ate System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=A" & _
        "dmin;Jet OLEDB:Global Bulk Transactions=1"
        '
        'HardwareTimer
        '
        Me.HardwareTimer.Interval = 500
        Me.HardwareTimer.SynchronizingObject = Me
        '
        'LabelMessage
        '
        Me.LabelMessage.BackColor = System.Drawing.Color.White
        Me.LabelMessage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessage.ForeColor = System.Drawing.Color.Black
        Me.LabelMessage.Location = New System.Drawing.Point(8, 952)
        Me.LabelMessage.Name = "LabelMessage"
        Me.LabelMessage.Size = New System.Drawing.Size(864, 32)
        Me.LabelMessage.TabIndex = 86
        Me.LabelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ButtonThermalSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonSPCLogging)
        Me.GroupBox1.Controls.Add(Me.ButtonDispenserSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonStationPositions)
        Me.GroupBox1.Controls.Add(Me.ButtonNeedleCalibSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonEventHandling)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 192)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(712, 296)
        Me.GroupBox1.TabIndex = 87
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Global Settings :"
        '
        'gbSystemSetup
        '
        Me.gbSystemSetup.Controls.Add(Me.ButtonGantrySetup)
        Me.gbSystemSetup.Controls.Add(Me.ButtonNeedleCalibSetup)
        Me.gbSystemSetup.Controls.Add(Me.ButtonSystemIO)
        Me.gbSystemSetup.Controls.Add(Me.ButtonLaserSetup)
        Me.gbSystemSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSystemSetup.Location = New System.Drawing.Point(16, 8)
        Me.gbSystemSetup.Name = "gbSystemSetup"
        Me.gbSystemSetup.Size = New System.Drawing.Size(712, 184)
        Me.gbSystemSetup.TabIndex = 88
        Me.gbSystemSetup.TabStop = False
        Me.gbSystemSetup.Text = "System Setup :"
        '
        'GroupBox5
        '
        Me.GroupBox5.Location = New System.Drawing.Point(0, 944)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(1272, 8)
        Me.GroupBox5.TabIndex = 90
        Me.GroupBox5.TabStop = False
        '
        'btLaserInterface
        '
        Me.btLaserInterface.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btLaserInterface.Location = New System.Drawing.Point(96, 24)
        Me.btLaserInterface.Name = "btLaserInterface"
        Me.btLaserInterface.Size = New System.Drawing.Size(224, 56)
        Me.btLaserInterface.TabIndex = 63
        Me.btLaserInterface.Text = "Laser"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.btLaserInterface)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(16, 488)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(712, 88)
        Me.GroupBox6.TabIndex = 64
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Hardware :"
        Me.GroupBox6.Visible = False
        '
        'HomingCheckTimer
        '
        Me.HomingCheckTimer.Interval = 30000
        '
        'btExit
        '
        Me.btExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btExit.Location = New System.Drawing.Point(1192, 896)
        Me.btExit.Name = "btExit"
        Me.btExit.Size = New System.Drawing.Size(72, 48)
        Me.btExit.TabIndex = 52
        Me.btExit.Text = "Exit"
        '
        'Setup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1272, 990)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.gbSystemSetup)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.LabelMessage)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.btExit)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "Setup"
        Me.Text = "System Setup"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.GpB_Configurations.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.HardwareTimer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.gbSystemSetup.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Shared m_RunState As Integer = 0
    Private Declare Function InterlockedExchange Lib "kernel32" (ByRef Target As Integer, ByVal Value As Integer) As Integer
    Private MouseTimer As System.Threading.Timer
    Private ThreadMonitor As Threading.Thread
    Private threadMain As Threading.Thread

    'jogging variables
    Dim m_TrackBall As New DLL_Export_Device_Motor.Mouse(Me)
    'Dim m_keyBoard As New DLL_Export_Device_Motor.Keyboard(Me)
    Private isJogON As Boolean = False
    Shared isPress As Boolean
    Dim deadzone As Integer = 3
    Dim jogspeed As Integer = 0
    Const maxSpeed As Integer = 100
    Const maxMouseRangeX = 600.0 '6000.0 '4500.0
    Const maxMouseRangeY = 300.0 '2000.0
    Const ratioLB = 0.4
    Const ratioUB = 1.05
    Dim ratio As Double
    Dim selectedButton As System.Windows.Forms.Button
    Dim homingDone As Boolean = False

    Private Sub MouseJogging(ByVal state As Object)

        If IDS.Data.SystemAtLogin = True Then Exit Sub
        If MachineRunning() Then Exit Sub

        'm_keyBoard.Poll()
        m_TrackBall.Poll()

        'isPress = m_keyBoard.State.Item(Key.LeftControl)
        isPress = KeyboardControl.ControlKeyPressed
        Dim x As Integer
        Dim y As Integer
        Dim CmdStr As String
        Dim VrData(3) As Single
        If isPress Then
            m_Tri.SteppingButtons.Enabled = False
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
                    VrData(1) = 0
                    VrData(2) = 0
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
                    VrData(1) = 0
                    VrData(2) = 0
                    m_Tri.SetTrioMotionValues("Jogging", VrData)
                    isJogON = True 'False
                End If
            End If
        Else
            If isJogON Then
                VrData(0) = 2
                VrData(1) = 0
                VrData(2) = 0
                m_Tri.SetTrioMotionValues("Jogging", VrData)
                isJogON = False
                m_Tri.SteppingButtons.Enabled = True
            End If

        End If

    End Sub

    Private Sub TheThreadMonitor()
        Do
            TraceDoEvents()
            MySleep(50)
            'we may want more detailed functionality here
            m_Tri.GetIDSState()
            TextBoxRobotX.Text = "X: " & m_Tri.XPosition
            TextBoxRobotY.Text = "Y: " & m_Tri.YPosition
            TextBoxRobotZ.Text = "Z: " & m_Tri.ZPosition
            If m_Tri.HomingFinished() Then
                GC.Collect()
                m_Tri.SteppingButtons.Enabled = True
                For Each ctrl As Control In Me.Controls
                    ctrl.Enabled = True
                Next
                homingDone = True
                HomingCheckTimer.Stop()
                HomingCheckTimer.Enabled = False
                ShowLabelMessage("Homing done. System is ready!")
            End If
        Loop
    End Sub

    Private Sub Setup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'display the stepping panel
        m_Tri2 = m_Tri
        Panel1.Controls.Add(m_Tri.SteppingButtons)
        m_Tri.SteppingButtons.BringToFront()
        'change GUI settings unique for system setup here
        m_Tri.SteppingButtons.Location = New System.Drawing.Point(350, GpB_Configurations.Location.Y)

        'gui
        Location = New Point(0, 0)
        MyBase.StartPosition = FormStartPosition.Manual

        Conn = OleDbConnection1
        Call DataLoad()
        FillFromSystemConfigureTable() 'lsgoh

        'kr
        If IDS.Data.Hardware.Dispenser.CurrentHeads = 1 Then
            rbOneHead.Checked = True
        Else
            rbTwoHead.Checked = True
        End If
        tbEquipmentID.Text = IDSData.Machine.EquipmentID
        'motion controller start and do homing
        m_Tri.SteppingButtons.Enabled = False
        m_Tri.SteppingButtons.Show()
        ShowLabelMessage("Initializing hardware......")
        For Each ctrl As Control In Me.Controls
            If ctrl.Name <> btExit.Name Then
                ctrl.Enabled = False
            End If
        Next
        AddPanel(PanelRight, MyGantrySetup.PanelToBeAdded)
        MyGantrySetup.RevertData()
        If Not (selectedButton Is Nothing) Then
            selectedButton.FlatStyle = FlatStyle.Standard
        End If
        ButtonGantrySetup.FlatStyle = FlatStyle.Flat
        selectedButton = ButtonGantrySetup
        MyGantrySetup.StationPosition.SelectedIndex = 0
        ButtonGantrySetup.Select()
        'timers
        HardwareTimer.Start()
        IDS.StartErrorCheck()
        ThreadMonitor = New Threading.Thread(AddressOf TheThreadMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.BelowNormal
        ThreadMonitor.Start()
        MouseTimer = New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 180)

        'hardware
        Laser.OpenPort()
        OnLaser()
    End Sub

    Private Sub Setup_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        isPress = False
        'gui clear
        UpdatetoSystemConfigureTable()
        'PanelVision.Controls.Clear()
        MySettings.Controls.Clear()
        PanelRight.Controls.Remove(CurrentControl)
        MyGantrySettings.Controls.Clear()

        'motion controller stop
        m_Tri.Disconnect_Controller()

        'disable stepping buttons!
        m_Tri.SteppingButtons.Enabled = False

        'timers stop
        IDS.StopErrorCheck()
        ThreadMonitor.Abort()
        MouseTimer.Dispose()
        HardwareTimer.Stop()

        'hardware
        OffLaser()
        'Laser.ClosePort()
        'Conveyor.ClosePort()
        'Weighting_Scale.ClosePort()

    End Sub

    Private Function Get_Limit(ByVal Direction As String) As Boolean

        Dim Read_Value As Integer = 0

        m_Tri.m_TriCtrl.In(0, 5, Read_Value)

        Select Case Direction
            Case "X+"
                Read_Value = Read_Value And &H2
            Case "X-"
                Read_Value = Read_Value And &H1
            Case "Y+"
                Read_Value = Read_Value And &H4
            Case "Y-"
                Read_Value = Read_Value And &H8
            Case "Up"
                Read_Value = Read_Value And &H10
            Case "Dn"
                Read_Value = Read_Value And &H20
            Case Else
                Return False
        End Select

        If (Read_Value = 0) Then
            MsgBox("Machine gets the limit position!", MsgBoxStyle.OKOnly, "Limit Error")
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub Setup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim mode As Integer = 0
        KeyboardControl.ReleaseControls()
        Laser.Instance.Hide()
        If m_Tri.IsHomingDone Then
            m_Tri.Move_Z(0)
            position(0) = IDS.Data.Hardware.Gantry.ParkPosition.X
            position(1) = IDS.Data.Hardware.Gantry.ParkPosition.Y
            m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
            If Not m_Tri.Move_XY(position) Then Exit Sub
        End If

        m_Tri.m_TriCtrl.GetTable(21, 1, mode)
        If mode <> 0 Then
            m_Tri.m_TriCtrl.SetTable(21, 1, 0)

            m_Tri.m_TriCtrl.Stop("Dispense")
            m_Tri.m_TriCtrl.Stop("SetDatum")

            m_Tri.m_TriCtrl.RapidStop()

            m_Tri.m_TriCtrl.Cancel(1, 0)
            m_Tri.m_TriCtrl.Cancel(1, 1)
            m_Tri.m_TriCtrl.Cancel(1, 2)
            m_Tri.m_TriCtrl.Cancel(0, 0)
            m_Tri.m_TriCtrl.Cancel(0, 1)
            m_Tri.m_TriCtrl.Cancel(0, 2)

            m_Tri.m_TriCtrl.Op(0)
            m_Tri.m_TriCtrl.Execute("STOP SETDATUM")
            m_Tri.RunTrioMotionProgram("EXITIDS")
        End If
    End Sub

    Private Sub SetPressure(ByVal PressureInput As Decimal, ByVal SuckBack As Decimal, ByVal Mode As String)

        'If Mode = "Time Pressure" Then
        '    If IDS.Data.Hardware.Dispenser.Left.SmartPressureUnit = False Then
        '        IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, PressureInput, -(SuckBack + 1))
        '    Else
        '        IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(1, PressureInput, -(SuckBack + 1))
        '    End If
        'ElseIf Mode = "Jetter" Then
        '    If IDS.Data.Hardware.Dispenser.Left.SmartPressureUnit = False Then
        '        IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, PressureInput, -(0 + 1))
        '    Else
        '        IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(1, PressureInput, -(0 + 1))
        '    End If
        'End If

    End Sub

    Friend Sub FillFromSystemConfigureTable()

        Dim TableName As String = "SystemConfigureTable"
        DBView = New DataView(DS.Tables(TableName))
        DBView.RowFilter = ""

        If DBView.Count > 0 Then
            Dim I As Integer = 0
            For I = 0 To DBView.Count - 1
                Dim Row As DataRow = DBView(I).Row

                'If Row("Hardware") = "HeightSensor" Then
                '    'CheckBoxHeightSensor.Checked = CBoolean(Row("Selected"))

                'ElseIf Row("Hardware") = "VolumeCalibration" Then
                '    CheckBoxVolume.Checked = CBoolean(Row("Selected"))

                'ElseIf Row("Hardware") = "ThermalController" Then
                '    CheckBoxHeater.Checked = CBoolean(Row("Selected"))

                'ElseIf Row("Hardware") = "Lifter&Vacuum" Then
                '    CheckBoxLifter.Checked = CBoolean(Row("Selected"))
                'End If
            Next
        End If
    End Sub
    'lsgoh
    Friend Sub UpdatetoSystemConfigureTable()

        Dim TableName As String = "SystemConfigureTable"
        DBView = New DataView(DS.Tables(TableName))
        DBView.RowFilter = ""

        If DBView.Count > 0 Then
            Dim I As Integer = 0
            For I = 0 To DBView.Count - 1
                Dim Row As DataRow = DBView(I).Row

                'If Row("Hardware") = "HeightSensor" Then
                '    'Row("Selected") = CheckBoxHeightSensor.Checked

                'ElseIf Row("Hardware") = "VolumeCalibration" Then
                '    Row("Selected") = CheckBoxVolume.Checked

                'ElseIf Row("Hardware") = "ThermalController" Then
                '    Row("Selected") = CheckBoxHeater.Checked

                'ElseIf Row("Hardware") = "Lifter&Vacuum" Then
                '    Row("Selected") = CheckBoxLifter.Checked

                'End If
            Next
        End If
        UpdateData("SELECT * FROM " + TableName, TableName)

    End Sub

    Private Sub CheckBoxVolume_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Present flag
        'If CheckBoxVolume.Checked = True Then
        '    ButtonVolumeCalibSettings.Visible = True
        'Else
        '    ButtonVolumeCalibSettings.Visible = False
        'End If ''yy
        MyHardwareCommunicationSetup.RefreshDisplay()
        MyHardwareCommunicationSetup.UpdateStatus()
    End Sub

    Private Sub CheckBoxHeater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Present flag
        'If CheckBoxHeater.Checked = True Then
        '    ButtonThermalSettings.Visible = True
        'Else
        '    ButtonThermalSettings.Visible = False
        'End If
        'MyHardwareCommunicationSetup.RefreshDisplay()
        'MyHardwareCommunicationSetup.UpdateStatus()
        'IDS.Data.Hardware.Thermal.HeaterFeatureEnabled = ButtonThermalSettings.Visible
        'IDS.Data.SaveLocalData()
    End Sub

    Private Sub ButtonSystemIO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSystemIO.Click
        'AddPanel(PanelRight, MySysIO.PanelToBeAdded)
        'PanelToBeAdded.Width = 1280
        'PanelRight.Width = 1280
        'PanelRight.Location = New Point(0, 0)
        'PanelRight.BringToFront()
        MySysIO.SetIOTable()
        IDS.Data.Admin.User.RunApplication = "System"
        MySysIO.IOReadClose = False
        MySysIO.IOReadStop = False
        MySysIO.TimerIO.Enabled = True
        MySysIO.TimerIO.Start()
        MySysIO.EditIO()
        MoveZToZero()
        MySysIO.ShowDialog()
    End Sub

    Private Sub ButtonGantry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGantrySetup.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyGantrySetup.PanelToBeAdded)
        If rbOneHead.Checked Then
            MyGantrySetup.RightHead.Visible = False
        Else
            MyGantrySetup.RightHead.Visible = True
        End If
        MyGantrySetup.StationPosition.SelectedIndex = 0
        MyGantrySetup.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonLaser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLaserSetup.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyLaser.PanelToBeAdded)
        MyLaser.LaserOffsetZ.Text = IDS.Data.Hardware.HeightSensor.Laser.HeightReference
        MyLaser.LaserOffsetX.Text = IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X
        MyLaser.LaserOffsetY.Text = IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y
        MyLaser.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonNeedle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCalibSetup.Click
        SetSelectButton(sender)
        If rbOneHead.Checked Then
            MyNeedleCalibrationSetup1.rbRightHead.Visible = False
            MyNeedleCalibrationSetup1.rbLeftHead.Checked = True
        Else
            MyNeedleCalibrationSetup1.rbRightHead.Visible = True
        End If
        MyNeedleCalibrationSetup1.Revert()
        AddPanel(PanelRight, MyNeedleCalibrationSetup1.PanelToBeAdded)
        MoveZToZero()
    End Sub

#Region " Data Reference "
    ' this function need to be removed after it is compiled as DLL .
    Private Sub SystemConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        KeyboardControl.GainControls()
        Conn = OleDbConnection1
        Call DataLoad()

    End Sub
#End Region

    Private Sub ButtonHardwareCommunicationSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHardwareCommunicationSetup.Click
        AddPanel(PanelRight, MyHardwareCommunicationSetup.PanelToBeAdded)
    End Sub

    Private Sub ButtonNeedleCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCalibSettings.Click
        SetSelectButton(sender)
        If rbOneHead.Checked Then
            MyNeedleCalibrationSettings.rbRightHead.Visible = False
            MyNeedleCalibrationSettings.rbLeftHead.Checked = True
        End If
        MyNeedleCalibrationSettings.Revert()
        AddPanel(PanelRight, MyNeedleCalibrationSettings.PanelToBeAdded)
        MoveZToZero()
    End Sub

    Private Sub ButtonStationPositions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStationPositions.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyGantrySettings.PanelToBeAdded)
        'initialize
        MyGantrySettings.RevertData()
        MyGantrySettings.StationPosition.SelectedIndex = 0
        MyGantrySettings.LeftHead.Checked = True
        MoveZToZero()
    End Sub

    Private Sub ButtonDispenserSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDispenserSettings.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyDispenserSettings.PanelToBeAdded)
        MyDispenserSettings.HeadType.SelectedIndex = 0
        MyDispenserSettings.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonSPCLogging_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPCLogging.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MySPCLogging.PanelToBeAdded)
        MySPCLogging.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonEventHandling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEventHandling.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyEventSettings.PanelToBeAdded)
        MyEventSettings.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonConveyorSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'AddPanel(PanelRight, MyConveyorSettings.PanelToBeAdded)
        'Conveyor.OpenPort()
        'Conveyor.PositionTimer.Start()
        'MyConveyorSettings.PositionTimer.Start()
        'MySettings.RevertData()
        'yy remove converyor
    End Sub

    Private Sub ButtonThermalSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonThermalSettings.Click
        SetSelectButton(sender)
        AddPanel(PanelRight, MyHeaterSettings.PanelToBeAdded)
        MyHeaterSettings.RevertData()
        MoveZToZero()
    End Sub

    Private Sub ButtonVolumeCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'AddPanel(PanelRight, MyVolumeCalibrationSettings.PanelToBeAdded)
        'MyVolumeCalibrationSettings.RevertData()
    End Sub


    Private Sub CheckBoxLifter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyHardwareCommunicationSetup.RefreshDisplay()
    End Sub

    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles HardwareTimer.Elapsed
        HardwareTimer.Stop()
        HardwareTimer.Enabled = False
        If m_Tri.Connect_Controller() Then
            m_Tri.SetMachineRun()
            ShowLabelMessage("Homing")
            m_Tri.m_TriCtrl.Execute("RUN SETDATUM")
        Else
            ShowLabelMessage("Cannot connect to motion controller. Please check network connectivity or reset controller.", True)
            For Each ctrl As Control In Me.Controls
                ctrl.Enabled = True
            Next
        End If
        
    End Sub

    Private Sub PanelVision_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs)

    End Sub

    Private Sub ShowLabelMessage(ByVal msg As String, Optional ByVal isError As Boolean = False)
        If isError Then
            LabelMessage.ForeColor = System.Drawing.Color.Red
        Else
            LabelMessage.ForeColor = System.Drawing.Color.Black
        End If
        LabelMessage.Text = msg
    End Sub

    Private Sub btLaserInterface_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btLaserInterface.Click
        OnLaser()
        Dim temp As CIDSService.CIDSServiceDevices.CIDSServiceLaser = Laser.Instance
        temp.Show()
    End Sub

    Private Sub SetSelectButton(ByVal sender As System.Object)
        If Not (selectedButton Is Nothing) Then
            selectedButton.FlatStyle = FlatStyle.Standard
            selectedButton.BackColor = System.Drawing.SystemColors.Control
        End If
        selectedButton = DirectCast(sender, System.Windows.Forms.Button)
        selectedButton.FlatStyle = FlatStyle.Flat
        selectedButton.BackColor = System.Drawing.SystemColors.ControlDarkDark
    End Sub

    Private Sub HomingCheckTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HomingCheckTimer.Tick
        HomingCheckTimer.Stop()
        HomingCheckTimer.Enabled = False
        If Not (homingDone) Then
            ShowLabelMessage("Homing failed......")
            For Each ctrl As Control In Me.Controls
                ctrl.Enabled = True
            Next
        End If
    End Sub

    Private Sub rbOneHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOneHead.CheckedChanged
        If rbOneHead.Checked Then
            IDS.Data.Hardware.Dispenser.CurrentHeads = 1
            MyGantrySetup.RightHead.Visible = False
            MyGantrySetup.LeftHead.Checked = True
            MyGantrySettings.RightHead.Visible = False
            MyGantrySettings.LeftHead.Checked = True
            MyNeedleCalibrationSettings.rbRightHead.Visible = False
            MyNeedleCalibrationSettings.rbLeftHead.Checked = True
            MyNeedleCalibrationSetup1.rbRightHead.Visible = False
            MyNeedleCalibrationSetup1.rbLeftHead.Checked = True
        Else
            IDS.Data.Hardware.Dispenser.CurrentHeads = 2
            MyGantrySetup.RightHead.Visible = True
            MyGantrySettings.RightHead.Visible = True
            MyNeedleCalibrationSettings.rbRightHead.Visible = True
            MyNeedleCalibrationSetup1.rbRightHead.Visible = True
        End If
        IDS.Data.SaveLocalData()
        MyDispenserSettings.RevertData()
    End Sub

    Private Sub MoveZToZero()
        If homingDone Then
            If m_Tri.ZPosition < -0.5 Then
                Me.Cursor = Cursors.WaitCursor
                m_Tri.Set_Z_Speed(50)
                m_Tri.Move_Z(0)
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub

    Private Sub tbEquipmentID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbEquipmentID.TextChanged
        CIDSData.Machine.EquipmentID = tbEquipmentID.Text
        IDS.Data.SaveGlobalData()
    End Sub

    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExit.Click
        Dim fm As InfoForm = New InfoForm
        fm.Height = 130
        fm.SetMessage("Are you sure you want to exit the program? Click Ok to exit.")
        If fm.ShowDialog() = DialogResult.Cancel Then
            Return
        End If
        selectedButton.FlatStyle = FlatStyle.Standard
        selectedButton.BackColor = System.Drawing.SystemColors.Control
        Close()
    End Sub
End Class
