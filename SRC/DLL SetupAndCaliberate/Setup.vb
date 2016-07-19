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
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents PanelVisionCtrl As System.Windows.Forms.Panel
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents PanelRight As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBoxRobotZ As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotY As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxRobotX As System.Windows.Forms.TextBox
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ButtonSystemIO As System.Windows.Forms.Button
    Friend WithEvents GpB_Configurations As System.Windows.Forms.GroupBox
    Friend WithEvents RBn_LVDT As System.Windows.Forms.RadioButton
    Friend WithEvents RBn_Laser As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBoxLifter As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxHeightSensor As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBoxHeater As System.Windows.Forms.CheckBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents CheckBoxVolume As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonGantrySetup As System.Windows.Forms.Button
    Friend WithEvents ButtonCameraSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonLVDTSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonLaserSetup As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenserSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonStationPositions As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonSPCLogging As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonThermalSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonEventHandling As System.Windows.Forms.Button
    Friend WithEvents ButtonVolumeCalibSettings As System.Windows.Forms.Button
    Friend WithEvents OneHead As System.Windows.Forms.CheckBox
    Friend WithEvents TwoHead As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonHardwareCommunicationSetup As System.Windows.Forms.Button
    Public WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents DisplayBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Timer1 As System.Timers.Timer
    Friend WithEvents HardwareInitTimer As System.Windows.Forms.Timer
    Friend WithEvents LabelMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Setup))
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.PanelVisionCtrl = New System.Windows.Forms.Panel
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GpB_Configurations = New System.Windows.Forms.GroupBox
        Me.RBn_LVDT = New System.Windows.Forms.RadioButton
        Me.RBn_Laser = New System.Windows.Forms.RadioButton
        Me.CheckBoxLifter = New System.Windows.Forms.CheckBox
        Me.CheckBoxHeightSensor = New System.Windows.Forms.CheckBox
        Me.CheckBoxHeater = New System.Windows.Forms.CheckBox
        Me.CheckBoxVolume = New System.Windows.Forms.CheckBox
        Me.OneHead = New System.Windows.Forms.CheckBox
        Me.TwoHead = New System.Windows.Forms.CheckBox
        Me.LabelMessage = New System.Windows.Forms.Label
        Me.PanelRight = New System.Windows.Forms.Panel
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label8 = New System.Windows.Forms.Label
        Me.ButtonGantrySetup = New System.Windows.Forms.Button
        Me.ButtonCameraSetup = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSetup = New System.Windows.Forms.Button
        Me.ButtonHardwareCommunicationSetup = New System.Windows.Forms.Button
        Me.ButtonSystemIO = New System.Windows.Forms.Button
        Me.ButtonLaserSetup = New System.Windows.Forms.Button
        Me.ButtonLVDTSetup = New System.Windows.Forms.Button
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label7 = New System.Windows.Forms.Label
        Me.ButtonStationPositions = New System.Windows.Forms.Button
        Me.ButtonVolumeCalibSettings = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSettings = New System.Windows.Forms.Button
        Me.ButtonEventHandling = New System.Windows.Forms.Button
        Me.ButtonConveyorSettings = New System.Windows.Forms.Button
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.DisplayBrightness = New System.Windows.Forms.NumericUpDown
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.Timer1 = New System.Timers.Timer
        Me.HardwareInitTimer = New System.Windows.Forms.Timer(Me.components)
        Me.PanelVision.SuspendLayout()
        Me.PanelVisionCtrl.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GpB_Configurations.SuspendLayout()
        Me.PanelRight.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DisplayBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.Controls.Add(Me.PanelVisionCtrl)
        Me.PanelVision.Location = New System.Drawing.Point(0, 340)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(768, 576)
        Me.PanelVision.TabIndex = 3
        '
        'PanelVisionCtrl
        '
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBox2)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBox4)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBox1)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBox3)
        Me.PanelVisionCtrl.Controls.Add(Me.CheckBox6)
        Me.PanelVisionCtrl.Controls.Add(Me.TextBox2)
        Me.PanelVisionCtrl.Controls.Add(Me.Label4)
        Me.PanelVisionCtrl.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelVisionCtrl.Location = New System.Drawing.Point(0, 920)
        Me.PanelVisionCtrl.Name = "PanelVisionCtrl"
        Me.PanelVisionCtrl.Size = New System.Drawing.Size(752, 24)
        Me.PanelVisionCtrl.TabIndex = 4
        '
        'CheckBox2
        '
        Me.CheckBox2.BackColor = System.Drawing.Color.Gainsboro
        Me.CheckBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Image = CType(resources.GetObject("CheckBox2.Image"), System.Drawing.Image)
        Me.CheckBox2.Location = New System.Drawing.Point(712, 3)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox2.TabIndex = 75
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(656, 2)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(56, 20)
        Me.TextBox4.TabIndex = 74
        Me.TextBox4.Text = "Z 100.000"
        '
        'CheckBox1
        '
        Me.CheckBox1.BackColor = System.Drawing.Color.Gainsboro
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Image = CType(resources.GetObject("CheckBox1.Image"), System.Drawing.Image)
        Me.CheckBox1.Location = New System.Drawing.Point(608, 3)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox1.TabIndex = 73
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(552, 2)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(56, 20)
        Me.TextBox3.TabIndex = 72
        Me.TextBox3.Text = "Y 100.000"
        '
        'CheckBox6
        '
        Me.CheckBox6.BackColor = System.Drawing.Color.Gainsboro
        Me.CheckBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox6.Image = CType(resources.GetObject("CheckBox6.Image"), System.Drawing.Image)
        Me.CheckBox6.Location = New System.Drawing.Point(504, 3)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox6.TabIndex = 71
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(448, 2)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(56, 20)
        Me.TextBox2.TabIndex = 6
        Me.TextBox2.Text = "X 100.000"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(408, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 23)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.GpB_Configurations)
        Me.Panel1.Controls.Add(Me.LabelMessage)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(768, 340)
        Me.Panel1.TabIndex = 5
        '
        'GpB_Configurations
        '
        Me.GpB_Configurations.Controls.Add(Me.RBn_LVDT)
        Me.GpB_Configurations.Controls.Add(Me.RBn_Laser)
        Me.GpB_Configurations.Controls.Add(Me.CheckBoxLifter)
        Me.GpB_Configurations.Controls.Add(Me.CheckBoxHeightSensor)
        Me.GpB_Configurations.Controls.Add(Me.CheckBoxHeater)
        Me.GpB_Configurations.Controls.Add(Me.CheckBoxVolume)
        Me.GpB_Configurations.Controls.Add(Me.OneHead)
        Me.GpB_Configurations.Controls.Add(Me.TwoHead)
        Me.GpB_Configurations.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GpB_Configurations.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GpB_Configurations.Location = New System.Drawing.Point(28, 8)
        Me.GpB_Configurations.Name = "GpB_Configurations"
        Me.GpB_Configurations.Size = New System.Drawing.Size(336, 280)
        Me.GpB_Configurations.TabIndex = 62
        Me.GpB_Configurations.TabStop = False
        Me.GpB_Configurations.Text = "Configurations"
        '
        'RBn_LVDT
        '
        Me.RBn_LVDT.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBn_LVDT.Location = New System.Drawing.Point(176, 112)
        Me.RBn_LVDT.Name = "RBn_LVDT"
        Me.RBn_LVDT.TabIndex = 68
        Me.RBn_LVDT.Text = "LVDT"
        '
        'RBn_Laser
        '
        Me.RBn_Laser.Checked = True
        Me.RBn_Laser.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RBn_Laser.Location = New System.Drawing.Point(176, 80)
        Me.RBn_Laser.Name = "RBn_Laser"
        Me.RBn_Laser.Size = New System.Drawing.Size(144, 24)
        Me.RBn_Laser.TabIndex = 67
        Me.RBn_Laser.TabStop = True
        Me.RBn_Laser.Text = "Laser Sensor"
        '
        'CheckBoxLifter
        '
        Me.CheckBoxLifter.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxLifter.Location = New System.Drawing.Point(16, 224)
        Me.CheckBoxLifter.Name = "CheckBoxLifter"
        Me.CheckBoxLifter.Size = New System.Drawing.Size(184, 23)
        Me.CheckBoxLifter.TabIndex = 66
        Me.CheckBoxLifter.Text = "Lifter Integration"
        '
        'CheckBoxHeightSensor
        '
        Me.CheckBoxHeightSensor.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxHeightSensor.Location = New System.Drawing.Point(16, 80)
        Me.CheckBoxHeightSensor.Name = "CheckBoxHeightSensor"
        Me.CheckBoxHeightSensor.Size = New System.Drawing.Size(144, 23)
        Me.CheckBoxHeightSensor.TabIndex = 65
        Me.CheckBoxHeightSensor.Text = "Height Sensor"
        '
        'CheckBoxHeater
        '
        Me.CheckBoxHeater.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxHeater.Location = New System.Drawing.Point(16, 144)
        Me.CheckBoxHeater.Name = "CheckBoxHeater"
        Me.CheckBoxHeater.Size = New System.Drawing.Size(184, 24)
        Me.CheckBoxHeater.TabIndex = 63
        Me.CheckBoxHeater.Text = "Heater Integration"
        '
        'CheckBoxVolume
        '
        Me.CheckBoxVolume.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBoxVolume.Location = New System.Drawing.Point(16, 184)
        Me.CheckBoxVolume.Name = "CheckBoxVolume"
        Me.CheckBoxVolume.Size = New System.Drawing.Size(184, 23)
        Me.CheckBoxVolume.TabIndex = 66
        Me.CheckBoxVolume.Text = "Volume Calibration"
        '
        'OneHead
        '
        Me.OneHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OneHead.Location = New System.Drawing.Point(16, 40)
        Me.OneHead.Name = "OneHead"
        Me.OneHead.Size = New System.Drawing.Size(144, 23)
        Me.OneHead.TabIndex = 65
        Me.OneHead.Text = "One Head"
        '
        'TwoHead
        '
        Me.TwoHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TwoHead.Location = New System.Drawing.Point(176, 40)
        Me.TwoHead.Name = "TwoHead"
        Me.TwoHead.Size = New System.Drawing.Size(144, 23)
        Me.TwoHead.TabIndex = 65
        Me.TwoHead.Text = "Two Heads"
        '
        'LabelMessage
        '
        Me.LabelMessage.BackColor = System.Drawing.Color.White
        Me.LabelMessage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessage.ForeColor = System.Drawing.Color.Black
        Me.LabelMessage.Location = New System.Drawing.Point(0, 304)
        Me.LabelMessage.Name = "LabelMessage"
        Me.LabelMessage.Size = New System.Drawing.Size(768, 32)
        Me.LabelMessage.TabIndex = 85
        Me.LabelMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PanelRight
        '
        Me.PanelRight.Controls.Add(Me.PanelToBeAdded)
        Me.PanelRight.Controls.Add(Me.Panel3)
        Me.PanelRight.Location = New System.Drawing.Point(768, 0)
        Me.PanelRight.Name = "PanelRight"
        Me.PanelRight.Size = New System.Drawing.Size(512, 944)
        Me.PanelRight.TabIndex = 6
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label8)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonGantrySetup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonCameraSetup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonNeedleCalibSetup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonHardwareCommunicationSetup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSystemIO)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonLaserSetup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonLVDTSetup)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(0, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 200)
        Me.PanelToBeAdded.TabIndex = 64
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(16, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(160, 32)
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "System Setup"
        '
        'ButtonGantrySetup
        '
        Me.ButtonGantrySetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonGantrySetup.Location = New System.Drawing.Point(264, 40)
        Me.ButtonGantrySetup.Name = "ButtonGantrySetup"
        Me.ButtonGantrySetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonGantrySetup.TabIndex = 43
        Me.ButtonGantrySetup.Text = "Gantry Parameters"
        '
        'ButtonCameraSetup
        '
        Me.ButtonCameraSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonCameraSetup.Location = New System.Drawing.Point(264, 96)
        Me.ButtonCameraSetup.Name = "ButtonCameraSetup"
        Me.ButtonCameraSetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonCameraSetup.TabIndex = 42
        Me.ButtonCameraSetup.Text = "Camera Setup"
        '
        'ButtonNeedleCalibSetup
        '
        Me.ButtonNeedleCalibSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSetup.Location = New System.Drawing.Point(24, 152)
        Me.ButtonNeedleCalibSetup.Name = "ButtonNeedleCalibSetup"
        Me.ButtonNeedleCalibSetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonNeedleCalibSetup.TabIndex = 44
        Me.ButtonNeedleCalibSetup.Text = "Needle Setup"
        '
        'ButtonHardwareCommunicationSetup
        '
        Me.ButtonHardwareCommunicationSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonHardwareCommunicationSetup.Location = New System.Drawing.Point(24, 40)
        Me.ButtonHardwareCommunicationSetup.Name = "ButtonHardwareCommunicationSetup"
        Me.ButtonHardwareCommunicationSetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonHardwareCommunicationSetup.TabIndex = 46
        Me.ButtonHardwareCommunicationSetup.Text = "Hardware Communication"
        '
        'ButtonSystemIO
        '
        Me.ButtonSystemIO.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSystemIO.Location = New System.Drawing.Point(264, 152)
        Me.ButtonSystemIO.Name = "ButtonSystemIO"
        Me.ButtonSystemIO.Size = New System.Drawing.Size(224, 30)
        Me.ButtonSystemIO.TabIndex = 37
        Me.ButtonSystemIO.Text = "System IO"
        '
        'ButtonLaserSetup
        '
        Me.ButtonLaserSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonLaserSetup.Location = New System.Drawing.Point(24, 96)
        Me.ButtonLaserSetup.Name = "ButtonLaserSetup"
        Me.ButtonLaserSetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonLaserSetup.TabIndex = 45
        Me.ButtonLaserSetup.Text = "Laser Setup"
        Me.ButtonLaserSetup.Visible = False
        '
        'ButtonLVDTSetup
        '
        Me.ButtonLVDTSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonLVDTSetup.Location = New System.Drawing.Point(24, 96)
        Me.ButtonLVDTSetup.Name = "ButtonLVDTSetup"
        Me.ButtonLVDTSetup.Size = New System.Drawing.Size(224, 30)
        Me.ButtonLVDTSetup.TabIndex = 50
        Me.ButtonLVDTSetup.Text = "LVDT Calibration"
        Me.ButtonLVDTSetup.Visible = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.ButtonStationPositions)
        Me.Panel3.Controls.Add(Me.ButtonVolumeCalibSettings)
        Me.Panel3.Controls.Add(Me.ButtonNeedleCalibSettings)
        Me.Panel3.Controls.Add(Me.ButtonEventHandling)
        Me.Panel3.Controls.Add(Me.ButtonConveyorSettings)
        Me.Panel3.Controls.Add(Me.ButtonThermalSettings)
        Me.Panel3.Controls.Add(Me.ButtonSPCLogging)
        Me.Panel3.Controls.Add(Me.ButtonDispenserSettings)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel3.Location = New System.Drawing.Point(0, 224)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(512, 272)
        Me.Panel3.TabIndex = 64
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(16, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(176, 32)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Global Settings"
        '
        'ButtonStationPositions
        '
        Me.ButtonStationPositions.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStationPositions.Location = New System.Drawing.Point(264, 48)
        Me.ButtonStationPositions.Name = "ButtonStationPositions"
        Me.ButtonStationPositions.Size = New System.Drawing.Size(224, 30)
        Me.ButtonStationPositions.TabIndex = 43
        Me.ButtonStationPositions.Text = "Gantry Settings"
        '
        'ButtonVolumeCalibSettings
        '
        Me.ButtonVolumeCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolumeCalibSettings.Location = New System.Drawing.Point(264, 104)
        Me.ButtonVolumeCalibSettings.Name = "ButtonVolumeCalibSettings"
        Me.ButtonVolumeCalibSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonVolumeCalibSettings.TabIndex = 46
        Me.ButtonVolumeCalibSettings.Text = "Volume Calibration"
        Me.ButtonVolumeCalibSettings.Visible = False
        '
        'ButtonNeedleCalibSettings
        '
        Me.ButtonNeedleCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSettings.Location = New System.Drawing.Point(24, 104)
        Me.ButtonNeedleCalibSettings.Name = "ButtonNeedleCalibSettings"
        Me.ButtonNeedleCalibSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonNeedleCalibSettings.TabIndex = 44
        Me.ButtonNeedleCalibSettings.Text = "Needle Calibration"
        '
        'ButtonEventHandling
        '
        Me.ButtonEventHandling.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEventHandling.Location = New System.Drawing.Point(264, 160)
        Me.ButtonEventHandling.Name = "ButtonEventHandling"
        Me.ButtonEventHandling.Size = New System.Drawing.Size(224, 30)
        Me.ButtonEventHandling.TabIndex = 47
        Me.ButtonEventHandling.Text = "Event Handling"
        '
        'ButtonConveyorSettings
        '
        Me.ButtonConveyorSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonConveyorSettings.Location = New System.Drawing.Point(24, 216)
        Me.ButtonConveyorSettings.Name = "ButtonConveyorSettings"
        Me.ButtonConveyorSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonConveyorSettings.TabIndex = 45
        Me.ButtonConveyorSettings.Text = "Conveyor"
        '
        'ButtonThermalSettings
        '
        Me.ButtonThermalSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonThermalSettings.Location = New System.Drawing.Point(264, 216)
        Me.ButtonThermalSettings.Name = "ButtonThermalSettings"
        Me.ButtonThermalSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonThermalSettings.TabIndex = 46
        Me.ButtonThermalSettings.Text = "Heater"
        Me.ButtonThermalSettings.Visible = False
        '
        'ButtonSPCLogging
        '
        Me.ButtonSPCLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCLogging.Location = New System.Drawing.Point(24, 160)
        Me.ButtonSPCLogging.Name = "ButtonSPCLogging"
        Me.ButtonSPCLogging.Size = New System.Drawing.Size(224, 30)
        Me.ButtonSPCLogging.TabIndex = 51
        Me.ButtonSPCLogging.Text = "SPC Logging"
        '
        'ButtonDispenserSettings
        '
        Me.ButtonDispenserSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenserSettings.Location = New System.Drawing.Point(24, 48)
        Me.ButtonDispenserSettings.Name = "ButtonDispenserSettings"
        Me.ButtonDispenserSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonDispenserSettings.TabIndex = 42
        Me.ButtonDispenserSettings.Text = "Dispenser"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.CheckBox3)
        Me.Panel2.Controls.Add(Me.TextBoxRobotZ)
        Me.Panel2.Controls.Add(Me.CheckBox4)
        Me.Panel2.Controls.Add(Me.TextBoxRobotY)
        Me.Panel2.Controls.Add(Me.CheckBox5)
        Me.Panel2.Controls.Add(Me.TextBoxRobotX)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.TextBox7)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.DisplayBrightness)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 916)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(768, 28)
        Me.Panel2.TabIndex = 5
        '
        'CheckBox3
        '
        Me.CheckBox3.BackColor = System.Drawing.SystemColors.Control
        Me.CheckBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox3.Image = CType(resources.GetObject("CheckBox3.Image"), System.Drawing.Image)
        Me.CheckBox3.Location = New System.Drawing.Point(725, 5)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox3.TabIndex = 75
        '
        'TextBoxRobotZ
        '
        Me.TextBoxRobotZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotZ.Location = New System.Drawing.Point(650, 4)
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
        Me.CheckBox4.Location = New System.Drawing.Point(611, 4)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox4.TabIndex = 73
        '
        'TextBoxRobotY
        '
        Me.TextBoxRobotY.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotY.Location = New System.Drawing.Point(536, 4)
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
        Me.CheckBox5.Location = New System.Drawing.Point(498, 4)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(40, 16)
        Me.CheckBox5.TabIndex = 71
        '
        'TextBoxRobotX
        '
        Me.TextBoxRobotX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotX.Location = New System.Drawing.Point(424, 4)
        Me.TextBoxRobotX.Name = "TextBoxRobotX"
        Me.TextBoxRobotX.ReadOnly = True
        Me.TextBoxRobotX.Size = New System.Drawing.Size(74, 21)
        Me.TextBoxRobotX.TabIndex = 6
        Me.TextBoxRobotX.Text = "X: 100.000"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(384, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Robot"
        '
        'TextBox7
        '
        Me.TextBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox7.Location = New System.Drawing.Point(55, 4)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.ReadOnly = True
        Me.TextBox7.Size = New System.Drawing.Size(145, 21)
        Me.TextBox7.TabIndex = 0
        Me.TextBox7.Text = "Z: 100.000,  Y: 100.000"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(245, 6)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(66, 23)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Brightness"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 6)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Cursor"
        '
        'DisplayBrightness
        '
        Me.DisplayBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DisplayBrightness.Location = New System.Drawing.Point(313, 4)
        Me.DisplayBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.DisplayBrightness.Name = "DisplayBrightness"
        Me.DisplayBrightness.Size = New System.Drawing.Size(50, 21)
        Me.DisplayBrightness.TabIndex = 5
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
        'Timer1
        '
        Me.Timer1.SynchronizingObject = Me
        '
        'HardwareInitTimer
        '
        '
        'Setup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1272, 990)
        Me.Controls.Add(Me.PanelRight)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.Panel1)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "Setup"
        Me.Text = "System Setup"
        Me.PanelVision.ResumeLayout(False)
        Me.PanelVisionCtrl.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.GpB_Configurations.ResumeLayout(False)
        Me.PanelRight.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DisplayBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()
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
    Dim m_keyBoard As New DLL_Export_Device_Motor.Keyboard(Me)
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
    Dim homingStart As Boolean
    Dim homingDone As Boolean
    Dim homingWaitCnt As Integer

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
            End If

        End If

    End Sub

    Private Sub TheThreadMonitor()
        Do
            TraceDoEvents()
            MySleep(50)

            'we may want more detailed functionality here
            m_Tri.GetIDSState()

            TextBoxRobotX.Text = "x: " & m_Tri.XPosition.ToString
            TextBoxRobotY.Text = "Y: " & m_Tri.YPosition.ToString
            TextBoxRobotZ.Text = "Z: " & m_Tri.ZPosition.ToString

            If m_Tri.HomingFinished() Then
                homingDone = True
                GC.Collect()
                m_Tri.SteppingButtons.Enabled = True
            End If
        Loop
    End Sub

    Private Sub Setup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        KeyboardControl.GainControls()

        ' get default value from the default pat file
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        OpenPathFileName("C:\IDS\Pattern_Dir\FactoryDefault.Pat")
        DisplayBrightness.Value = IDSData.Hardware.Camera.Brightness
        IDS.Devices.Vision.FrmVision.SetBrightness(IDSData.Hardware.Camera.Brightness)

        homingStart = False
        homingDone = False
        homingWaitCnt = 0
        'display the stepping panel
        Panel1.Controls.Add(m_Tri.SteppingButtons)
        m_Tri.SteppingButtons.BringToFront()
        m_Tri.SteppingButtons.Enabled = False
        m_Tri.SteppingButtons.Show()
        'change GUI settings unique for system setup here
        m_Tri.SteppingButtons.Location = New System.Drawing.Point(404, GpB_Configurations.Location.Y)
        'gui
        PanelVision.Controls.Add(IDS.Devices.Vision.FrmVision.PanelVision) 'lsgoh
        Location = New Point(0, 0)
        MyBase.StartPosition = FormStartPosition.Manual

        Conn = OleDbConnection1
        Call DataLoad()
        FillFromSystemConfigureTable() 'lsgoh

        'kr
        If IDS.Data.Hardware.Dispenser.CurrentHeads = 1 Then
            OneHead.Checked = True
        Else
            OneHead.Checked = False
        End If

        EnableControls(False)
        'timers
        IDS.StartErrorCheck()
        Conveyor.PositionTimer.Start()
        Timer1.Start()
        ThreadMonitor = New Threading.Thread(AddressOf TheThreadMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.BelowNormal
        ThreadMonitor.Start()
        MouseTimer = New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 180)
        'hardware
        Laser.OpenPort()
        OnLaser()
        LabelMessage.Text = "Initialize Hardware......"
        HardwareInitTimer.Enabled = True
        HardwareInitTimer.Start()
    End Sub

    Private Sub Setup_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        KeyboardControl.ReleaseControls()
        'gui clear
        UpdatetoSystemConfigureTable()
        PanelVision.Controls.Clear()
        MySettings.Controls.Clear()
        PanelRight.Controls.Remove(CurrentControl)
        MyGantrySettings.Controls.Clear()

        Weighting_Scale.Instance.Hide()
        Laser.Instance.Hide()
        Conveyor.Instance.Hide()

        'motion controller stop
        m_Tri.Disconnect_Controller()

        'disable stepping buttons!
        m_Tri.SteppingButtons.Enabled = False

        'timers stop
        IDS.StopErrorCheck()
        ThreadMonitor.Abort()
        MouseTimer.Dispose()
        Timer1.Stop()
        Conveyor.PositionTimer.Stop()

        'hardware
        OffLaser()
        Laser.ClosePort()
        Conveyor.ClosePort()
        Weighting_Scale.ClosePort()

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

    Private Sub DisplayBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DisplayBrightness.ValueChanged
        IDS.Devices.Vision.IDSV_SetBrightness(DisplayBrightness.Value)
    End Sub

    Private Sub Setup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim mode As Integer = 0

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

            m_Tri.RunTrioMotionProgram("EXITIDS")
        End If
        HardwareInitTimer.Stop()
        HardwareInitTimer.Enabled = False
        If Not (MySysIO.IOReadClose) Then
            MySysIO.ClearandClose()
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

                If Row("Hardware") = "HeightSensor" Then
                    CheckBoxHeightSensor.Checked = CBoolean(Row("Selected"))

                ElseIf Row("Hardware") = "VolumeCalibration" Then
                    CheckBoxVolume.Checked = CBoolean(Row("Selected"))

                ElseIf Row("Hardware") = "ThermalController" Then
                    CheckBoxHeater.Checked = CBoolean(Row("Selected"))

                ElseIf Row("Hardware") = "Lifter&Vacuum" Then
                    CheckBoxLifter.Checked = CBoolean(Row("Selected"))
                End If
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

                If Row("Hardware") = "HeightSensor" Then
                    Row("Selected") = CheckBoxHeightSensor.Checked

                ElseIf Row("Hardware") = "VolumeCalibration" Then
                    Row("Selected") = CheckBoxVolume.Checked

                ElseIf Row("Hardware") = "ThermalController" Then
                    Row("Selected") = CheckBoxHeater.Checked

                ElseIf Row("Hardware") = "Lifter&Vacuum" Then
                    Row("Selected") = CheckBoxLifter.Checked

                End If
            Next
        End If
        UpdateData("SELECT * FROM " + TableName, TableName)

    End Sub

    Private Sub CheckBoxHeightSensor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxHeightSensor.CheckedChanged
        'Present flag
        If CheckBoxHeightSensor.Checked = True Then
            Me.RBn_Laser.Enabled = True
            Me.RBn_LVDT.Enabled = True
            If Me.RBn_Laser.Checked = True Then
                ButtonLaserSetup.Visible = True
                ButtonLVDTSetup.Visible = False

            Else
                ButtonLaserSetup.Visible = False
                ButtonLVDTSetup.Visible = True
            End If
            IDSData.Hardware.HeightSensor.TP.OffsetPos.X = 1
            IDSData.Hardware.HeightSensor.TP.OffsetPos.Y = 1
            'use TP OffsetPos=0 to indicate that no height sensor is in use 'Xu Long
        Else
            ButtonLaserSetup.Visible = False
            ButtonLVDTSetup.Visible = False
            IDSData.Hardware.HeightSensor.TP.OffsetPos.X = 0
            IDSData.Hardware.HeightSensor.TP.OffsetPos.Y = 0
            'use TP OffsetPos=0 to indicate that no height sensor is in use 'Xu Long
        End If
    End Sub


    Private Sub RBn_Laser_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBn_Laser.CheckedChanged
        If Me.RBn_Laser.Checked = True Then
            IDSData.Hardware.HeightSensor.SelectType = False       'False - Laser sensor is selected
            ButtonLaserSetup.Visible = True
            ButtonLVDTSetup.Visible = False
        Else
            IDSData.Hardware.HeightSensor.SelectType = True        'True - LVDT is selected
            ButtonLaserSetup.Visible = False
            ButtonLVDTSetup.Visible = True
        End If
    End Sub
    Private Sub RBn_LVDT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RBn_LVDT.CheckedChanged
        If Me.RBn_LVDT.Checked = True Then
            IDSData.Hardware.HeightSensor.SelectType = True
            ButtonLaserSetup.Visible = False
            ButtonLVDTSetup.Visible = True
        Else
            IDSData.Hardware.HeightSensor.SelectType = False
            ButtonLaserSetup.Visible = True
            ButtonLVDTSetup.Visible = False
        End If
    End Sub

    Private Sub CheckBoxVolume_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxVolume.CheckedChanged
        'Present flag
        If CheckBoxVolume.Checked = True Then
            ButtonVolumeCalibSettings.Visible = True
        Else
            ButtonVolumeCalibSettings.Visible = False
        End If
        MyHardwareCommunicationSetup.RefreshDisplay()
        MyHardwareCommunicationSetup.UpdateStatus()
    End Sub

    Private Sub CheckBoxHeater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxHeater.CheckedChanged
        'Present flag
        If CheckBoxHeater.Checked = True Then
            ButtonThermalSettings.Visible = True
        Else
            ButtonThermalSettings.Visible = False
        End If
        MyHardwareCommunicationSetup.RefreshDisplay()
        MyHardwareCommunicationSetup.UpdateStatus()
        IDS.Data.Hardware.Thermal.HeaterFeatureEnabled = ButtonThermalSettings.Visible
        IDS.Data.SaveData()
    End Sub

    Private Sub ButtonSystemIO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSystemIO.Click
        AddPanel(PanelRight, MySysIO.PanelToBeAdded)
        PanelToBeAdded.Width = 1280
        PanelRight.Width = 1280
        PanelRight.Location = New Point(0, 0)
        PanelRight.BringToFront()
        MySysIO.SetIOTable()

        IDS.Data.Admin.User.RunApplication = "System"
        MySysIO.IOReadClose = False
        MySysIO.IOReadStop = False
        MySysIO.TimerIO.Enabled = True
        MySysIO.TimerIO.Start()

        MySysIO.EditIO()
    End Sub

    Private Sub ButtonGantry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGantrySetup.Click
        AddPanel(PanelRight, MyGantrySetup.PanelToBeAdded)
        If OneHead.Checked Then
            MyGantrySetup.RightHead.Visible = False
        Else
            MyGantrySetup.RightHead.Visible = True
        End If
        MyGantrySetup.RevertData()
    End Sub

    Private Sub ButtonLaser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLaserSetup.Click
        OnLaser()
        AddPanel(PanelRight, MyLaser.PanelToBeAdded)
        MyLaser.LaserOffsetZ.Text = IDS.Data.Hardware.HeightSensor.Laser.HeightReference
        MyLaser.LaserOffsetX.Text = IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X
        MyLaser.LaserOffsetY.Text = IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y
        MyLaser.RevertData()
    End Sub

    Private Sub ButtonNeedle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCalibSetup.Click
        AddPanel(PanelRight, MyNeedleCalibrationSetup1.PanelToBeAdded)
        MyNeedleCalibrationSetup1.RevertData()
        With IDS.Data.Hardware
            MyNeedleCalibrationSetup1.NSensorXPos.Text = .Needle.Right.CalibratorPos.X - .Camera.ReferencePos.X
            MyNeedleCalibrationSetup1.NSensorYPos.Text = .Needle.Right.CalibratorPos.Y - .Camera.ReferencePos.Y
            MyNeedleCalibrationSetup1.NSensorZPos.Text = .Needle.Right.TouchSensorZPosition - .Needle.Left.CalibratorPos.Z

            MyNeedleCalibrationSetup1.LNeedleXPos.Text = .Needle.Left.NeedleCalibrationPosition.X - .Needle.Right.CalibratorPos.X
            MyNeedleCalibrationSetup1.LNeedleYPos.Text = .Needle.Left.NeedleCalibrationPosition.Y - .Needle.Right.CalibratorPos.Y
            MyNeedleCalibrationSetup1.LNeedleZPos.Text = .Needle.Left.NeedleCalibrationPosition.Z

            MyNeedleCalibrationSetup1.RNeedleXPos.Text = .Needle.Right.NeedleCalibrationPosition.X - .Needle.Right.CalibratorPos.X
            MyNeedleCalibrationSetup1.RNeedleYPos.Text = .Needle.Right.NeedleCalibrationPosition.Y - .Needle.Right.CalibratorPos.Y
            MyNeedleCalibrationSetup1.RNeedleZPos.Text = .Needle.Right.NeedleCalibrationPosition.Z
        End With
    End Sub

    Private Sub ButtonLVDT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLVDTSetup.Click
        AddPanel(PanelRight, MyLVDTSetup.PanelToBeAdded)
    End Sub

#Region " Data Reference "
    ' this function need to be removed after it is compiled as DLL .
    Private Sub SystemConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        IDS.Devices.Vision.IDSV_SetBrightness(DisplayBrightness.Value)
        Conn = OleDbConnection1
        Call DataLoad()

    End Sub
#End Region

    Private Sub ButtonCamera_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCameraSetup.Click
        AddPanel(PanelRight, MyVisionSetup.PanelToBeAdded)
        MyVisionSetup.SetTable_a()
        MyVisionSetup.SetTable_b()
    End Sub


    Private Sub ButtonHardwareCommunicationSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHardwareCommunicationSetup.Click
        AddPanel(PanelRight, MyHardwareCommunicationSetup.PanelToBeAdded)
    End Sub

    Private Sub ButtonNeedleCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCalibSettings.Click
        OnLaser()
        AddPanel(PanelRight, MyNeedleCalibrationSettings.PanelToBeAdded)
        If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Then
            MyNeedleCalibrationSettings.BoxStep1.Visible = False
            MyNeedleCalibrationSettings.BoxStep1Jetting.Visible = True
        Else
            MyNeedleCalibrationSettings.BoxStep1.Visible = True
            MyNeedleCalibrationSettings.BoxStep1Jetting.Visible = False
        End If
        MyNeedleCalibrationSettings.RevertData()
    End Sub

    Private Sub ButtonStationPositions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStationPositions.Click
        AddPanel(PanelRight, MyGantrySettings.PanelToBeAdded)
        'initialize
        MyGantrySettings.RevertData()
        MyGantrySettings.StationPosition.SelectedIndex = 0
        MyGantrySettings.LeftHead.Checked = True
    End Sub

    Private Sub ButtonDispenserSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDispenserSettings.Click
        AddPanel(PanelRight, MyDispenserSettings.PanelToBeAdded)
        MyDispenserSettings.HeadType.SelectedIndex = 0
        MyDispenserSettings.RevertData()
    End Sub

    Private Sub ButtonSPCLogging_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPCLogging.Click
        AddPanel(PanelRight, MySPCLogging.PanelToBeAdded)
        MySPCLogging.RevertData()
    End Sub

    Private Sub ButtonEventHandling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEventHandling.Click
        AddPanel(PanelRight, MyEventSettings.PanelToBeAdded)
        MyEventSettings.RevertData()
    End Sub

    Private Sub ButtonConveyorSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorSettings.Click
        AddPanel(PanelRight, MyConveyorSettings.PanelToBeAdded)
        Conveyor.OpenPort()
        Conveyor.PositionTimer.Start()
        MyConveyorSettings.PositionTimer.Start()
        MySettings.RevertData()
    End Sub

    Private Sub ButtonThermalSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonThermalSettings.Click
        AddPanel(PanelRight, MyHeaterSettings.PanelToBeAdded)
        MyHeaterSettings.RevertData()
    End Sub

    Private Sub ButtonVolumeCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolumeCalibSettings.Click
        AddPanel(PanelRight, MyVolumeCalibrationSettings.PanelToBeAdded)
        MyVolumeCalibrationSettings.RevertData()
    End Sub

    Private Sub OneHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OneHead.CheckStateChanged
        If OneHead.Checked Then
            TwoHead.Checked = False
            IDS.Data.Hardware.Dispenser.CurrentHeads = 1
            MyGantrySetup.RightHead.Visible = False
        Else
            TwoHead.Checked = True
            IDS.Data.Hardware.Dispenser.CurrentHeads = 2
            MyGantrySetup.RightHead.Visible = True
        End If
        IDS.Data.SaveData()
    End Sub

    Private Sub TwoHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TwoHead.CheckStateChanged
        If TwoHead.Checked Then
            OneHead.Checked = False
            IDS.Data.Hardware.Dispenser.CurrentHeads = 2
            MyDispenserSettings.RevertData()
        Else
            OneHead.Checked = True
            IDS.Data.Hardware.Dispenser.CurrentHeads = 1
            MyDispenserSettings.RevertData()
        End If
        IDS.Data.SaveData()
    End Sub

    Private Sub CheckBoxLifter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxLifter.CheckedChanged
        MyHardwareCommunicationSetup.RefreshDisplay()
    End Sub

    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
        MyConveyorSettings.e_stop_T1_Tick()
    End Sub

    'Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Keys.ControlKey Then
    '        isPress = True
    '    End If
    'End Sub

    'Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Keys.ControlKey Then
    '        isPress = False
    '    End If
    'End Sub

    Private Sub HardwareInitTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HardwareInitTimer.Tick
        HardwareInitTimer.Stop()
        HardwareInitTimer.Enabled = False
        'motion controller start and do homing
        If Not (homingStart) Then
            m_Tri.Connect_Controller()
            m_Tri.SetMachineRun()
            m_Tri.m_TriCtrl.Execute("RUN SETDATUM")
            LabelMessage.Text = "Homing......"
            homingStart = True
            homingDone = False
            HardwareInitTimer.Start()
            HardwareInitTimer.Enabled = True
            homingWaitCnt = 0
        Else
            If Not (homingDone) Then
                homingWaitCnt += 1
                If homingWaitCnt > 350 Then
                    LabelMessage.Text = "Homing timeout(> 35 Seconds) error! Please Check Axes Status!"
                    EnableControls(True)
                Else 'keep waiting for homing done
                    HardwareInitTimer.Start()
                    HardwareInitTimer.Enabled = True
                End If
            Else
                LabelMessage.Text = "System Ready"
                EnableControls(True)
            End If
            End If
    End Sub

    Private Sub EnableControls(ByVal enable As Boolean)
        For Each ctrl As System.Windows.Forms.Control In Me.Controls
            ctrl.Enabled = enable
            m_Tri.SteppingButtons.Enabled = enable
        Next
    End Sub
End Class
