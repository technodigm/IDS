Imports System.IO
Imports DLL_Export_Service

Public Class NeedleCalibrationSettings
    Inherits System.Windows.Forms.Form

    Friend Z_Offset As Double
    Public CalibrationTable(16) As Double
    Friend x_pitch_l, y_pitch_l As Double
    Private RegionX = 768 / 2
    Private RegionY = 576 / 2
    Private ROI_X = 700
    Private ROI_Y = 550
    Private success_num As Integer
    Private MeasuredOffsetX, MeasuredOffsetY As Double

    Friend LeftRoughNeedleCalibrationOffset(2) As Double
    Public NeedleCalibrationState As String = "Stopped"

    Private m_Rough, m_Open, m_Close, m_Threshold, m_Brightness, m_maxRad, m_minRad, m_Compact

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
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents ClearanceHeight As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractHeight As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractDelay As System.Windows.Forms.NumericUpDown
    Friend WithEvents ApproachHeight As System.Windows.Forms.NumericUpDown
    Friend WithEvents NeedleGap As System.Windows.Forms.NumericUpDown
    Friend WithEvents DispenseDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Public WithEvents RightHead As System.Windows.Forms.RadioButton
    Public WithEvents LeftHead As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonCalibrateVision As System.Windows.Forms.Button
    Friend WithEvents BoxStep3 As System.Windows.Forms.GroupBox
    Friend WithEvents Compactness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Roughness As System.Windows.Forms.NumericUpDown
    Friend WithEvents MinRadius As System.Windows.Forms.NumericUpDown
    Friend WithEvents MaxRadius As System.Windows.Forms.NumericUpDown
    Friend WithEvents CloseValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents OpenValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Threshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents Brightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents BlackBackground As System.Windows.Forms.RadioButton
    Friend WithEvents WhiteBackground As System.Windows.Forms.RadioButton
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents BoxStep4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Calibrate As System.Windows.Forms.Button
    Friend WithEvents BoxStep1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblNotice As System.Windows.Forms.Label
    Friend WithEvents BoxStep2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ButtonTest As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents ButtonCalibrateZPosition As System.Windows.Forms.Button
    Friend WithEvents BoxStep1Jetting As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonCalibrateZPositionJetting As System.Windows.Forms.Button
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents btCheckDotOnce As System.Windows.Forms.Button
    Friend WithEvents btSaveDotSize As System.Windows.Forms.Button
    Friend WithEvents btPurgeNow As System.Windows.Forms.Button
    Friend WithEvents tbStatus As System.Windows.Forms.TextBox
    Friend WithEvents btCheckCalibrationAccuracy As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSettings))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label21 = New System.Windows.Forms.Label
        Me.BoxStep2 = New System.Windows.Forms.GroupBox
        Me.btPurgeNow = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.RetractHeight = New System.Windows.Forms.NumericUpDown
        Me.Label31 = New System.Windows.Forms.Label
        Me.RetractSpeed = New System.Windows.Forms.NumericUpDown
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.NeedleGap = New System.Windows.Forms.NumericUpDown
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.ApproachHeight = New System.Windows.Forms.NumericUpDown
        Me.Label25 = New System.Windows.Forms.Label
        Me.RetractDelay = New System.Windows.Forms.NumericUpDown
        Me.ClearanceHeight = New System.Windows.Forms.NumericUpDown
        Me.Label30 = New System.Windows.Forms.Label
        Me.DispenseDuration = New System.Windows.Forms.NumericUpDown
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonCalibrateVision = New System.Windows.Forms.Button
        Me.RightHead = New System.Windows.Forms.RadioButton
        Me.LeftHead = New System.Windows.Forms.RadioButton
        Me.Label11 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.BoxStep3 = New System.Windows.Forms.GroupBox
        Me.tbStatus = New System.Windows.Forms.TextBox
        Me.btSaveDotSize = New System.Windows.Forms.Button
        Me.btCheckDotOnce = New System.Windows.Forms.Button
        Me.Compactness = New System.Windows.Forms.NumericUpDown
        Me.Roughness = New System.Windows.Forms.NumericUpDown
        Me.MinRadius = New System.Windows.Forms.NumericUpDown
        Me.MaxRadius = New System.Windows.Forms.NumericUpDown
        Me.CloseValue = New System.Windows.Forms.NumericUpDown
        Me.OpenValue = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Threshold = New System.Windows.Forms.NumericUpDown
        Me.Brightness = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.BlackBackground = New System.Windows.Forms.RadioButton
        Me.WhiteBackground = New System.Windows.Forms.RadioButton
        Me.Label19 = New System.Windows.Forms.Label
        Me.ButtonTest = New System.Windows.Forms.Button
        Me.btCheckCalibrationAccuracy = New System.Windows.Forms.Button
        Me.BoxStep4 = New System.Windows.Forms.GroupBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.Calibrate = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.BoxStep1 = New System.Windows.Forms.GroupBox
        Me.ButtonCalibrateZPosition = New System.Windows.Forms.Button
        Me.lblNotice = New System.Windows.Forms.Label
        Me.BoxStep1Jetting = New System.Windows.Forms.GroupBox
        Me.ButtonCalibrateZPositionJetting = New System.Windows.Forms.Button
        Me.Label22 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PanelToBeAdded.SuspendLayout()
        Me.BoxStep2.SuspendLayout()
        CType(Me.RetractHeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NeedleGap, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ApproachHeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ClearanceHeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DispenseDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BoxStep3.SuspendLayout()
        CType(Me.Compactness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Roughness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MinRadius, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxRadius, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CloseValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.OpenValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Threshold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Brightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BoxStep4.SuspendLayout()
        Me.BoxStep1.SuspendLayout()
        Me.BoxStep1Jetting.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.Label21)
        Me.PanelToBeAdded.Controls.Add(Me.BoxStep2)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.BoxStep3)
        Me.PanelToBeAdded.Controls.Add(Me.BoxStep4)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.BoxStep1)
        Me.PanelToBeAdded.Controls.Add(Me.BoxStep1Jetting)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(24, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 68
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label21.Location = New System.Drawing.Point(0, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(296, 32)
        Me.Label21.TabIndex = 90
        Me.Label21.Text = "Needle Calibration Settings"
        '
        'BoxStep2
        '
        Me.BoxStep2.Controls.Add(Me.btPurgeNow)
        Me.BoxStep2.Controls.Add(Me.Label17)
        Me.BoxStep2.Controls.Add(Me.RetractHeight)
        Me.BoxStep2.Controls.Add(Me.Label31)
        Me.BoxStep2.Controls.Add(Me.RetractSpeed)
        Me.BoxStep2.Controls.Add(Me.Label29)
        Me.BoxStep2.Controls.Add(Me.Label26)
        Me.BoxStep2.Controls.Add(Me.NeedleGap)
        Me.BoxStep2.Controls.Add(Me.Label13)
        Me.BoxStep2.Controls.Add(Me.Label14)
        Me.BoxStep2.Controls.Add(Me.Label27)
        Me.BoxStep2.Controls.Add(Me.Label32)
        Me.BoxStep2.Controls.Add(Me.ApproachHeight)
        Me.BoxStep2.Controls.Add(Me.Label25)
        Me.BoxStep2.Controls.Add(Me.RetractDelay)
        Me.BoxStep2.Controls.Add(Me.ClearanceHeight)
        Me.BoxStep2.Controls.Add(Me.Label30)
        Me.BoxStep2.Controls.Add(Me.DispenseDuration)
        Me.BoxStep2.Controls.Add(Me.Label12)
        Me.BoxStep2.Controls.Add(Me.Label15)
        Me.BoxStep2.Controls.Add(Me.Label16)
        Me.BoxStep2.Controls.Add(Me.Label18)
        Me.BoxStep2.Controls.Add(Me.Label1)
        Me.BoxStep2.Controls.Add(Me.ButtonCalibrateVision)
        Me.BoxStep2.Controls.Add(Me.RightHead)
        Me.BoxStep2.Controls.Add(Me.LeftHead)
        Me.BoxStep2.Controls.Add(Me.Label11)
        Me.BoxStep2.Location = New System.Drawing.Point(8, 176)
        Me.BoxStep2.Name = "BoxStep2"
        Me.BoxStep2.Size = New System.Drawing.Size(496, 272)
        Me.BoxStep2.TabIndex = 89
        Me.BoxStep2.TabStop = False
        Me.BoxStep2.Text = "Step 2: Single Dot Dispense"
        '
        'btPurgeNow
        '
        Me.btPurgeNow.Location = New System.Drawing.Point(360, 192)
        Me.btPurgeNow.Name = "btPurgeNow"
        Me.btPurgeNow.Size = New System.Drawing.Size(128, 64)
        Me.btPurgeNow.TabIndex = 92
        Me.btPurgeNow.Text = "Purge At Current Position"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(256, 136)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(104, 24)
        Me.Label17.TabIndex = 57
        Me.Label17.Text = "Needle Gap"
        '
        'RetractHeight
        '
        Me.RetractHeight.DecimalPlaces = 3
        Me.RetractHeight.Location = New System.Drawing.Point(376, 104)
        Me.RetractHeight.Name = "RetractHeight"
        Me.RetractHeight.Size = New System.Drawing.Size(80, 27)
        Me.RetractHeight.TabIndex = 66
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.Location = New System.Drawing.Point(456, 136)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(40, 23)
        Me.Label31.TabIndex = 82
        Me.Label31.Text = "mm"
        '
        'RetractSpeed
        '
        Me.RetractSpeed.Location = New System.Drawing.Point(152, 168)
        Me.RetractSpeed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.RetractSpeed.Name = "RetractSpeed"
        Me.RetractSpeed.Size = New System.Drawing.Size(64, 27)
        Me.RetractSpeed.TabIndex = 64
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(456, 104)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(40, 23)
        Me.Label29.TabIndex = 84
        Me.Label29.Text = "mm"
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.Location = New System.Drawing.Point(216, 168)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(56, 23)
        Me.Label26.TabIndex = 87
        Me.Label26.Text = "mm/s"
        '
        'NeedleGap
        '
        Me.NeedleGap.DecimalPlaces = 3
        Me.NeedleGap.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NeedleGap.Location = New System.Drawing.Point(376, 136)
        Me.NeedleGap.Name = "NeedleGap"
        Me.NeedleGap.Size = New System.Drawing.Size(80, 27)
        Me.NeedleGap.TabIndex = 58
        Me.NeedleGap.Value = New Decimal(New Integer() {5, 0, 0, 65536})
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(256, 104)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(120, 24)
        Me.Label13.TabIndex = 65
        Me.Label13.Text = "Retract Height"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 168)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 23)
        Me.Label14.TabIndex = 63
        Me.Label14.Text = "Retract Speed"
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(216, 136)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(40, 23)
        Me.Label27.TabIndex = 86
        Me.Label27.Text = "mm"
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(216, 72)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(32, 23)
        Me.Label32.TabIndex = 81
        Me.Label32.Text = "ms"
        '
        'ApproachHeight
        '
        Me.ApproachHeight.DecimalPlaces = 3
        Me.ApproachHeight.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ApproachHeight.Location = New System.Drawing.Point(152, 104)
        Me.ApproachHeight.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.ApproachHeight.Name = "ApproachHeight"
        Me.ApproachHeight.Size = New System.Drawing.Size(64, 27)
        Me.ApproachHeight.TabIndex = 60
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label25.Location = New System.Drawing.Point(456, 72)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(32, 23)
        Me.Label25.TabIndex = 88
        Me.Label25.Text = "ms"
        '
        'RetractDelay
        '
        Me.RetractDelay.Location = New System.Drawing.Point(376, 72)
        Me.RetractDelay.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.RetractDelay.Name = "RetractDelay"
        Me.RetractDelay.Size = New System.Drawing.Size(80, 27)
        Me.RetractDelay.TabIndex = 62
        '
        'ClearanceHeight
        '
        Me.ClearanceHeight.DecimalPlaces = 3
        Me.ClearanceHeight.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ClearanceHeight.Location = New System.Drawing.Point(152, 136)
        Me.ClearanceHeight.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.ClearanceHeight.Name = "ClearanceHeight"
        Me.ClearanceHeight.Size = New System.Drawing.Size(64, 27)
        Me.ClearanceHeight.TabIndex = 68
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(216, 104)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(40, 23)
        Me.Label30.TabIndex = 83
        Me.Label30.Text = "mm"
        '
        'DispenseDuration
        '
        Me.DispenseDuration.Location = New System.Drawing.Point(152, 72)
        Me.DispenseDuration.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.DispenseDuration.Name = "DispenseDuration"
        Me.DispenseDuration.Size = New System.Drawing.Size(64, 27)
        Me.DispenseDuration.TabIndex = 56
        Me.DispenseDuration.Value = New Decimal(New Integer() {1000, 0, 0, 0})
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 136)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(144, 23)
        Me.Label12.TabIndex = 67
        Me.Label12.Text = "Clearance Height"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(256, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(120, 23)
        Me.Label15.TabIndex = 61
        Me.Label15.Text = "Retract Delay"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(8, 104)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(144, 23)
        Me.Label16.TabIndex = 59
        Me.Label16.Text = "Approach Height"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(8, 72)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(128, 23)
        Me.Label18.TabIndex = 55
        Me.Label18.Text = "Dispense Time"
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.Brown
        Me.Label1.Location = New System.Drawing.Point(8, 216)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 48)
        Me.Label1.TabIndex = 91
        Me.Label1.Text = "Clean the calibration block before pressing."
        '
        'ButtonCalibrateVision
        '
        Me.ButtonCalibrateVision.Location = New System.Drawing.Point(224, 216)
        Me.ButtonCalibrateVision.Name = "ButtonCalibrateVision"
        Me.ButtonCalibrateVision.Size = New System.Drawing.Size(128, 40)
        Me.ButtonCalibrateVision.TabIndex = 61
        Me.ButtonCalibrateVision.Text = "Dispense Dot"
        '
        'RightHead
        '
        Me.RightHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RightHead.Location = New System.Drawing.Point(320, 32)
        Me.RightHead.Name = "RightHead"
        Me.RightHead.Size = New System.Drawing.Size(72, 24)
        Me.RightHead.TabIndex = 3
        Me.RightHead.Text = "Right"
        '
        'LeftHead
        '
        Me.LeftHead.Checked = True
        Me.LeftHead.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LeftHead.Location = New System.Drawing.Point(240, 32)
        Me.LeftHead.Name = "LeftHead"
        Me.LeftHead.Size = New System.Drawing.Size(64, 24)
        Me.LeftHead.TabIndex = 2
        Me.LeftHead.TabStop = True
        Me.LeftHead.Text = "Left"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(104, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(120, 24)
        Me.Label11.TabIndex = 3
        Me.Label11.Text = "Current Head"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(392, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 73
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'BoxStep3
        '
        Me.BoxStep3.Controls.Add(Me.tbStatus)
        Me.BoxStep3.Controls.Add(Me.btSaveDotSize)
        Me.BoxStep3.Controls.Add(Me.btCheckDotOnce)
        Me.BoxStep3.Controls.Add(Me.Compactness)
        Me.BoxStep3.Controls.Add(Me.Roughness)
        Me.BoxStep3.Controls.Add(Me.MinRadius)
        Me.BoxStep3.Controls.Add(Me.MaxRadius)
        Me.BoxStep3.Controls.Add(Me.CloseValue)
        Me.BoxStep3.Controls.Add(Me.OpenValue)
        Me.BoxStep3.Controls.Add(Me.Label2)
        Me.BoxStep3.Controls.Add(Me.Label3)
        Me.BoxStep3.Controls.Add(Me.Label8)
        Me.BoxStep3.Controls.Add(Me.Label4)
        Me.BoxStep3.Controls.Add(Me.Label5)
        Me.BoxStep3.Controls.Add(Me.Label6)
        Me.BoxStep3.Controls.Add(Me.Threshold)
        Me.BoxStep3.Controls.Add(Me.Brightness)
        Me.BoxStep3.Controls.Add(Me.Label7)
        Me.BoxStep3.Controls.Add(Me.Label9)
        Me.BoxStep3.Controls.Add(Me.BlackBackground)
        Me.BoxStep3.Controls.Add(Me.WhiteBackground)
        Me.BoxStep3.Controls.Add(Me.Label19)
        Me.BoxStep3.Controls.Add(Me.ButtonTest)
        Me.BoxStep3.Controls.Add(Me.btCheckCalibrationAccuracy)
        Me.BoxStep3.Location = New System.Drawing.Point(8, 456)
        Me.BoxStep3.Name = "BoxStep3"
        Me.BoxStep3.Size = New System.Drawing.Size(496, 296)
        Me.BoxStep3.TabIndex = 4
        Me.BoxStep3.TabStop = False
        Me.BoxStep3.Text = "Step 3: Check Dot"
        '
        'tbStatus
        '
        Me.tbStatus.Location = New System.Drawing.Point(16, 256)
        Me.tbStatus.Name = "tbStatus"
        Me.tbStatus.ReadOnly = True
        Me.tbStatus.Size = New System.Drawing.Size(464, 27)
        Me.tbStatus.TabIndex = 108
        Me.tbStatus.Text = ""
        '
        'btSaveDotSize
        '
        Me.btSaveDotSize.Location = New System.Drawing.Point(424, 200)
        Me.btSaveDotSize.Name = "btSaveDotSize"
        Me.btSaveDotSize.Size = New System.Drawing.Size(88, 48)
        Me.btSaveDotSize.TabIndex = 107
        Me.btSaveDotSize.Text = "Save Dot Size"
        Me.btSaveDotSize.Visible = False
        '
        'btCheckDotOnce
        '
        Me.btCheckDotOnce.Location = New System.Drawing.Point(80, 200)
        Me.btCheckDotOnce.Name = "btCheckDotOnce"
        Me.btCheckDotOnce.Size = New System.Drawing.Size(160, 48)
        Me.btCheckDotOnce.TabIndex = 106
        Me.btCheckDotOnce.Text = "Check Once"
        '
        'Compactness
        '
        Me.Compactness.DecimalPlaces = 1
        Me.Compactness.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.Compactness.Location = New System.Drawing.Point(355, 168)
        Me.Compactness.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.Compactness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.Compactness.Name = "Compactness"
        Me.Compactness.Size = New System.Drawing.Size(64, 27)
        Me.Compactness.TabIndex = 99
        Me.Compactness.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Roughness
        '
        Me.Roughness.DecimalPlaces = 1
        Me.Roughness.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.Roughness.Location = New System.Drawing.Point(355, 136)
        Me.Roughness.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.Roughness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.Roughness.Name = "Roughness"
        Me.Roughness.Size = New System.Drawing.Size(64, 27)
        Me.Roughness.TabIndex = 100
        Me.Roughness.Value = New Decimal(New Integer() {3, 0, 0, 0})
        '
        'MinRadius
        '
        Me.MinRadius.DecimalPlaces = 2
        Me.MinRadius.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.MinRadius.Location = New System.Drawing.Point(355, 104)
        Me.MinRadius.Maximum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.MinRadius.Name = "MinRadius"
        Me.MinRadius.Size = New System.Drawing.Size(64, 27)
        Me.MinRadius.TabIndex = 97
        Me.MinRadius.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'MaxRadius
        '
        Me.MaxRadius.DecimalPlaces = 2
        Me.MaxRadius.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.MaxRadius.Location = New System.Drawing.Point(355, 72)
        Me.MaxRadius.Maximum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.MaxRadius.Name = "MaxRadius"
        Me.MaxRadius.Size = New System.Drawing.Size(64, 27)
        Me.MaxRadius.TabIndex = 104
        '
        'CloseValue
        '
        Me.CloseValue.Location = New System.Drawing.Point(123, 168)
        Me.CloseValue.Maximum = New Decimal(New Integer() {9, 0, 0, 0})
        Me.CloseValue.Name = "CloseValue"
        Me.CloseValue.Size = New System.Drawing.Size(64, 27)
        Me.CloseValue.TabIndex = 95
        '
        'OpenValue
        '
        Me.OpenValue.Location = New System.Drawing.Point(123, 136)
        Me.OpenValue.Maximum = New Decimal(New Integer() {9, 0, 0, 0})
        Me.OpenValue.Name = "OpenValue"
        Me.OpenValue.Size = New System.Drawing.Size(64, 27)
        Me.OpenValue.TabIndex = 96
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(243, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 94
        Me.Label2.Text = "Compactness"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(243, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 23)
        Me.Label3.TabIndex = 93
        Me.Label3.Text = "Roughness"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(224, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(120, 23)
        Me.Label8.TabIndex = 92
        Me.Label8.Text = "Min Diameter"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(224, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 23)
        Me.Label4.TabIndex = 91
        Me.Label4.Text = "Max Diameter"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(27, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 23)
        Me.Label5.TabIndex = 90
        Me.Label5.Text = "Open"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(27, 168)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 23)
        Me.Label6.TabIndex = 89
        Me.Label6.Text = "Close"
        '
        'Threshold
        '
        Me.Threshold.Location = New System.Drawing.Point(123, 104)
        Me.Threshold.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.Threshold.Name = "Threshold"
        Me.Threshold.Size = New System.Drawing.Size(64, 27)
        Me.Threshold.TabIndex = 105
        '
        'Brightness
        '
        Me.Brightness.Location = New System.Drawing.Point(123, 72)
        Me.Brightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.Brightness.Name = "Brightness"
        Me.Brightness.Size = New System.Drawing.Size(64, 27)
        Me.Brightness.TabIndex = 53
        Me.Brightness.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(27, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 24)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Brightness"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(27, 104)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(88, 23)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Threshold"
        '
        'BlackBackground
        '
        Me.BlackBackground.Checked = True
        Me.BlackBackground.Location = New System.Drawing.Point(235, 32)
        Me.BlackBackground.Name = "BlackBackground"
        Me.BlackBackground.Size = New System.Drawing.Size(76, 24)
        Me.BlackBackground.TabIndex = 0
        Me.BlackBackground.TabStop = True
        Me.BlackBackground.Text = "Black"
        '
        'WhiteBackground
        '
        Me.WhiteBackground.Location = New System.Drawing.Point(323, 32)
        Me.WhiteBackground.Name = "WhiteBackground"
        Me.WhiteBackground.Size = New System.Drawing.Size(76, 24)
        Me.WhiteBackground.TabIndex = 1
        Me.WhiteBackground.Text = "White"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(136, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(88, 24)
        Me.Label19.TabIndex = 3
        Me.Label19.Text = "Dot Color"
        '
        'ButtonTest
        '
        Me.ButtonTest.Location = New System.Drawing.Point(256, 200)
        Me.ButtonTest.Name = "ButtonTest"
        Me.ButtonTest.Size = New System.Drawing.Size(160, 48)
        Me.ButtonTest.TabIndex = 61
        Me.ButtonTest.Text = "Continous Check Dot"
        '
        'btCheckCalibrationAccuracy
        '
        Me.btCheckCalibrationAccuracy.Location = New System.Drawing.Point(392, 16)
        Me.btCheckCalibrationAccuracy.Name = "btCheckCalibrationAccuracy"
        Me.btCheckCalibrationAccuracy.Size = New System.Drawing.Size(96, 48)
        Me.btCheckCalibrationAccuracy.TabIndex = 93
        Me.btCheckCalibrationAccuracy.Text = "Check Accuracy"
        '
        'BoxStep4
        '
        Me.BoxStep4.Controls.Add(Me.Label20)
        Me.BoxStep4.Controls.Add(Me.Calibrate)
        Me.BoxStep4.Location = New System.Drawing.Point(8, 808)
        Me.BoxStep4.Name = "BoxStep4"
        Me.BoxStep4.Size = New System.Drawing.Size(496, 88)
        Me.BoxStep4.TabIndex = 4
        Me.BoxStep4.TabStop = False
        Me.BoxStep4.Text = "Step 4: Calibrate Needle to Camera Offset"
        '
        'Label20
        '
        Me.Label20.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label20.Location = New System.Drawing.Point(32, 32)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(176, 48)
        Me.Label20.TabIndex = 91
        Me.Label20.Text = "Final calibration with dispensing of 8 dots."
        '
        'Calibrate
        '
        Me.Calibrate.Location = New System.Drawing.Point(216, 32)
        Me.Calibrate.Name = "Calibrate"
        Me.Calibrate.Size = New System.Drawing.Size(88, 40)
        Me.Calibrate.TabIndex = 9
        Me.Calibrate.Text = "Calibrate"
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(219, 760)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 40)
        Me.ButtonSave.TabIndex = 72
        Me.ButtonSave.Text = "Save"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Location = New System.Drawing.Point(384, 760)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(75, 40)
        Me.ButtonRevert.TabIndex = 71
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'BoxStep1
        '
        Me.BoxStep1.Controls.Add(Me.ButtonCalibrateZPosition)
        Me.BoxStep1.Controls.Add(Me.lblNotice)
        Me.BoxStep1.Cursor = System.Windows.Forms.Cursors.Default
        Me.BoxStep1.Location = New System.Drawing.Point(8, 64)
        Me.BoxStep1.Name = "BoxStep1"
        Me.BoxStep1.Size = New System.Drawing.Size(496, 104)
        Me.BoxStep1.TabIndex = 0
        Me.BoxStep1.TabStop = False
        Me.BoxStep1.Text = "Step 1: Set Needle Datum"
        '
        'ButtonCalibrateZPosition
        '
        Me.ButtonCalibrateZPosition.Location = New System.Drawing.Point(272, 40)
        Me.ButtonCalibrateZPosition.Name = "ButtonCalibrateZPosition"
        Me.ButtonCalibrateZPosition.Size = New System.Drawing.Size(176, 40)
        Me.ButtonCalibrateZPosition.TabIndex = 61
        Me.ButtonCalibrateZPosition.Text = "Calibrate Position"
        '
        'lblNotice
        '
        Me.lblNotice.ForeColor = System.Drawing.Color.Brown
        Me.lblNotice.Location = New System.Drawing.Point(32, 24)
        Me.lblNotice.Name = "lblNotice"
        Me.lblNotice.Size = New System.Drawing.Size(216, 64)
        Me.lblNotice.TabIndex = 91
        Me.lblNotice.Text = "Move the needle tip to above the center of touch sensor before pressing."
        '
        'BoxStep1Jetting
        '
        Me.BoxStep1Jetting.Controls.Add(Me.ButtonCalibrateZPositionJetting)
        Me.BoxStep1Jetting.Controls.Add(Me.Label22)
        Me.BoxStep1Jetting.Cursor = System.Windows.Forms.Cursors.Default
        Me.BoxStep1Jetting.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BoxStep1Jetting.Location = New System.Drawing.Point(8, 64)
        Me.BoxStep1Jetting.Name = "BoxStep1Jetting"
        Me.BoxStep1Jetting.Size = New System.Drawing.Size(496, 104)
        Me.BoxStep1Jetting.TabIndex = 0
        Me.BoxStep1Jetting.TabStop = False
        Me.BoxStep1Jetting.Text = "Step 1"
        '
        'ButtonCalibrateZPositionJetting
        '
        Me.ButtonCalibrateZPositionJetting.Location = New System.Drawing.Point(272, 40)
        Me.ButtonCalibrateZPositionJetting.Name = "ButtonCalibrateZPositionJetting"
        Me.ButtonCalibrateZPositionJetting.Size = New System.Drawing.Size(176, 40)
        Me.ButtonCalibrateZPositionJetting.TabIndex = 61
        Me.ButtonCalibrateZPositionJetting.Text = "Set Z Position"
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(32, 32)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(216, 64)
        Me.Label22.TabIndex = 91
        Me.Label22.Text = "Set the Z position of the Jetting Valve for dispensing dots."
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "(*.txt)*|"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.Filter = "(*.txt)|*"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'NeedleCalibrationSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSettings"
        Me.Text = "NeedleCalibration1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.BoxStep2.ResumeLayout(False)
        CType(Me.RetractHeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NeedleGap, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ApproachHeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ClearanceHeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DispenseDuration, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BoxStep3.ResumeLayout(False)
        CType(Me.Compactness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Roughness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MinRadius, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxRadius, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CloseValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.OpenValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Threshold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Brightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BoxStep4.ResumeLayout(False)
        Me.BoxStep1.ResumeLayout(False)
        Me.BoxStep1Jetting.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Calibration Procedures"

    Private Function CheckLocalDot()
        Return IDS.Devices.Vision.IDSV_NC(WhiteBackground.Checked, Threshold.Value, MaxRadius.Value / 2, MinRadius.Value / 2, CloseValue.Value, OpenValue.Value, Roughness.Value, Compactness.Value, RegionX, RegionY, ROI_X, ROI_Y, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y)
    End Function

    Private Function CheckDot()
        MySleep(500)
        Vision.FrmVision.WaitForImageToStabilize()
        With IDS.Data.Hardware.Needle.Left
            Dim rtn As Boolean
            rtn = Vision.IDSV_NC(.CalBackground, .CalThreshold, .CalMaxRadius, .CalMinRadius, .CalClose, .CalOpen, .CalRoughness, .CalCompactness, RegionX, RegionY, ROI_X, ROI_Y, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y)
            Vision.FrmVision.ClearDisplay()
            If rtn Then
                tbStatus.Text = Vision.FrmVision.diameter
                tbStatus.Refresh()
            End If
            Return rtn
        End With
    End Function

    Public Sub SetNeedleXYCal(ByVal NumOfDots As Integer, ByVal DispenseIO As Integer)

        With IDS.Data.Hardware.Needle
            LeftRoughNeedleCalibrationOffset(0) = .Left.NeedleCalibrationPosition.X - .Left.CalibratorPos.X
            LeftRoughNeedleCalibrationOffset(1) = .Left.NeedleCalibrationPosition.Y - .Left.CalibratorPos.Y
            x_pitch_l = (.Left.ArrayDotPos3.X - .Left.ArrayDotPos1.X) / 2
            y_pitch_l = (.Left.ArrayDotPos3.Y - .Left.ArrayDotPos1.Y) / 2
            CalibrationTable(0) = .Left.DispenseDot.DispenseDuration '110
            CalibrationTable(1) = .Left.DispenseDot.NeedleGap
            CalibrationTable(2) = .Left.DispenseDot.ApproachHeight
            CalibrationTable(3) = .Left.DispenseDot.RetractDelay
            CalibrationTable(4) = .Left.DispenseDot.RetractSpeed
            CalibrationTable(5) = .Left.DispenseDot.RetractHeight '115
            CalibrationTable(6) = .Left.DispenseDot.ClearanceHeight
            CalibrationTable(7) = IDS.Data.Hardware.Gantry.ServiceXYSpeed  'xy speed
            CalibrationTable(8) = IDS.Data.Hardware.Gantry.ServiceZSpeed  'z speed
            CalibrationTable(9) = DispenseIO
            CalibrationTable(10) = NumOfDots 'dot number: 8 or 1
            CalibrationTable(11) = .Left.ArrayDotPos1.X + LeftRoughNeedleCalibrationOffset(0)
            CalibrationTable(12) = .Left.ArrayDotPos1.Y + LeftRoughNeedleCalibrationOffset(1)
            CalibrationTable(13) = Z_Offset
            CalibrationTable(14) = x_pitch_l
            CalibrationTable(15) = y_pitch_l '125
            CalibrationTable(16) = 0 '125
        End With

    End Sub

    Public Function CalibrateNeedleZPosition(Optional ByVal moveToZ As Boolean = True) As Boolean
        'call the z calibration sub routine in the program CALIBRATIONS inside Motion Perfect 
        m_Tri.SetMachineRun()
        m_Tri.SetCalibrationType("Needle Z Calibration")
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not TrioMotionCalibrating() Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        'check for failure
        If m_Tri.GetCalibrationFlag = 2 Then
            m_Tri.ResetCalibrationFlag()
            'MsgBox("Nothing sensed.")
            m_Tri.SetMachineStop()
            Return False
        End If

        'record the x,y,z coordinates at which the needle tip strikes the touch sensor and save it.
        m_Tri.GetIDSState()
        IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.X = m_Tri.XPosition
        IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Y = m_Tri.YPosition
        IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z = m_Tri.ZPosition + IDS.Data.Hardware.Needle.Left.CalibratorPos.Z - IDS.Data.Hardware.Needle.Left.TouchSensorZPosition
        IDS.Data.SaveData()

        'move to 0 when done.
        SetServiceSpeed()
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        m_Tri.SetMachineStop()
        Return True

StopCalibration:
        'MsgBox("Checking the Z position for needle calibration stopped prematurely.")
        m_Tri.SetMachineStop()

        SetServiceSpeed()
        m_Tri.Move_Z(0)
       
        Calibrate.Enabled = True
        Calibrate.Text = "Calibrate"
        'If localCalib Then
        '    MessageBox.Show("Needle Calibration stopped.")
        '    localCalib = False
        'End If
        Return False
    End Function

    Public Function CalibrateDispenseDotForVision() As Boolean

        'set the flag to dispense the middle dot
        '1 - one mid dot,  8 - eight XYCal dots
        SetNeedleXYCal(1, 25)

        'gantry moves to position laser at calibration block
        'position(0) = -5     'position has to be found by laser-camera offset, etc...
        'position(1) = -393   'white area
        'position(0) = 77     'position has to be found by laser-camera offset, etc...
        'position(1) = -361   'black area
        position(0) = CalibrationTable(11) + CalibrationTable(14) + LeftNeedleOffsetX - LaserOffX
        position(1) = CalibrationTable(12) + CalibrationTable(15) + LeftNeedleOffsetY - LaserOffY
        m_Tri.SetMachineRun()
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration

        'laser checks the height of the disp. plane
        MySleep(50)
        If Not Laser.WaitForReadingToStabilize() Then GoTo stopcalibration
        Z_Offset = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z + IDS.Data.Hardware.HeightSensor.Laser.CurrentPos.Z - Laser.LASER_Reading

        SetNeedleXYCal(1, 25)
        m_Tri.m_TriCtrl.SetTable(110, 17, CalibrationTable)

        'download dispenser settings
        If LeftHead.Checked Then
            MyDispenserSettings.DownloadDispenserSettings("Left")
        ElseIf RightHead.Checked Then
            MyDispenserSettings.DownloadDispenserSettings("Right")
        End If

        'dispense a dot and move back to 0 when done
        m_Tri.SetCalibrationType("Needle XY Calibration")
        If Not TrioMotionCalibrating() Then GoTo StopCalibration
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration

        'move the camera to the dot
        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X + x_pitch_l
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y + y_pitch_l

        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        m_Tri.SetMachineStop()
        Return True

StopCalibration:
        'MsgBox("Dispensing the dot for needle calibration stopped prematurely.")
        m_Tri.SetMachineStop()
        Return False
    End Function

    Public Function CalibrateNeedleXYPosition(Optional ByVal moveToZ As Boolean = True) As Boolean

        'gantry moves to position laser at calibration block
        'position(0) = -5     'position has to be found by laser-camera offset, etc...
        'position(1) = -393   'white area
        'position(0) = 77     'position has to be found by laser-camera offset, etc...
        'position(1) = -361   'black area
        SetNeedleXYCal(1, 25)
        position(0) = CalibrationTable(11) + CalibrationTable(14) + LeftNeedleOffsetX - LaserOffX
        position(1) = CalibrationTable(12) + CalibrationTable(15) + LeftNeedleOffsetY - LaserOffY
        m_Tri.SetMachineRun()
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration

        'laser checks the height of the disp. plane
        MySleep(50)
        OnLaser()
        If Not Laser.WaitForReadingToStabilize() Then GoTo stopcalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        Z_Offset = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z + IDS.Data.Hardware.HeightSensor.Laser.CurrentPos.Z - Laser.LASER_Reading
        OffLaser()

        'download dispenser settings
        Try
            If LeftHead.Checked Then
                MyDispenserSettings.DownloadDispenserSettings("Left")
            ElseIf RightHead.Checked Then
                MyDispenserSettings.DownloadDispenserSettings("Right")
            End If
        Catch ex As Exception
            MessageBox.Show("Error occur when downloading dispenser settings. Make sure devices is all connected.")
            GoTo StopCalibration
        End Try
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        'xy dispensing
        SetNeedleXYCal(8, 25)
        m_Tri.m_TriCtrl.SetTable(110, 17, CalibrationTable)
        m_Tri.SetCalibrationType("Needle XY Calibration")
        If Not TrioMotionCalibrating() Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration

        MeasuredOffsetX = 0
        MeasuredOffsetY = 0
        Vision.FrmVision.SwitchCamera("Teach Camera")
        Vision.FrmVision.DisplayIndicator()
        IDS.Devices.Vision.IDSV_SetBrightness(IDS.Data.Hardware.Camera.Brightness)
        'vision checking
        IDS.Devices.Vision.IDSV_SetBrightness(IDS.Data.Hardware.Needle.Left.CalBrightness)
        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X + x_pitch_l
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y + y_pitch_l
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X + x_pitch_l
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos3.Y
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        distance(0) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X
        distance(1) = IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y + y_pitch_l
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If Not m_Tri.Move_XY(distance) Then GoTo StopCalibration
        If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
        If CheckDot() Then
            success_num += 1
            MeasuredOffsetX += MyNeedleCalibrationSetup1.Offset_X
            MeasuredOffsetY += MyNeedleCalibrationSetup1.Offset_Y
        End If

        Vision.FrmVision.DisplayIndicator()
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)

        If success_num > 0 Then
            If (MeasuredOffsetX / success_num) ^ 2 <= 64 And (MeasuredOffsetY / success_num) ^ 2 <= 36 Then
                IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.X = -MeasuredOffsetX / success_num + LeftRoughNeedleCalibrationOffset(0) + IDS.Data.Hardware.Needle.Left.CalibratorPos.X
                IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Y = MeasuredOffsetY / success_num + LeftRoughNeedleCalibrationOffset(1) + IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
                IDS.Data.SaveData()
                MessageBox.Show(success_num & " Dots detected.Needle to camera calibration successful and data saved!")
                success_num = 0
                m_Tri.SetMachineStop()
                Return True
            Else
                MessageBox.Show("Offset is too big. Erroneous reading. Needle calibration fail.")
                success_num = 0
                m_Tri.SetMachineStop()
                Return False
            End If
        Else
            MessageBox.Show("No dots detected. Needle calibration fail.")
            m_Tri.SetMachineStop()
            Return False
        End If

StopCalibration:
        success_num = 0
        OffLaser()
        'MsgBox("Needle calibration process stopped prematurely.")
        m_Tri.SetMachineStop()

        m_Tri.Move_Z(0)

        Calibrate.Enabled = True
        Calibrate.Text = "Calibrate"
        'If localCalib Then
        '    MessageBox.Show("Needle Calibration stopped.")
        '    localCalib = False
        'End If

        Return False

    End Function
#End Region


    Public Sub SaveData()

        With IDS.Data.Hardware.Needle
            If LeftHead.Checked Then
                .Left.DispenseDot.DispenseDuration = DispenseDuration.Value
                .Left.DispenseDot.NeedleGap = NeedleGap.Value
                .Left.DispenseDot.ApproachHeight = ApproachHeight.Value
                .Left.DispenseDot.RetractDelay = RetractDelay.Value
                .Left.DispenseDot.RetractSpeed = RetractSpeed.Value
                .Left.DispenseDot.RetractHeight = RetractHeight.Value
                .Left.DispenseDot.ClearanceHeight = ClearanceHeight.Value
                .Left.CalBackground = BlackBackground.Checked
                .Left.CalBrightness = Brightness.Value
                .Left.CalCompactness = Compactness.Value
                .Left.CalClose = CloseValue.Value
                .Left.CalOpen = OpenValue.Value
                .Left.CalThreshold = Threshold.Value
                .Left.CalRoughness = Roughness.Value
                .Left.CalMinRadius = MinRadius.Value / 2
                .Left.CalMaxRadius = MaxRadius.Value / 2
            ElseIf RightHead.Checked Then
                .Right.DispenseDot.DispenseDuration = DispenseDuration.Value
                .Right.DispenseDot.NeedleGap = NeedleGap.Value
                .Right.DispenseDot.ApproachHeight = ApproachHeight.Value
                .Right.DispenseDot.RetractDelay = RetractDelay.Value
                .Right.DispenseDot.RetractSpeed = RetractSpeed.Value
                .Right.DispenseDot.RetractHeight = RetractHeight.Value
                .Right.DispenseDot.ClearanceHeight = ClearanceHeight.Value
            End If
        End With

        IDS.Data.SaveData()
    End Sub

    Public Sub RevertData(Optional ByVal hideexit As Boolean = False)
        ButtonExit.Visible = Not hideexit
        IDS.Data.OpenData()
        With IDS.Data.Hardware.Needle
            If LeftHead.Checked Then
                DispenseDuration.Text = .Left.DispenseDot.DispenseDuration
                NeedleGap.Text = .Left.DispenseDot.NeedleGap
                ApproachHeight.Text = .Left.DispenseDot.ApproachHeight
                RetractDelay.Text = .Left.DispenseDot.RetractDelay
                RetractSpeed.Text = .Left.DispenseDot.RetractSpeed
                ClearanceHeight.Text = .Left.DispenseDot.ClearanceHeight
                RetractHeight.Text = .Left.DispenseDot.RetractHeight
                BlackBackground.Checked = .Left.CalBackground
                If Not BlackBackground.Checked Then
                    Me.WhiteBackground.Checked = True
                End If
                Brightness.Text = .Left.CalBrightness
                Me.m_Brightness = .Left.CalBrightness
                Compactness.Text = .Left.CalCompactness
                m_Compact = .Left.CalCompactness
                CloseValue.Text = .Left.CalClose
                m_Close = .Left.CalClose
                OpenValue.Text = .Left.CalOpen
                m_Open = .Left.CalOpen
                Threshold.Text = .Left.CalThreshold
                m_Threshold = .Left.CalThreshold
                Roughness.Text = .Left.CalRoughness
                m_Rough = .Left.CalRoughness
                MinRadius.Text = .Left.CalMinRadius * 2
                Me.m_minRad = .Left.CalMinRadius * 2
                MaxRadius.Text = .Left.CalMaxRadius * 2
                Me.m_maxRad = .Left.CalMaxRadius * 2
            ElseIf RightHead.Checked Then
                DispenseDuration.Text = .Right.DispenseDot.DispenseDuration
                NeedleGap.Text = .Right.DispenseDot.NeedleGap
                ApproachHeight.Text = .Right.DispenseDot.ApproachHeight
                RetractDelay.Text = .Right.DispenseDot.RetractDelay
                RetractSpeed.Text = .Right.DispenseDot.RetractSpeed
                RetractHeight.Text = .Right.DispenseDot.RetractHeight
                ClearanceHeight.Text = .Right.DispenseDot.ClearanceHeight
            End If
        End With
        '        Jetting(Valve)
        '        Auger(Valve)
        '        Slider(Valve)
        'Time Pressure Valve
        'Time Pressure Syringe
        'BoxStep(False)
        'If IDS.Data.Hardware.Dispenser.Left.HeadType.ToUpper = "AUGER VALVE" Then
        '    If Not IDS.Data.Hardware.Dispenser.Left.AugerCalOnce Then
        '        BoxStep(False)
        '    End If
        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType.ToUpper = "JETTING VALVE" Then
        '    If Not IDS.Data.Hardware.Dispenser.Left.JettingCalOnce Then
        '        BoxStep(False)
        '    End If
        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType.ToUpper = "SLIDER VALVE" Then
        '    If Not IDS.Data.Hardware.Dispenser.Left.SlideValveCalOnce Then
        '        BoxStep(False)
        '    End If
        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType.ToUpper = "TIME PRESSURE VALVE" Then
        '    If Not IDS.Data.Hardware.Dispenser.Left.TimePressureValveCalOnce Then
        '        BoxStep(False)
        '    End If
        'ElseIf IDS.Data.Hardware.Dispenser.Left.HeadType.ToUpper = "TIME PRESSURE SYRINGE VALVE" Then
        '    If Not IDS.Data.Hardware.Dispenser.Left.TimePressureSyringeCalOnce Then
        '        BoxStep(False)
        '    End If
        'End If
    End Sub
    Private Sub BoxStep(ByVal enable As Boolean)
        'Me.BoxStep2.Enabled = enable
        'Me.BoxStep3.Enabled = enable
        Me.BoxStep4.Enabled = enable
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        Timer1.Enabled = False
        ButtonTest.Text = "Continous Check Dot"
        OffLaser()
        RemovePanel(CurrentControl)
    End Sub

    Private Sub ButtonCalibratePosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrateZPosition.Click
        If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Then
            MessageBox.Show("Jetting valve do not need to calibrate Z")
            Return
        End If
        Me.NeedleCalibrationState = "Running"
        PanelToBeAdded.Enabled = False
        SetServiceSpeed()
        If CalibrateNeedleZPosition() Then
            MessageBox.Show("Z calibration success!")
        End If
        BoxStep(True)
        PanelToBeAdded.Enabled = True
    End Sub
    Private localCalib As Boolean = False
    Private Sub Calibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Calibrate.Click
        If Calibrate.Text = "Calibrate" Then
            Calibrate.Text = "Stop Calibrate"
            Vision.FrmVision.SwitchCamera("View Camera")
            localCalib = True
            Timer1.Enabled = False
            Timer1.Stop()
            Dim flag As Boolean = False
            BoxStep1.Enabled = flag
            BoxStep2.Enabled = flag
            BoxStep3.Enabled = flag
            BoxStep1Jetting.Enabled = flag
            ButtonExit.Enabled = flag
            ButtonRevert.Enabled = flag
            ButtonSave.Enabled = flag
            m_Tri.SteppingButtons.Enabled = flag
            'PanelToBeAdded.Enabled = False
            NeedleCalibrationState = "Running"
            If PanelToBeAdded.Visible = True Then SaveData()
            NeedleCalibration()
            Calibrate.Text = "Calibrate"
            localCalib = False
            flag = True
            BoxStep1.Enabled = flag
            ButtonRevert.Enabled = flag
            ButtonSave.Enabled = flag
            BoxStep2.Enabled = flag
            BoxStep3.Enabled = flag
            BoxStep1Jetting.Enabled = flag
            ButtonExit.Enabled = flag
            m_Tri.SteppingButtons.Enabled = flag
            Calibrate.Enabled = True
            'PanelToBeAdded.Enabled = True
        Else
            Dim pos(0) As Integer
            pos(0) = 1
            m_Tri.m_TriCtrl.SetTable(126, 1, pos)
            Calibrate.Enabled = False
            Me.NeedleCalibrationState = "Stopped"
            Calibrate.Text = "Stopping"
        End If

    End Sub

    Private Sub ButtonCalibrateVision_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrateVision.Click
        If Not CheckZ() Then
            Return
        End If
        PanelToBeAdded.Enabled = False
        SetServiceSpeed()
        OnLaser()
        CalibrateDispenseDotForVision()
        PanelToBeAdded.Enabled = True
    End Sub
    Private Function CheckState() As Boolean
        If NeedleCalibrationState = "Stopped" Then
            Return False
        ElseIf NeedleCalibrationState = "Running" Then
            Return True
        End If
    End Function
    Public NCRunning As Boolean = False
    Public Function NeedleCalibration(Optional ByVal moveToZ As Boolean = True) As Boolean
        SetServiceSpeed()
        NCRunning = True
        OnLaser()
        If Not IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Then
            If Not MoveToZCalibratorPosition() Then
                NCRunning = False
                Return False
            End If
            If Me.NeedleCalibrationState = "Stopped" Then
                NCRunning = False
                Return False
            End If
            If Not CalibrateNeedleZPosition(moveToZ) Then
                NCRunning = False
                Return False
            End If


        End If
        If Not (CalibrateNeedleXYPosition(moveToZ)) Then
            NCRunning = False
            Return False
        End If
        NCRunning = False
        Return True
    End Function

    Public Function MoveToZCalibratorPosition() As Boolean
        With IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition
            position(0) = .X
            position(1) = .Y
            position(2) = .Z
            m_Tri.SetMachineRun()
            SetServiceSpeed()
            If Not m_Tri.Move_XY(position) Then GoTo StopCalibration
            If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
            If Not m_Tri.Move_Z(position(2)) Then GoTo StopCalibration
            If Me.NeedleCalibrationState = "Stopped" Then GoTo StopCalibration
            m_Tri.SetMachineStop()
            Return True
        End With
StopCalibration:
        'MsgBox("Checking the Z position for needle calibration stopped prematurely.")
        m_Tri.SetMachineStop()
        SetServiceSpeed()
        m_Tri.Move_Z(0)
        Calibrate.Enabled = True
        Return False
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        Timer1.Enabled = False
        If stopTimer Then
            Timer1.Stop()
            Timer1.Enabled = False
            Return
        End If
        tbStatus.Text = ""
        Dim left_tip_x As Double
        Dim left_tip_y As Double
        Dim right_tip_x As Double
        Dim right_tip_y As Double
        Vision.FrmVision.DisplayIndicator()
        'try to make a way to automate this process in the future
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)
        Try
            'If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, Threshold.Value, MaxRadius.Value, MinRadius.Value, CloseValue.Value, OpenValue.Value, Roughness.Value, Compactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
            If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, m_Threshold, m_maxRad / 2, m_minRad / 2, m_Close, m_Open, m_Rough, m_Compact, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
                IDS.Data.Hardware.Needle.Left.DotDiameter = Vision.FrmVision.diameter
                tbStatus.Text = "Dot found. Diameter = " + IDS.Data.Hardware.Needle.Left.DotDiameter.ToString("0.000") + "mm"
            Else
                tbStatus.Text = "Dot not found"
            End If
            'Vision.FrmVision.DisplayIndicator()

        Catch ex As Exception
            ExceptionDisplay(ex)
            Return
        End Try
        Timer1.Enabled = True
        Timer1.Start()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub
    Dim stopTimer As Boolean = False
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTest.Click
        If ButtonTest.Text = "Continous Check Dot" Then
            ButtonTest.Text = "Stop Dot Checking"
            btCheckDotOnce.Enabled = False
            Me.btSaveDotSize.Enabled = False
            stopTimer = False
            Vision.FrmVision.SwitchCamera("Teach Camera")
            Vision.FrmVision.DisplayIndicator()
            Timer1.Enabled = True
            Timer1.Start()

        Else
            Vision.FrmVision.DisplayIndicator()
            ButtonTest.Text = "Continous Check Dot"
            btCheckDotOnce.Enabled = True
            Me.btSaveDotSize.Enabled = True
            stopTimer = True
            Timer1.Enabled = False
            Timer1.Stop()
        End If
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveData()
    End Sub

    Private Sub ButtonCalibrateZPositionJetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrateZPositionJetting.Click
        m_Tri.GetIDSState()
        IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z = m_Tri.ZPosition
        IDS.Data.SaveData()
    End Sub

    Private Sub Brightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Brightness.TextChanged
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)
    End Sub

    Private Sub RetractHeight_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetractHeight.TextChanged
        Try
            If CDbl(RetractHeight.Text) > CDbl(ClearanceHeight.Text) Then
                RetractHeight.Text = ClearanceHeight.Text
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ClearanceHeight_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearanceHeight.TextChanged
        If CDbl(RetractHeight.Text) > CDbl(ClearanceHeight.Text) Then
            ClearanceHeight.Text = RetractHeight.Text
        End If
    End Sub

    Private Sub btCheckDotOnce_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCheckDotOnce.Click
        Timer1.Stop()
        tbStatus.Text = ""
        Timer1.Enabled = False
        btCheckDotOnce.Enabled = False
        ButtonTest.Text = "Continous Check Dot"
        Vision.FrmVision.SwitchCamera("Teach Camera")
        Vision.FrmVision.DisplayIndicator()
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)
        If MinRadius.Value >= MaxRadius.Value Then
            MinRadius.Value = MaxRadius.Value - 0.1
            MessageBox.Show("Minimum diameter cannot be equal to or greater than maximum diameter. It will be reset to smaller than Maximum diamter - " & MinRadius.Value & " now")
            Return
        End If
        Try
            'If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, Threshold.Value, MaxRadius.Value / 2, MinRadius.Value / 2, CloseValue.Value, OpenValue.Value, Roughness.Value, Compactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
            '    'Vision.FrmVision.DisplayIndicator()
            '    'MessageBox.Show("Dot found! Dot diameter(mm) is " & Vision.FrmVision.diameter)
            '    IDS.Data.Hardware.Needle.Left.DotDiameter = Vision.FrmVision.diameter
            '    tbStatus.Text = "Dot found. Diameter = " + IDS.Data.Hardware.Needle.Left.DotDiameter.ToString("0.000") + "mm"
            'Else
            '    'MessageBox.Show("Dot not found!")
            '    tbStatus.Text = "Dot not found"
            'End If
            If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, m_Threshold, m_maxRad / 2, m_minRad / 2, m_Close, m_Open, m_Rough, m_Compact, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
                IDS.Data.Hardware.Needle.Left.DotDiameter = Vision.FrmVision.diameter
                tbStatus.Text = "Dot found. Diameter = " + IDS.Data.Hardware.Needle.Left.DotDiameter.ToString("0.000") + "mm"
            Else
                tbStatus.Text = "Dot not found"
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
            btCheckDotOnce.Enabled = True
            Return
        End Try
        btCheckDotOnce.Enabled = True
    End Sub

    Private Sub btSaveDotSize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSaveDotSize.Click
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)
        Try
            If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, Threshold.Value, MaxRadius.Value / 2, MinRadius.Value / 2, CloseValue.Value, OpenValue.Value, Roughness.Value, Compactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
                Vision.FrmVision.DisplayIndicator()
                MessageBox.Show("Dot found! Dot diameter(mm) is " & Vision.FrmVision.diameter + " and saved")
                IDS.Data.Hardware.Needle.Left.DotDiameter = Vision.FrmVision.diameter
            Else
                MessageBox.Show("Dot not found! Can't save dot size")
            End If

        Catch ex As Exception
            ExceptionDisplay(ex)
            Return
        End Try
    End Sub

    Private Sub btPurgeNow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btPurgeNow.Click
        'download dispenser settings
       
        If btPurgeNow.Text = "Off Purge" Then
            btPurgeNow.Text = "Purge at current position"
            m_Tri.TurnOff("Left Needle IO")
        Else
            If LeftHead.Checked Then
                MyDispenserSettings.DownloadDispenserSettings("Left")
            ElseIf RightHead.Checked Then
                MyDispenserSettings.DownloadDispenserSettings("Right")
            End If
            btPurgeNow.Text = "Off Purge"
            m_Tri.TurnOn("Left Needle IO")
        End If

    End Sub

    Private Sub Threshold_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Threshold.ValueChanged
        Me.m_Threshold = Me.Threshold.Value
    End Sub

    Private Sub Brightness_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Brightness.ValueChanged
        m_Brightness = Me.Brightness.Value
    End Sub

    Private Sub OpenValue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenValue.ValueChanged
        m_Open = Me.OpenValue.Value
    End Sub

    Private Sub CloseValue_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseValue.ValueChanged
        Me.m_Close = Me.CloseValue.Value
    End Sub

    Private Sub MaxRadius_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxRadius.ValueChanged
        Me.m_maxRad = Me.MaxRadius.Value
    End Sub

    Private Sub MinRadius_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinRadius.ValueChanged
        If MinRadius.Value > MaxRadius.Value Then
            MinRadius.Value = MaxRadius.Value - 0.1
            Return
        End If
        Me.m_minRad = Me.MinRadius.Value
    End Sub

    Private Sub Roughness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Roughness.ValueChanged
        Me.m_Rough = Me.Roughness.Value
    End Sub

    Private Sub Compactness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Compactness.ValueChanged
        Me.m_Compact = Me.Compactness.Value
    End Sub

    Private Sub btCheckCalibrationAccuracy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCheckCalibrationAccuracy.Click
        Vision.FrmVision.DisplayIndicator()
        Vision.FrmVision.SwitchCamera("Teach Camera")
        IDS.Devices.Vision.IDSV_SetBrightness(Brightness.Value)
        Try
            If IDS.Devices.Vision.IDSV_NC(BlackBackground.Checked, Threshold.Value, MaxRadius.Value / 2, MinRadius.Value / 2, CloseValue.Value, OpenValue.Value, Roughness.Value, Compactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.Offset_X, MyNeedleCalibrationSetup1.Offset_Y) Then
                IDS.Data.Hardware.Needle.Left.DotDiameter = Vision.FrmVision.diameter
                tbStatus.Text = "Dot found. Dot offset X = " + MyNeedleCalibrationSetup1.Offset_X.ToString("0.000") + "mm" + " Y = " + MyNeedleCalibrationSetup1.Offset_Y.ToString("0.000") + " mm"
            Else
                tbStatus.Text = "Dot not found"
            End If

        Catch ex As Exception
            ExceptionDisplay(ex)
            Return
        End Try
    End Sub

    'Private Sub ApproachHeight_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApproachHeight.ValueChanged
    '    CheckZ(ApproachHeight)
    'End Sub

    'Private Sub RetractHeight_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RetractHeight.ValueChanged
    '    CheckZ(RetractHeight)
    'End Sub

    'Private Sub ClearanceHeight_ValueChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearanceHeight.ValueChanged
    '    CheckZ(ClearanceHeight)
    'End Sub

    Private Function CheckZ() As Boolean
        If CDbl(Me.NeedleGap.Text) > CDbl(ApproachHeight.Text) Then
            'ctrl.Text = Me.NeedleGap.Text
            MessageBox.Show("Approach height cannot be smaller than needle gap!")
            Return False
        End If
        If CDbl(Me.NeedleGap.Text) > CDbl(RetractHeight.Text) Then
            'ctrl.Text = Me.NeedleGap.Text
            MessageBox.Show("Retract height cannot be smaller than needle gap!")
            Return False
        End If
        If CDbl(Me.NeedleGap.Text) > CDbl(ClearanceHeight.Text) Then
            'ctrl.Text = Me.NeedleGap.Text
            MessageBox.Show("Clearance height cannot be smaller than needle gap!")
            Return False
        End If
        Return True
    End Function
End Class
