Imports DLL_Export_Service

Public Class DispenserSettings

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
    Friend WithEvents MaterialAirPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents SuckbackPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents Download As System.Windows.Forms.Button
    Friend WithEvents SuckbackPressureLabel1 As System.Windows.Forms.Label
    Friend WithEvents RPMLabel As System.Windows.Forms.Label
    Friend WithEvents MaterialAirPressureLabel1 As System.Windows.Forms.Label
    Friend WithEvents HeadType As System.Windows.Forms.ComboBox
    Friend WithEvents MaterialAirPressureLabel2 As System.Windows.Forms.Label
    Friend WithEvents SuckbackPressureLabel2 As System.Windows.Forms.Label
    Friend WithEvents RetractDelayLabel As System.Windows.Forms.Label
    Friend WithEvents RetractTimeLabel As System.Windows.Forms.Label
    Friend WithEvents RetractDelayLabel2 As System.Windows.Forms.Label
    Friend WithEvents RetractTimeLabel2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents AutoPurgingOption As System.Windows.Forms.CheckBox
    Friend WithEvents PotLifeOption As System.Windows.Forms.CheckBox
    Friend WithEvents AutoCleaningDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents AutoPurgingDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents AutoPurgingDurationLabel1 As System.Windows.Forms.Label
    Friend WithEvents AutoPurgingDurationLabel2 As System.Windows.Forms.Label
    Friend WithEvents AutoCleaningDurationLabel1 As System.Windows.Forms.Label
    Friend WithEvents AutoCleaningDurationLabel2 As System.Windows.Forms.Label
    Friend WithEvents NeedleColorLabel As System.Windows.Forms.Label
    Friend WithEvents NeedleColor As System.Windows.Forms.ComboBox
    Friend WithEvents NeedleTipLength As System.Windows.Forms.ComboBox
    Friend WithEvents AutoPurgingIntervalLabel3 As System.Windows.Forms.Label
    Friend WithEvents AutoPurgingIntervalHours As System.Windows.Forms.NumericUpDown
    Friend WithEvents AutoPurgingIntervalLabel1 As System.Windows.Forms.Label
    Friend WithEvents AutoPurgingIntervalMinutes As System.Windows.Forms.NumericUpDown
    Friend WithEvents AutoPurgingIntervalLabel2 As System.Windows.Forms.Label
    Friend WithEvents PotLifeDurationHours As System.Windows.Forms.NumericUpDown
    Friend WithEvents PotLifeDurationLabel3 As System.Windows.Forms.Label
    Friend WithEvents PotLifeDurationLabel1 As System.Windows.Forms.Label
    Friend WithEvents PotLifeDurationLabel2 As System.Windows.Forms.Label
    Friend WithEvents PotLifeDurationMinutes As System.Windows.Forms.NumericUpDown
    Friend WithEvents CurrentHeads As System.Windows.Forms.ComboBox
    Friend WithEvents MaterialInfo As System.Windows.Forms.TextBox
    Friend WithEvents TipLengthLabel As System.Windows.Forms.Label
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents SliderOrJetting As System.Windows.Forms.GroupBox
    Friend WithEvents Auger As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Pulse As System.Windows.Forms.NumericUpDown
    Friend WithEvents pause As System.Windows.Forms.NumericUpDown
    Friend WithEvents Count As System.Windows.Forms.NumericUpDown
    Friend WithEvents SliderOrJettingTemperature As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractDelay As System.Windows.Forms.NumericUpDown
    Friend WithEvents RPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents AugerTemperature As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractTime As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents cbEnableGlobalQC As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btZeroPressure As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DispenserSettings))
        Me.MaterialAirPressure = New System.Windows.Forms.NumericUpDown
        Me.SuckbackPressure = New System.Windows.Forms.NumericUpDown
        Me.Download = New System.Windows.Forms.Button
        Me.SuckbackPressureLabel1 = New System.Windows.Forms.Label
        Me.RPMLabel = New System.Windows.Forms.Label
        Me.MaterialAirPressureLabel1 = New System.Windows.Forms.Label
        Me.HeadType = New System.Windows.Forms.ComboBox
        Me.MaterialAirPressureLabel2 = New System.Windows.Forms.Label
        Me.SuckbackPressureLabel2 = New System.Windows.Forms.Label
        Me.RetractDelayLabel = New System.Windows.Forms.Label
        Me.RetractTimeLabel = New System.Windows.Forms.Label
        Me.RetractDelayLabel2 = New System.Windows.Forms.Label
        Me.RetractTimeLabel2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btZeroPressure = New System.Windows.Forms.Button
        Me.SliderOrJetting = New System.Windows.Forms.GroupBox
        Me.Pulse = New System.Windows.Forms.NumericUpDown
        Me.pause = New System.Windows.Forms.NumericUpDown
        Me.Count = New System.Windows.Forms.NumericUpDown
        Me.SliderOrJettingTemperature = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Auger = New System.Windows.Forms.GroupBox
        Me.RetractDelay = New System.Windows.Forms.NumericUpDown
        Me.RPM = New System.Windows.Forms.NumericUpDown
        Me.AugerTemperature = New System.Windows.Forms.NumericUpDown
        Me.RetractTime = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.cbEnableGlobalQC = New System.Windows.Forms.CheckBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.MaterialInfo = New System.Windows.Forms.TextBox
        Me.TipLengthLabel = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.NeedleColorLabel = New System.Windows.Forms.Label
        Me.NeedleColor = New System.Windows.Forms.ComboBox
        Me.NeedleTipLength = New System.Windows.Forms.ComboBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.AutoPurgingOption = New System.Windows.Forms.CheckBox
        Me.AutoPurgingIntervalLabel3 = New System.Windows.Forms.Label
        Me.AutoPurgingIntervalHours = New System.Windows.Forms.NumericUpDown
        Me.AutoPurgingDuration = New System.Windows.Forms.NumericUpDown
        Me.AutoPurgingDurationLabel1 = New System.Windows.Forms.Label
        Me.AutoPurgingIntervalLabel1 = New System.Windows.Forms.Label
        Me.AutoPurgingIntervalMinutes = New System.Windows.Forms.NumericUpDown
        Me.AutoPurgingIntervalLabel2 = New System.Windows.Forms.Label
        Me.AutoPurgingDurationLabel2 = New System.Windows.Forms.Label
        Me.AutoCleaningDuration = New System.Windows.Forms.NumericUpDown
        Me.AutoCleaningDurationLabel1 = New System.Windows.Forms.Label
        Me.AutoCleaningDurationLabel2 = New System.Windows.Forms.Label
        Me.PotLifeDurationHours = New System.Windows.Forms.NumericUpDown
        Me.PotLifeDurationLabel3 = New System.Windows.Forms.Label
        Me.PotLifeDurationLabel1 = New System.Windows.Forms.Label
        Me.PotLifeOption = New System.Windows.Forms.CheckBox
        Me.PotLifeDurationLabel2 = New System.Windows.Forms.Label
        Me.PotLifeDurationMinutes = New System.Windows.Forms.NumericUpDown
        Me.CurrentHeads = New System.Windows.Forms.ComboBox
        CType(Me.MaterialAirPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SuckbackPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SliderOrJetting.SuspendLayout()
        CType(Me.Pulse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pause, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Count, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SliderOrJettingTemperature, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Auger.SuspendLayout()
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AugerTemperature, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.AutoPurgingIntervalHours, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoPurgingDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoPurgingIntervalMinutes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoCleaningDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PotLifeDurationHours, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PotLifeDurationMinutes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MaterialAirPressure
        '
        Me.MaterialAirPressure.DecimalPlaces = 1
        Me.MaterialAirPressure.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.MaterialAirPressure.Location = New System.Drawing.Point(96, 88)
        Me.MaterialAirPressure.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.MaterialAirPressure.Name = "MaterialAirPressure"
        Me.MaterialAirPressure.Size = New System.Drawing.Size(72, 27)
        Me.MaterialAirPressure.TabIndex = 54
        Me.MaterialAirPressure.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'SuckbackPressure
        '
        Me.SuckbackPressure.DecimalPlaces = 2
        Me.SuckbackPressure.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.SuckbackPressure.Location = New System.Drawing.Point(96, 128)
        Me.SuckbackPressure.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.SuckbackPressure.Minimum = New Decimal(New Integer() {1, 0, 0, -2147483648})
        Me.SuckbackPressure.Name = "SuckbackPressure"
        Me.SuckbackPressure.Size = New System.Drawing.Size(72, 27)
        Me.SuckbackPressure.TabIndex = 55
        Me.SuckbackPressure.Value = New Decimal(New Integer() {3, 0, 0, -2147418112})
        Me.SuckbackPressure.Visible = False
        '
        'Download
        '
        Me.Download.Location = New System.Drawing.Point(0, 208)
        Me.Download.Name = "Download"
        Me.Download.Size = New System.Drawing.Size(200, 32)
        Me.Download.TabIndex = 53
        Me.Download.Text = "Download"
        '
        'SuckbackPressureLabel1
        '
        Me.SuckbackPressureLabel1.Location = New System.Drawing.Point(8, 128)
        Me.SuckbackPressureLabel1.Name = "SuckbackPressureLabel1"
        Me.SuckbackPressureLabel1.Size = New System.Drawing.Size(88, 24)
        Me.SuckbackPressureLabel1.TabIndex = 52
        Me.SuckbackPressureLabel1.Text = "Suckback"
        Me.SuckbackPressureLabel1.Visible = False
        '
        'RPMLabel
        '
        Me.RPMLabel.Location = New System.Drawing.Point(8, 40)
        Me.RPMLabel.Name = "RPMLabel"
        Me.RPMLabel.Size = New System.Drawing.Size(96, 24)
        Me.RPMLabel.TabIndex = 57
        Me.RPMLabel.Text = "Valve RPM"
        '
        'MaterialAirPressureLabel1
        '
        Me.MaterialAirPressureLabel1.Location = New System.Drawing.Point(8, 88)
        Me.MaterialAirPressureLabel1.Name = "MaterialAirPressureLabel1"
        Me.MaterialAirPressureLabel1.Size = New System.Drawing.Size(80, 24)
        Me.MaterialAirPressureLabel1.TabIndex = 51
        Me.MaterialAirPressureLabel1.Text = "Pressure"
        '
        'HeadType
        '
        Me.HeadType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.HeadType.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.HeadType.ItemHeight = 20
        Me.HeadType.Items.AddRange(New Object() {"Jetting Valve", "Auger Valve", "Slider Valve", "Time Pressure Valve", "Time Pressure Syringe"})
        Me.HeadType.Location = New System.Drawing.Point(16, 40)
        Me.HeadType.Name = "HeadType"
        Me.HeadType.Size = New System.Drawing.Size(184, 28)
        Me.HeadType.TabIndex = 46
        '
        'MaterialAirPressureLabel2
        '
        Me.MaterialAirPressureLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.MaterialAirPressureLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.MaterialAirPressureLabel2.Location = New System.Drawing.Point(168, 88)
        Me.MaterialAirPressureLabel2.Name = "MaterialAirPressureLabel2"
        Me.MaterialAirPressureLabel2.Size = New System.Drawing.Size(32, 24)
        Me.MaterialAirPressureLabel2.TabIndex = 47
        Me.MaterialAirPressureLabel2.Text = "bar"
        Me.MaterialAirPressureLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'SuckbackPressureLabel2
        '
        Me.SuckbackPressureLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.SuckbackPressureLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.SuckbackPressureLabel2.Location = New System.Drawing.Point(168, 128)
        Me.SuckbackPressureLabel2.Name = "SuckbackPressureLabel2"
        Me.SuckbackPressureLabel2.Size = New System.Drawing.Size(32, 24)
        Me.SuckbackPressureLabel2.TabIndex = 48
        Me.SuckbackPressureLabel2.Text = "bar"
        Me.SuckbackPressureLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.SuckbackPressureLabel2.Visible = False
        '
        'RetractDelayLabel
        '
        Me.RetractDelayLabel.Location = New System.Drawing.Point(8, 120)
        Me.RetractDelayLabel.Name = "RetractDelayLabel"
        Me.RetractDelayLabel.Size = New System.Drawing.Size(168, 24)
        Me.RetractDelayLabel.TabIndex = 59
        Me.RetractDelayLabel.Text = "Valve Retract Delay"
        '
        'RetractTimeLabel
        '
        Me.RetractTimeLabel.Location = New System.Drawing.Point(8, 80)
        Me.RetractTimeLabel.Name = "RetractTimeLabel"
        Me.RetractTimeLabel.Size = New System.Drawing.Size(160, 24)
        Me.RetractTimeLabel.TabIndex = 56
        Me.RetractTimeLabel.Text = "Valve Retract Time"
        '
        'RetractDelayLabel2
        '
        Me.RetractDelayLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.RetractDelayLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RetractDelayLabel2.Location = New System.Drawing.Point(232, 128)
        Me.RetractDelayLabel2.Name = "RetractDelayLabel2"
        Me.RetractDelayLabel2.Size = New System.Drawing.Size(32, 16)
        Me.RetractDelayLabel2.TabIndex = 50
        Me.RetractDelayLabel2.Text = "s"
        Me.RetractDelayLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RetractTimeLabel2
        '
        Me.RetractTimeLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.RetractTimeLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RetractTimeLabel2.Location = New System.Drawing.Point(232, 88)
        Me.RetractTimeLabel2.Name = "RetractTimeLabel2"
        Me.RetractTimeLabel2.Size = New System.Drawing.Size(32, 16)
        Me.RetractTimeLabel2.TabIndex = 49
        Me.RetractTimeLabel2.Text = "s"
        Me.RetractTimeLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btZeroPressure)
        Me.GroupBox1.Controls.Add(Me.SliderOrJetting)
        Me.GroupBox1.Controls.Add(Me.MaterialAirPressureLabel2)
        Me.GroupBox1.Controls.Add(Me.SuckbackPressureLabel2)
        Me.GroupBox1.Controls.Add(Me.MaterialAirPressure)
        Me.GroupBox1.Controls.Add(Me.SuckbackPressure)
        Me.GroupBox1.Controls.Add(Me.Download)
        Me.GroupBox1.Controls.Add(Me.SuckbackPressureLabel1)
        Me.GroupBox1.Controls.Add(Me.MaterialAirPressureLabel1)
        Me.GroupBox1.Controls.Add(Me.HeadType)
        Me.GroupBox1.Controls.Add(Me.Auger)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 256)
        Me.GroupBox1.TabIndex = 64
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Dispenser Information"
        '
        'btZeroPressure
        '
        Me.btZeroPressure.Location = New System.Drawing.Point(0, 168)
        Me.btZeroPressure.Name = "btZeroPressure"
        Me.btZeroPressure.Size = New System.Drawing.Size(200, 32)
        Me.btZeroPressure.TabIndex = 57
        Me.btZeroPressure.Text = "Off Pressure/Suckback"
        '
        'SliderOrJetting
        '
        Me.SliderOrJetting.Controls.Add(Me.Pulse)
        Me.SliderOrJetting.Controls.Add(Me.pause)
        Me.SliderOrJetting.Controls.Add(Me.Count)
        Me.SliderOrJetting.Controls.Add(Me.SliderOrJettingTemperature)
        Me.SliderOrJetting.Controls.Add(Me.Label1)
        Me.SliderOrJetting.Controls.Add(Me.Label3)
        Me.SliderOrJetting.Controls.Add(Me.Label4)
        Me.SliderOrJetting.Controls.Add(Me.Label6)
        Me.SliderOrJetting.Controls.Add(Me.Label8)
        Me.SliderOrJetting.Controls.Add(Me.Label9)
        Me.SliderOrJetting.Controls.Add(Me.Label5)
        Me.SliderOrJetting.Location = New System.Drawing.Point(208, 40)
        Me.SliderOrJetting.Name = "SliderOrJetting"
        Me.SliderOrJetting.Size = New System.Drawing.Size(280, 200)
        Me.SliderOrJetting.TabIndex = 56
        Me.SliderOrJetting.TabStop = False
        Me.SliderOrJetting.Text = "Slider/Jetting Valve Controller"
        '
        'Pulse
        '
        Me.Pulse.DecimalPlaces = 3
        Me.Pulse.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.Pulse.Location = New System.Drawing.Point(160, 40)
        Me.Pulse.Name = "Pulse"
        Me.Pulse.Size = New System.Drawing.Size(80, 27)
        Me.Pulse.TabIndex = 71
        '
        'pause
        '
        Me.pause.DecimalPlaces = 3
        Me.pause.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.pause.Location = New System.Drawing.Point(160, 80)
        Me.pause.Name = "pause"
        Me.pause.Size = New System.Drawing.Size(80, 27)
        Me.pause.TabIndex = 72
        '
        'Count
        '
        Me.Count.Location = New System.Drawing.Point(160, 120)
        Me.Count.Name = "Count"
        Me.Count.Size = New System.Drawing.Size(80, 27)
        Me.Count.TabIndex = 69
        '
        'SliderOrJettingTemperature
        '
        Me.SliderOrJettingTemperature.DecimalPlaces = 2
        Me.SliderOrJettingTemperature.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.SliderOrJettingTemperature.Location = New System.Drawing.Point(160, 160)
        Me.SliderOrJettingTemperature.Name = "SliderOrJettingTemperature"
        Me.SliderOrJettingTemperature.Size = New System.Drawing.Size(80, 27)
        Me.SliderOrJettingTemperature.TabIndex = 70
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 24)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Pulse On Time"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 24)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "No Of Dispense"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 24)
        Me.Label4.TabIndex = 66
        Me.Label4.Text = "Pulse Off Time"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(240, 88)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 64
        Me.Label6.Text = "ms"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 160)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 24)
        Me.Label8.TabIndex = 68
        Me.Label8.Text = "Temperature"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(240, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 16)
        Me.Label9.TabIndex = 65
        Me.Label9.Text = "ms"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(240, 168)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(16, 16)
        Me.Label5.TabIndex = 64
        Me.Label5.Text = "C"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Auger
        '
        Me.Auger.Controls.Add(Me.RetractDelay)
        Me.Auger.Controls.Add(Me.RPM)
        Me.Auger.Controls.Add(Me.AugerTemperature)
        Me.Auger.Controls.Add(Me.RetractTime)
        Me.Auger.Controls.Add(Me.RetractDelayLabel2)
        Me.Auger.Controls.Add(Me.RetractTimeLabel2)
        Me.Auger.Controls.Add(Me.RPMLabel)
        Me.Auger.Controls.Add(Me.RetractDelayLabel)
        Me.Auger.Controls.Add(Me.RetractTimeLabel)
        Me.Auger.Controls.Add(Me.Label7)
        Me.Auger.Location = New System.Drawing.Point(208, 40)
        Me.Auger.Name = "Auger"
        Me.Auger.Size = New System.Drawing.Size(280, 200)
        Me.Auger.TabIndex = 56
        Me.Auger.TabStop = False
        Me.Auger.Text = "Auger Valve Controller"
        '
        'RetractDelay
        '
        Me.RetractDelay.DecimalPlaces = 2
        Me.RetractDelay.Location = New System.Drawing.Point(168, 120)
        Me.RetractDelay.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RetractDelay.Name = "RetractDelay"
        Me.RetractDelay.Size = New System.Drawing.Size(72, 27)
        Me.RetractDelay.TabIndex = 68
        '
        'RPM
        '
        Me.RPM.DecimalPlaces = 1
        Me.RPM.Location = New System.Drawing.Point(168, 40)
        Me.RPM.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RPM.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.RPM.Name = "RPM"
        Me.RPM.Size = New System.Drawing.Size(72, 27)
        Me.RPM.TabIndex = 69
        Me.RPM.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'AugerTemperature
        '
        Me.AugerTemperature.Location = New System.Drawing.Point(168, 160)
        Me.AugerTemperature.Name = "AugerTemperature"
        Me.AugerTemperature.Size = New System.Drawing.Size(72, 27)
        Me.AugerTemperature.TabIndex = 66
        Me.AugerTemperature.Visible = False
        '
        'RetractTime
        '
        Me.RetractTime.DecimalPlaces = 2
        Me.RetractTime.Location = New System.Drawing.Point(168, 80)
        Me.RetractTime.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RetractTime.Name = "RetractTime"
        Me.RetractTime.Size = New System.Drawing.Size(72, 27)
        Me.RetractTime.TabIndex = 67
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(112, 24)
        Me.Label7.TabIndex = 68
        Me.Label7.Text = "Temperature"
        Me.Label7.Visible = False
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox4)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox3)
        Me.PanelToBeAdded.Controls.Add(Me.CurrentHeads)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(24, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 65
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cbEnableGlobalQC)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 664)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(496, 72)
        Me.GroupBox4.TabIndex = 67
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "QC Settings"
        Me.GroupBox4.Visible = False
        '
        'cbEnableGlobalQC
        '
        Me.cbEnableGlobalQC.Location = New System.Drawing.Point(16, 32)
        Me.cbEnableGlobalQC.Name = "cbEnableGlobalQC"
        Me.cbEnableGlobalQC.Size = New System.Drawing.Size(384, 24)
        Me.cbEnableGlobalQC.TabIndex = 0
        Me.cbEnableGlobalQC.Text = "Enable QC Check for all Dots in the recipe."
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(208, 32)
        Me.Label17.TabIndex = 66
        Me.Label17.Text = "Dispenser Settings"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 65
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(288, 792)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 40)
        Me.ButtonSave.TabIndex = 50
        Me.ButtonSave.Text = "Save"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(400, 792)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(75, 40)
        Me.ButtonRevert.TabIndex = 49
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.MaterialInfo)
        Me.GroupBox2.Controls.Add(Me.TipLengthLabel)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.NeedleColorLabel)
        Me.GroupBox2.Controls.Add(Me.NeedleColor)
        Me.GroupBox2.Controls.Add(Me.NeedleTipLength)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 360)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(240, 296)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Syringe Information"
        '
        'MaterialInfo
        '
        Me.MaterialInfo.AutoSize = False
        Me.MaterialInfo.Location = New System.Drawing.Point(16, 144)
        Me.MaterialInfo.Name = "MaterialInfo"
        Me.MaterialInfo.Size = New System.Drawing.Size(208, 136)
        Me.MaterialInfo.TabIndex = 53
        Me.MaterialInfo.Text = "TextBox1"
        '
        'TipLengthLabel
        '
        Me.TipLengthLabel.Location = New System.Drawing.Point(8, 32)
        Me.TipLengthLabel.Name = "TipLengthLabel"
        Me.TipLengthLabel.Size = New System.Drawing.Size(96, 24)
        Me.TipLengthLabel.TabIndex = 52
        Me.TipLengthLabel.Text = "Tip Length"
        Me.TipLengthLabel.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(68, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 24)
        Me.Label2.TabIndex = 51
        Me.Label2.Text = "Material Info"
        '
        'NeedleColorLabel
        '
        Me.NeedleColorLabel.Location = New System.Drawing.Point(8, 72)
        Me.NeedleColorLabel.Name = "NeedleColorLabel"
        Me.NeedleColorLabel.Size = New System.Drawing.Size(88, 24)
        Me.NeedleColorLabel.TabIndex = 52
        Me.NeedleColorLabel.Text = "Tip Color"
        Me.NeedleColorLabel.Visible = False
        '
        'NeedleColor
        '
        Me.NeedleColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.NeedleColor.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NeedleColor.ItemHeight = 20
        Me.NeedleColor.Items.AddRange(New Object() {"Olive", "Brown", "Grey", "White", "Green", "Black", "Pink", "Purple", "Blue", "Orange", "Peach", "Red", "Clear", "Light Blue", "Lavender", "Yellow"})
        Me.NeedleColor.Location = New System.Drawing.Point(112, 72)
        Me.NeedleColor.Name = "NeedleColor"
        Me.NeedleColor.Size = New System.Drawing.Size(112, 28)
        Me.NeedleColor.TabIndex = 46
        '
        'NeedleTipLength
        '
        Me.NeedleTipLength.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.NeedleTipLength.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.NeedleTipLength.ItemHeight = 20
        Me.NeedleTipLength.Items.AddRange(New Object() {"0.25 inch", "0.50 inch", "1.00 inch", "1.50 inch"})
        Me.NeedleTipLength.Location = New System.Drawing.Point(112, 32)
        Me.NeedleTipLength.Name = "NeedleTipLength"
        Me.NeedleTipLength.Size = New System.Drawing.Size(112, 28)
        Me.NeedleTipLength.TabIndex = 46
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.AutoPurgingOption)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingIntervalLabel3)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingIntervalHours)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingDuration)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingDurationLabel1)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingIntervalLabel1)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingIntervalMinutes)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingIntervalLabel2)
        Me.GroupBox3.Controls.Add(Me.AutoPurgingDurationLabel2)
        Me.GroupBox3.Controls.Add(Me.AutoCleaningDuration)
        Me.GroupBox3.Controls.Add(Me.AutoCleaningDurationLabel1)
        Me.GroupBox3.Controls.Add(Me.AutoCleaningDurationLabel2)
        Me.GroupBox3.Controls.Add(Me.PotLifeDurationHours)
        Me.GroupBox3.Controls.Add(Me.PotLifeDurationLabel3)
        Me.GroupBox3.Controls.Add(Me.PotLifeDurationLabel1)
        Me.GroupBox3.Controls.Add(Me.PotLifeOption)
        Me.GroupBox3.Controls.Add(Me.PotLifeDurationLabel2)
        Me.GroupBox3.Controls.Add(Me.PotLifeDurationMinutes)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(256, 360)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(248, 296)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Auto Purging Setup"
        '
        'AutoPurgingOption
        '
        Me.AutoPurgingOption.Location = New System.Drawing.Point(16, 32)
        Me.AutoPurgingOption.Name = "AutoPurgingOption"
        Me.AutoPurgingOption.Size = New System.Drawing.Size(128, 24)
        Me.AutoPurgingOption.TabIndex = 56
        Me.AutoPurgingOption.Text = "Auto Purging"
        '
        'AutoPurgingIntervalLabel3
        '
        Me.AutoPurgingIntervalLabel3.BackColor = System.Drawing.SystemColors.Control
        Me.AutoPurgingIntervalLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.AutoPurgingIntervalLabel3.Location = New System.Drawing.Point(200, 72)
        Me.AutoPurgingIntervalLabel3.Name = "AutoPurgingIntervalLabel3"
        Me.AutoPurgingIntervalLabel3.Size = New System.Drawing.Size(16, 24)
        Me.AutoPurgingIntervalLabel3.TabIndex = 47
        Me.AutoPurgingIntervalLabel3.Text = "m"
        Me.AutoPurgingIntervalLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AutoPurgingIntervalHours
        '
        Me.AutoPurgingIntervalHours.Location = New System.Drawing.Point(80, 72)
        Me.AutoPurgingIntervalHours.Maximum = New Decimal(New Integer() {24, 0, 0, 0})
        Me.AutoPurgingIntervalHours.Name = "AutoPurgingIntervalHours"
        Me.AutoPurgingIntervalHours.Size = New System.Drawing.Size(48, 27)
        Me.AutoPurgingIntervalHours.TabIndex = 54
        Me.AutoPurgingIntervalHours.Value = New Decimal(New Integer() {1, 0, 0, 131072})
        '
        'AutoPurgingDuration
        '
        Me.AutoPurgingDuration.Location = New System.Drawing.Point(128, 120)
        Me.AutoPurgingDuration.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.AutoPurgingDuration.Name = "AutoPurgingDuration"
        Me.AutoPurgingDuration.Size = New System.Drawing.Size(72, 27)
        Me.AutoPurgingDuration.TabIndex = 55
        Me.AutoPurgingDuration.Value = New Decimal(New Integer() {500, 0, 0, 0})
        Me.AutoPurgingDuration.Visible = False
        '
        'AutoPurgingDurationLabel1
        '
        Me.AutoPurgingDurationLabel1.Location = New System.Drawing.Point(16, 120)
        Me.AutoPurgingDurationLabel1.Name = "AutoPurgingDurationLabel1"
        Me.AutoPurgingDurationLabel1.Size = New System.Drawing.Size(104, 24)
        Me.AutoPurgingDurationLabel1.TabIndex = 52
        Me.AutoPurgingDurationLabel1.Text = "Purge Time"
        Me.AutoPurgingDurationLabel1.Visible = False
        '
        'AutoPurgingIntervalLabel1
        '
        Me.AutoPurgingIntervalLabel1.Location = New System.Drawing.Point(16, 72)
        Me.AutoPurgingIntervalLabel1.Name = "AutoPurgingIntervalLabel1"
        Me.AutoPurgingIntervalLabel1.Size = New System.Drawing.Size(56, 24)
        Me.AutoPurgingIntervalLabel1.TabIndex = 51
        Me.AutoPurgingIntervalLabel1.Text = "Every"
        '
        'AutoPurgingIntervalMinutes
        '
        Me.AutoPurgingIntervalMinutes.Location = New System.Drawing.Point(152, 72)
        Me.AutoPurgingIntervalMinutes.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.AutoPurgingIntervalMinutes.Name = "AutoPurgingIntervalMinutes"
        Me.AutoPurgingIntervalMinutes.Size = New System.Drawing.Size(48, 27)
        Me.AutoPurgingIntervalMinutes.TabIndex = 54
        Me.AutoPurgingIntervalMinutes.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'AutoPurgingIntervalLabel2
        '
        Me.AutoPurgingIntervalLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.AutoPurgingIntervalLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.AutoPurgingIntervalLabel2.Location = New System.Drawing.Point(128, 72)
        Me.AutoPurgingIntervalLabel2.Name = "AutoPurgingIntervalLabel2"
        Me.AutoPurgingIntervalLabel2.Size = New System.Drawing.Size(16, 24)
        Me.AutoPurgingIntervalLabel2.TabIndex = 47
        Me.AutoPurgingIntervalLabel2.Text = "h"
        Me.AutoPurgingIntervalLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'AutoPurgingDurationLabel2
        '
        Me.AutoPurgingDurationLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.AutoPurgingDurationLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.AutoPurgingDurationLabel2.Location = New System.Drawing.Point(200, 120)
        Me.AutoPurgingDurationLabel2.Name = "AutoPurgingDurationLabel2"
        Me.AutoPurgingDurationLabel2.Size = New System.Drawing.Size(32, 24)
        Me.AutoPurgingDurationLabel2.TabIndex = 47
        Me.AutoPurgingDurationLabel2.Text = "ms"
        Me.AutoPurgingDurationLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.AutoPurgingDurationLabel2.Visible = False
        '
        'AutoCleaningDuration
        '
        Me.AutoCleaningDuration.Location = New System.Drawing.Point(128, 160)
        Me.AutoCleaningDuration.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.AutoCleaningDuration.Name = "AutoCleaningDuration"
        Me.AutoCleaningDuration.Size = New System.Drawing.Size(72, 27)
        Me.AutoCleaningDuration.TabIndex = 55
        Me.AutoCleaningDuration.Value = New Decimal(New Integer() {500, 0, 0, 0})
        Me.AutoCleaningDuration.Visible = False
        '
        'AutoCleaningDurationLabel1
        '
        Me.AutoCleaningDurationLabel1.Location = New System.Drawing.Point(16, 160)
        Me.AutoCleaningDurationLabel1.Name = "AutoCleaningDurationLabel1"
        Me.AutoCleaningDurationLabel1.Size = New System.Drawing.Size(96, 24)
        Me.AutoCleaningDurationLabel1.TabIndex = 52
        Me.AutoCleaningDurationLabel1.Text = "Clean Time"
        Me.AutoCleaningDurationLabel1.Visible = False
        '
        'AutoCleaningDurationLabel2
        '
        Me.AutoCleaningDurationLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.AutoCleaningDurationLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.AutoCleaningDurationLabel2.Location = New System.Drawing.Point(200, 160)
        Me.AutoCleaningDurationLabel2.Name = "AutoCleaningDurationLabel2"
        Me.AutoCleaningDurationLabel2.Size = New System.Drawing.Size(32, 24)
        Me.AutoCleaningDurationLabel2.TabIndex = 47
        Me.AutoCleaningDurationLabel2.Text = "ms"
        Me.AutoCleaningDurationLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PotLifeDurationHours
        '
        Me.PotLifeDurationHours.Location = New System.Drawing.Point(80, 256)
        Me.PotLifeDurationHours.Maximum = New Decimal(New Integer() {24, 0, 0, 0})
        Me.PotLifeDurationHours.Name = "PotLifeDurationHours"
        Me.PotLifeDurationHours.Size = New System.Drawing.Size(48, 27)
        Me.PotLifeDurationHours.TabIndex = 54
        Me.PotLifeDurationHours.Value = New Decimal(New Integer() {1, 0, 0, 131072})
        '
        'PotLifeDurationLabel3
        '
        Me.PotLifeDurationLabel3.BackColor = System.Drawing.SystemColors.Control
        Me.PotLifeDurationLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PotLifeDurationLabel3.Location = New System.Drawing.Point(200, 256)
        Me.PotLifeDurationLabel3.Name = "PotLifeDurationLabel3"
        Me.PotLifeDurationLabel3.Size = New System.Drawing.Size(16, 24)
        Me.PotLifeDurationLabel3.TabIndex = 47
        Me.PotLifeDurationLabel3.Text = "m"
        Me.PotLifeDurationLabel3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PotLifeDurationLabel1
        '
        Me.PotLifeDurationLabel1.Location = New System.Drawing.Point(16, 256)
        Me.PotLifeDurationLabel1.Name = "PotLifeDurationLabel1"
        Me.PotLifeDurationLabel1.Size = New System.Drawing.Size(56, 24)
        Me.PotLifeDurationLabel1.TabIndex = 51
        Me.PotLifeDurationLabel1.Text = "Every"
        '
        'PotLifeOption
        '
        Me.PotLifeOption.Location = New System.Drawing.Point(16, 216)
        Me.PotLifeOption.Name = "PotLifeOption"
        Me.PotLifeOption.Size = New System.Drawing.Size(176, 24)
        Me.PotLifeOption.TabIndex = 56
        Me.PotLifeOption.Text = "Pot Life Monitoring"
        '
        'PotLifeDurationLabel2
        '
        Me.PotLifeDurationLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.PotLifeDurationLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PotLifeDurationLabel2.Location = New System.Drawing.Point(128, 256)
        Me.PotLifeDurationLabel2.Name = "PotLifeDurationLabel2"
        Me.PotLifeDurationLabel2.Size = New System.Drawing.Size(16, 24)
        Me.PotLifeDurationLabel2.TabIndex = 47
        Me.PotLifeDurationLabel2.Text = "h"
        Me.PotLifeDurationLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PotLifeDurationMinutes
        '
        Me.PotLifeDurationMinutes.Location = New System.Drawing.Point(152, 256)
        Me.PotLifeDurationMinutes.Maximum = New Decimal(New Integer() {60, 0, 0, 0})
        Me.PotLifeDurationMinutes.Name = "PotLifeDurationMinutes"
        Me.PotLifeDurationMinutes.Size = New System.Drawing.Size(48, 27)
        Me.PotLifeDurationMinutes.TabIndex = 54
        Me.PotLifeDurationMinutes.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'CurrentHeads
        '
        Me.CurrentHeads.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CurrentHeads.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CurrentHeads.ItemHeight = 20
        Me.CurrentHeads.Items.AddRange(New Object() {"Left Head", "Right Head"})
        Me.CurrentHeads.Location = New System.Drawing.Point(24, 40)
        Me.CurrentHeads.Name = "CurrentHeads"
        Me.CurrentHeads.Size = New System.Drawing.Size(200, 28)
        Me.CurrentHeads.TabIndex = 46
        '
        'DispenserSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(544, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "DispenserSettings"
        Me.Text = "Form1"
        CType(Me.MaterialAirPressure, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SuckbackPressure, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.SliderOrJetting.ResumeLayout(False)
        CType(Me.Pulse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pause, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Count, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SliderOrJettingTemperature, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Auger.ResumeLayout(False)
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RPM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AugerTemperature, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.AutoPurgingIntervalHours, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoPurgingDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoPurgingIntervalMinutes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoCleaningDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PotLifeDurationHours, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PotLifeDurationMinutes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Sub SaveData()

        If MySetup.TwoHead.Checked Then
            IDS.Data.Hardware.Dispenser.CurrentHeads = 2
        ElseIf MySetup.OneHead.Checked Then
            IDS.Data.Hardware.Dispenser.CurrentHeads = 1
        End If

        With IDS.Data.Hardware.Dispenser.Left
            .MaterialAirPressure = MaterialAirPressure.Value
            .SuckbackPressure = SuckbackPressure.Value
            .RPM = RPM.Text
            .RetractTime = RetractTime.Text
            .RetractDelay = RetractDelay.Text
            .Pulse = Pulse.Text
            .Pause = pause.Text
            .Count = Count.Text
            If HeadType.SelectedItem = "Auger Valve" Then
                ' .ValveTemperature = AugerTemperature.Value
            ElseIf HeadType.SelectedItem = "Slider Valve" Or HeadType.SelectedItem = "Jetting Valve" Then
                .ValveTemperature = SliderOrJettingTemperature.Value
            End If
            .NeedleTipLength = NeedleTipLength.SelectedItem
            .NeedleColor = NeedleColor.SelectedItem
            .MaterialInfo = MaterialInfo.Text
            .AutoPurgingDuration = AutoPurgingDuration.Value
            .AutoCleaningDuration = AutoCleaningDuration.Value
            .AutoPurgingInterval = (60 * AutoPurgingIntervalHours.Value) + AutoPurgingIntervalMinutes.Value
            .PotLifeDuration = (60 * PotLifeDurationHours.Value) + PotLifeDurationMinutes.Value
            .PotLifeOption = PotLifeOption.Checked
            .AutoPurgingOption = AutoPurgingOption.Checked
            .HeadType = HeadType.SelectedItem
        End With
        IDS.Data.Hardware.Camera.DotQCEnable = cbEnableGlobalQC.Checked
        IDS.Data.SaveData()

    End Sub

    Private Sub Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveData()
    End Sub

    Public Sub RevertData(Optional ByVal hideExit As Boolean = False)

        ButtonExit.Visible = Not (hideExit)

        IDS.Data.OpenData()
        If IDS.Data.Hardware.Dispenser.CurrentHeads = 1 Then
            CurrentHeads.Enabled = False
        Else
            CurrentHeads.Enabled = True
        End If

        CurrentHeads.SelectedItem = "Left Head"

        With IDS.Data.Hardware.Dispenser.Left
            MaterialAirPressure.Value = .MaterialAirPressure
            SuckbackPressure.Value = .SuckbackPressure
            RPM.Text = .RPM
            RetractTime.Text = .RetractTime
            RetractDelay.Text = .RetractDelay
            Pulse.Text = .Pulse
            pause.Text = .Pause
            Count.Text = .Count
            If HeadType.SelectedItem = "Auger Valve" Then
                AugerTemperature.Value = .ValveTemperature
            ElseIf HeadType.SelectedItem = "Slider Valve" Or HeadType.SelectedItem = "Jetting Valve" Then
                SliderOrJettingTemperature.Value = .ValveTemperature
            End If
            NeedleTipLength.SelectedItem = .NeedleTipLength
            NeedleColor.SelectedItem = .NeedleColor
            MaterialInfo.Text = .MaterialInfo
            AutoPurgingDuration.Value = .AutoPurgingDuration
            AutoCleaningDuration.Value = .AutoCleaningDuration
            AutoPurgingIntervalHours.Value = Math.Floor(.AutoPurgingInterval / 60)
            AutoPurgingIntervalMinutes.Value = .AutoPurgingInterval - Math.Floor(.AutoPurgingInterval / 60) * 60
            PotLifeDurationHours.Value = Math.Floor(.PotLifeDuration / 60)
            PotLifeDurationMinutes.Value = .PotLifeDuration - Math.Floor(.PotLifeDuration / 60) * 60
            PotLifeOption.Checked = .PotLifeOption
            AutoPurgingOption.Checked = .AutoPurgingOption
            HeadType.SelectedItem = .HeadType
        End With
        cbEnableGlobalQC.Checked = IDS.Data.Hardware.Camera.DotQCEnable
    End Sub

    Public Sub Revert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub

    Private Sub HeadType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HeadType.SelectedIndexChanged
        If HeadType.Text = "Jetting Valve" Then
            HideSuckback()
            HideClean()
            HideSyringeInfo()
            Auger.Hide()
            SliderOrJetting.Show()
        ElseIf HeadType.Text = "Auger Valve" Then
            ShowSuckback()
            ShowClean()
            ShowSyringeInfo()
            SliderOrJetting.Hide()
            Auger.Show()
        ElseIf HeadType.Text = "Slider Valve" Then
            ShowSuckback()
            ShowClean()
            ShowSyringeInfo()
            Auger.Hide()
            SliderOrJetting.Show()
        ElseIf HeadType.Text = "Time Pressure Valve" Then
            ShowSuckback()
            ShowClean()
            ShowSyringeInfo()
            Auger.Hide()
            SliderOrJetting.Hide()
        ElseIf HeadType.Text = "Time Pressure Syringe" Then
            ShowSuckback()
            ShowClean()
            ShowSyringeInfo()
            Auger.Hide()
            SliderOrJetting.Hide()
        End If
    End Sub

#Region "GUI show/hide functions"
    Private Sub ShowSuckback()
        SuckbackPressureLabel1.Visible = True
        SuckbackPressureLabel2.Visible = True
        SuckbackPressure.Visible = True
    End Sub
    Private Sub HideSuckback()
        SuckbackPressureLabel1.Visible = False
        SuckbackPressureLabel2.Visible = False
        SuckbackPressure.Visible = False
    End Sub
    Private Sub ShowClean()
        AutoCleaningDurationLabel1.Visible = True
        AutoCleaningDurationLabel2.Visible = True
        AutoCleaningDuration.Visible = True
    End Sub
    Private Sub HideClean()
        AutoCleaningDurationLabel1.Visible = False
        AutoCleaningDurationLabel2.Visible = False
        AutoCleaningDuration.Visible = False
    End Sub
    Private Sub ShowSyringeInfo()
        TipLengthLabel.Visible = True
        NeedleTipLength.Visible = True
        NeedleColorLabel.Visible = True
        NeedleColor.Visible = True
    End Sub
    Private Sub HideSyringeInfo()
        TipLengthLabel.Visible = False
        NeedleTipLength.Visible = False
        NeedleColorLabel.Visible = False
        NeedleColor.Visible = False
    End Sub
    Private Sub ShowAutoPurging()
        AutoPurgingDurationLabel1.Visible = True
        AutoPurgingDurationLabel2.Visible = True
        AutoPurgingIntervalLabel1.Visible = True
        AutoPurgingIntervalLabel2.Visible = True
        AutoPurgingIntervalLabel3.Visible = True
        AutoCleaningDurationLabel1.Visible = True
        AutoCleaningDurationLabel2.Visible = True
        AutoPurgingIntervalHours.Visible = True
        AutoPurgingIntervalMinutes.Visible = True
        AutoCleaningDuration.Visible = True
        AutoPurgingDuration.Visible = True
    End Sub
    Private Sub HideAutoPurging()
        AutoPurgingDurationLabel1.Visible = False
        AutoPurgingDurationLabel2.Visible = False
        AutoPurgingIntervalLabel1.Visible = False
        AutoPurgingIntervalLabel2.Visible = False
        AutoPurgingIntervalLabel3.Visible = False
        AutoCleaningDurationLabel1.Visible = False
        AutoCleaningDurationLabel2.Visible = False
        AutoPurgingIntervalHours.Visible = False
        AutoPurgingIntervalMinutes.Visible = False
        AutoCleaningDuration.Visible = False
        AutoPurgingDuration.Visible = False
    End Sub
    Private Sub ShowPotLife()
        PotLifeDurationLabel1.Visible = True
        PotLifeDurationLabel2.Visible = True
        PotLifeDurationLabel3.Visible = True
        PotLifeDurationHours.Visible = True
        PotLifeDurationMinutes.Visible = True
    End Sub
    Private Sub HidePotLife()
        PotLifeDurationLabel1.Visible = False
        PotLifeDurationLabel2.Visible = False
        PotLifeDurationLabel3.Visible = False
        PotLifeDurationHours.Visible = False
        PotLifeDurationMinutes.Visible = False
    End Sub
#End Region

    Private Sub AutoPurgingOption_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoPurgingOption.CheckStateChanged
        If AutoPurgingOption.Checked Then
            ShowAutoPurging()
            If HeadType.Text = "Jetting Valve" Then HideClean()
        Else
            HideAutoPurging()
        End If
    End Sub

    Private Sub PotLifeOption_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PotLifeOption.CheckedChanged
        If PotLifeOption.Checked Then
            ShowPotLife()
        Else
            HidePotLife()
        End If
    End Sub

    Private Sub Download_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Download.Click
        Dim SendData, SendVal As String
        Try
            If HeadType.SelectedItem.ToString = "Auger Valve" Then
                Dispenser.DownloadAugerRPM(RPM.Value, RetractTime.Value, RetractDelay.Value)
            ElseIf HeadType.SelectedItem.ToString = "Slider Valve" Or HeadType.SelectedItem.ToString = "Jetting Valve" Then
                Dispenser.DownloadJettingParameters(Pulse.Value, pause.Value, Count.Value, SliderOrJettingTemperature.Value)
            End If
            DownloadMaterialAirPressure(MaterialAirPressure.Text, SuckbackPressure.Text)
            IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure = MaterialAirPressure.Value
            IDS.Data.Hardware.Dispenser.Left.SuckbackPressure = SuckbackPressure.Value
            SaveData()
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

    End Sub

    Public Sub DownloadMaterialAirPressure(ByVal MaterialAirPressure As Double, ByVal SuckbackPressure As Double)
        IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(0, MaterialAirPressure, -(SuckbackPressure + 1))
    End Sub

    Public Sub DownloadDispenserSettings(ByVal CurrentHead As String)

        Dim SendData, SendVal As String
        Dim MaterialAirPressure, SuckbackPressure As Double
        Dim RPM, RetractTime, RetractDelay, Pulse, Pause, Count, Temperature As String
        Dim HeadType As String

        With IDS.Data.Hardware.Dispenser.Left
            MaterialAirPressure = .MaterialAirPressure
            SuckbackPressure = .SuckbackPressure
            RPM = .RPM
            RetractTime = .RetractTime
            RetractDelay = .RetractDelay
            Pulse = .Pulse
            Pause = .Pause
            Count = .Count
            Temperature = .ValveTemperature
            HeadType = .HeadType
        End With

        If HeadType = "Auger Valve" Then
            Dispenser.DownloadAugerRPM(RPM, RetractTime, RetractDelay)
        ElseIf HeadType = "Slider Valve" Or HeadType = "Jetting Valve" Then
            Dispenser.DownloadJettingParameters(Pulse, Pause, Count, Temperature)
        End If

        'for psi: IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(1, pressure, suckback)
        'we use bar as the default value. see above if you want to change.
        DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)

    End Sub

    Public Sub DownloadAugerRPM(ByVal RPM As Double, ByVal RetractTime As Double, ByVal RetractDelay As Double)
        'With IDS.Data.Hardware.Dispenser.Left
        Dispenser.DownloadAugerRPM(RPM, RetractTime, RetractDelay)
        'End With
    End Sub
    Public Sub DownloadJettingParameters(ByVal PulseOnDuration As Double, ByVal PulseOffDuration As Double, ByVal noOfDispense As Double, ByVal temperature As Double)
        Dispenser.DownloadJettingParameters(PulseOnDuration, PulseOffDuration, noOfDispense, temperature)
    End Sub
    Public Sub DownloadJettingParameters(ByVal PulseDuration As Double)
        With IDS.Data.Hardware.Dispenser.Left
            'Dispenser.DownloadJettingParameters(PulseDuration, .RetractTime, .RetractDelay, .ValveTemperature)
            Dispenser.DownloadJettingParameters(PulseDuration, .Pause, .Count, .ValveTemperature)
        End With
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub

    Private Sub btZeroPressure_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btZeroPressure.Click
        DownloadMaterialAirPressure(0, 0)
    End Sub
End Class
