Imports DLL_Export_Service

Public Class VolumeCalibrationSettings
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
    Friend WithEvents Label17 As System.Windows.Forms.Label


    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents gpbDualHead As System.Windows.Forms.GroupBox
    Public WithEvents chkDualHead As System.Windows.Forms.CheckBox
    Friend WithEvents DesiredWeight As System.Windows.Forms.NumericUpDown
    Friend WithEvents LabelSetTolerance As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents NumberOfAttempts As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents DispenserType As System.Windows.Forms.Label
    Friend WithEvents ButtonCalibrate As System.Windows.Forms.Button
    Friend WithEvents SetupMaterialAirPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents SetupDispenseDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents Tolerance As System.Windows.Forms.NumericUpDown
    Friend WithEvents AdjustedDispenseDuration As System.Windows.Forms.NumericUpDown
    Friend WithEvents AdjustedMaterialAirPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents AdjustedRPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents CalibrationType As System.Windows.Forms.ComboBox
    Friend WithEvents LabelAdjustedDispenseDuration1 As System.Windows.Forms.Label
    Friend WithEvents LabelAdjustedMaterialAirPressure1 As System.Windows.Forms.Label
    Friend WithEvents LabelAdjustedDispenseDuration2 As System.Windows.Forms.Label
    Friend WithEvents LabelAdjustedRPM1 As System.Windows.Forms.Label
    Friend WithEvents LabelAdjustedMaterialAirPressure2 As System.Windows.Forms.Label
    Friend WithEvents BoxSettings As System.Windows.Forms.GroupBox
    Friend WithEvents LabelSetupMaterialAirPressure1 As System.Windows.Forms.Label
    Friend WithEvents LabelDesiredWeight2 As System.Windows.Forms.Label
    Friend WithEvents LabelSetupDispenseDuration2 As System.Windows.Forms.Label
    Friend WithEvents LabelDesiredWeight1 As System.Windows.Forms.Label
    Friend WithEvents LabelSetupDispenseDuration1 As System.Windows.Forms.Label
    Friend WithEvents LabelSetupMaterialAirPressure2 As System.Windows.Forms.Label
    Friend WithEvents SetupRPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents LabelSetupRPM As System.Windows.Forms.Label
    Friend WithEvents Status As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ButtonTeachCalibrate As System.Windows.Forms.Button
    Friend WithEvents WeighingScaleButton As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents RetractDelayValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractSpeedValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents RetractHeightValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents LabelPressureStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelPressureStep1 As System.Windows.Forms.Label
    Friend WithEvents LabelDurationStep1 As System.Windows.Forms.Label
    Friend WithEvents LabelDurationStep2 As System.Windows.Forms.Label
    Friend WithEvents LabelRPMStep1 As System.Windows.Forms.Label
    Friend WithEvents MaterialAirPressureStepValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents RPMStepValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents DurationStepValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents rtbResult As System.Windows.Forms.RichTextBox
    Friend WithEvents jettingPulseOffValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents jettingNumOfDispense As System.Windows.Forms.NumericUpDown
    Friend WithEvents jettingpulseOnValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbJettingPulseOff As System.Windows.Forms.Label
    Friend WithEvents lbJettingNoOfDispense As System.Windows.Forms.Label
    Friend WithEvents lbJettingPulseOn As System.Windows.Forms.Label
    Friend WithEvents lbValveTemp As System.Windows.Forms.Label
    Friend WithEvents valveTempValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents AugerRetractTimeValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbAugerRetractTime As System.Windows.Forms.Label
    Friend WithEvents AugerValveRetractDelay As System.Windows.Forms.NumericUpDown
    Friend WithEvents lbAugerRetractDelay As System.Windows.Forms.Label


    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VolumeCalibrationSettings))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.BoxSettings = New System.Windows.Forms.GroupBox
        Me.AugerValveRetractDelay = New System.Windows.Forms.NumericUpDown
        Me.lbAugerRetractDelay = New System.Windows.Forms.Label
        Me.AugerRetractTimeValue = New System.Windows.Forms.NumericUpDown
        Me.lbAugerRetractTime = New System.Windows.Forms.Label
        Me.lbValveTemp = New System.Windows.Forms.Label
        Me.valveTempValue = New System.Windows.Forms.NumericUpDown
        Me.lbJettingPulseOn = New System.Windows.Forms.Label
        Me.jettingpulseOnValue = New System.Windows.Forms.NumericUpDown
        Me.lbJettingNoOfDispense = New System.Windows.Forms.Label
        Me.jettingNumOfDispense = New System.Windows.Forms.NumericUpDown
        Me.lbJettingPulseOff = New System.Windows.Forms.Label
        Me.jettingPulseOffValue = New System.Windows.Forms.NumericUpDown
        Me.SetupDispenseDuration = New System.Windows.Forms.NumericUpDown
        Me.SetupMaterialAirPressure = New System.Windows.Forms.NumericUpDown
        Me.Tolerance = New System.Windows.Forms.NumericUpDown
        Me.DesiredWeight = New System.Windows.Forms.NumericUpDown
        Me.NumberOfAttempts = New System.Windows.Forms.NumericUpDown
        Me.Label49 = New System.Windows.Forms.Label
        Me.SetupRPM = New System.Windows.Forms.NumericUpDown
        Me.LabelSetupMaterialAirPressure1 = New System.Windows.Forms.Label
        Me.LabelDesiredWeight2 = New System.Windows.Forms.Label
        Me.LabelSetTolerance = New System.Windows.Forms.Label
        Me.LabelSetupDispenseDuration2 = New System.Windows.Forms.Label
        Me.LabelSetupRPM = New System.Windows.Forms.Label
        Me.LabelDesiredWeight1 = New System.Windows.Forms.Label
        Me.LabelSetupDispenseDuration1 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.CalibrationType = New System.Windows.Forms.ComboBox
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.LabelSetupMaterialAirPressure2 = New System.Windows.Forms.Label
        Me.LabelAdjustedDispenseDuration1 = New System.Windows.Forms.Label
        Me.LabelAdjustedMaterialAirPressure1 = New System.Windows.Forms.Label
        Me.LabelAdjustedDispenseDuration2 = New System.Windows.Forms.Label
        Me.LabelAdjustedRPM1 = New System.Windows.Forms.Label
        Me.AdjustedMaterialAirPressure = New System.Windows.Forms.NumericUpDown
        Me.AdjustedDispenseDuration = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.RetractDelayValue = New System.Windows.Forms.NumericUpDown
        Me.RetractSpeedValue = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.RetractHeightValue = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.LabelAdjustedMaterialAirPressure2 = New System.Windows.Forms.Label
        Me.LabelDurationStep2 = New System.Windows.Forms.Label
        Me.LabelRPMStep1 = New System.Windows.Forms.Label
        Me.LabelPressureStep1 = New System.Windows.Forms.Label
        Me.LabelPressureStep2 = New System.Windows.Forms.Label
        Me.LabelDurationStep1 = New System.Windows.Forms.Label
        Me.MaterialAirPressureStepValue = New System.Windows.Forms.NumericUpDown
        Me.RPMStepValue = New System.Windows.Forms.NumericUpDown
        Me.DurationStepValue = New System.Windows.Forms.NumericUpDown
        Me.AdjustedRPM = New System.Windows.Forms.NumericUpDown
        Me.WeighingScaleButton = New System.Windows.Forms.Button
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.ButtonTeachCalibrate = New System.Windows.Forms.Button
        Me.DispenserType = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.rtbResult = New System.Windows.Forms.RichTextBox
        Me.Status = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.gpbDualHead = New System.Windows.Forms.GroupBox
        Me.chkDualHead = New System.Windows.Forms.CheckBox
        Me.PanelToBeAdded.SuspendLayout()
        Me.BoxSettings.SuspendLayout()
        CType(Me.AugerValveRetractDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AugerRetractTimeValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.valveTempValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.jettingpulseOnValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.jettingNumOfDispense, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.jettingPulseOffValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SetupDispenseDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SetupMaterialAirPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tolerance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DesiredWeight, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumberOfAttempts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SetupRPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AdjustedMaterialAirPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AdjustedDispenseDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractDelayValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractSpeedValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractHeightValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaterialAirPressureStepValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RPMStepValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DurationStepValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AdjustedRPM, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpbDualHead.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.Label40)
        Me.PanelToBeAdded.Controls.Add(Me.BoxSettings)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonCalibrate)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonTeachCalibrate)
        Me.PanelToBeAdded.Controls.Add(Me.DispenserType)
        Me.PanelToBeAdded.Controls.Add(Me.Label2)
        Me.PanelToBeAdded.Controls.Add(Me.rtbResult)
        Me.PanelToBeAdded.Controls.Add(Me.Status)
        Me.PanelToBeAdded.Controls.Add(Me.Label1)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 0
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(400, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 47
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(272, 32)
        Me.Label17.TabIndex = 35
        Me.Label17.Text = "Volume Calibration Setup"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(8, 32)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(136, 23)
        Me.Label40.TabIndex = 3
        Me.Label40.Text = "Dispenser Type:"
        '
        'BoxSettings
        '
        Me.BoxSettings.Controls.Add(Me.AugerValveRetractDelay)
        Me.BoxSettings.Controls.Add(Me.lbAugerRetractDelay)
        Me.BoxSettings.Controls.Add(Me.AugerRetractTimeValue)
        Me.BoxSettings.Controls.Add(Me.lbAugerRetractTime)
        Me.BoxSettings.Controls.Add(Me.lbValveTemp)
        Me.BoxSettings.Controls.Add(Me.valveTempValue)
        Me.BoxSettings.Controls.Add(Me.lbJettingPulseOn)
        Me.BoxSettings.Controls.Add(Me.jettingpulseOnValue)
        Me.BoxSettings.Controls.Add(Me.lbJettingNoOfDispense)
        Me.BoxSettings.Controls.Add(Me.jettingNumOfDispense)
        Me.BoxSettings.Controls.Add(Me.lbJettingPulseOff)
        Me.BoxSettings.Controls.Add(Me.jettingPulseOffValue)
        Me.BoxSettings.Controls.Add(Me.SetupDispenseDuration)
        Me.BoxSettings.Controls.Add(Me.SetupMaterialAirPressure)
        Me.BoxSettings.Controls.Add(Me.Tolerance)
        Me.BoxSettings.Controls.Add(Me.DesiredWeight)
        Me.BoxSettings.Controls.Add(Me.NumberOfAttempts)
        Me.BoxSettings.Controls.Add(Me.Label49)
        Me.BoxSettings.Controls.Add(Me.SetupRPM)
        Me.BoxSettings.Controls.Add(Me.LabelSetupMaterialAirPressure1)
        Me.BoxSettings.Controls.Add(Me.LabelDesiredWeight2)
        Me.BoxSettings.Controls.Add(Me.LabelSetTolerance)
        Me.BoxSettings.Controls.Add(Me.LabelSetupDispenseDuration2)
        Me.BoxSettings.Controls.Add(Me.LabelSetupRPM)
        Me.BoxSettings.Controls.Add(Me.LabelDesiredWeight1)
        Me.BoxSettings.Controls.Add(Me.LabelSetupDispenseDuration1)
        Me.BoxSettings.Controls.Add(Me.Label45)
        Me.BoxSettings.Controls.Add(Me.CalibrationType)
        Me.BoxSettings.Controls.Add(Me.ButtonRevert)
        Me.BoxSettings.Controls.Add(Me.ButtonSave)
        Me.BoxSettings.Controls.Add(Me.LabelSetupMaterialAirPressure2)
        Me.BoxSettings.Controls.Add(Me.LabelAdjustedDispenseDuration1)
        Me.BoxSettings.Controls.Add(Me.LabelAdjustedMaterialAirPressure1)
        Me.BoxSettings.Controls.Add(Me.LabelAdjustedDispenseDuration2)
        Me.BoxSettings.Controls.Add(Me.LabelAdjustedRPM1)
        Me.BoxSettings.Controls.Add(Me.AdjustedMaterialAirPressure)
        Me.BoxSettings.Controls.Add(Me.AdjustedDispenseDuration)
        Me.BoxSettings.Controls.Add(Me.Label3)
        Me.BoxSettings.Controls.Add(Me.Label4)
        Me.BoxSettings.Controls.Add(Me.Label6)
        Me.BoxSettings.Controls.Add(Me.RetractDelayValue)
        Me.BoxSettings.Controls.Add(Me.RetractSpeedValue)
        Me.BoxSettings.Controls.Add(Me.Label5)
        Me.BoxSettings.Controls.Add(Me.RetractHeightValue)
        Me.BoxSettings.Controls.Add(Me.Label7)
        Me.BoxSettings.Controls.Add(Me.Label8)
        Me.BoxSettings.Controls.Add(Me.LabelAdjustedMaterialAirPressure2)
        Me.BoxSettings.Controls.Add(Me.LabelDurationStep2)
        Me.BoxSettings.Controls.Add(Me.LabelRPMStep1)
        Me.BoxSettings.Controls.Add(Me.LabelPressureStep1)
        Me.BoxSettings.Controls.Add(Me.LabelPressureStep2)
        Me.BoxSettings.Controls.Add(Me.LabelDurationStep1)
        Me.BoxSettings.Controls.Add(Me.MaterialAirPressureStepValue)
        Me.BoxSettings.Controls.Add(Me.RPMStepValue)
        Me.BoxSettings.Controls.Add(Me.DurationStepValue)
        Me.BoxSettings.Controls.Add(Me.AdjustedRPM)
        Me.BoxSettings.Controls.Add(Me.WeighingScaleButton)
        Me.BoxSettings.Location = New System.Drawing.Point(16, 64)
        Me.BoxSettings.Name = "BoxSettings"
        Me.BoxSettings.Size = New System.Drawing.Size(480, 616)
        Me.BoxSettings.TabIndex = 49
        Me.BoxSettings.TabStop = False
        Me.BoxSettings.Text = "Settings"
        '
        'AugerValveRetractDelay
        '
        Me.AugerValveRetractDelay.DecimalPlaces = 2
        Me.AugerValveRetractDelay.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.AugerValveRetractDelay.Location = New System.Drawing.Point(268, 448)
        Me.AugerValveRetractDelay.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.AugerValveRetractDelay.Name = "AugerValveRetractDelay"
        Me.AugerValveRetractDelay.Size = New System.Drawing.Size(72, 27)
        Me.AugerValveRetractDelay.TabIndex = 63
        '
        'lbAugerRetractDelay
        '
        Me.lbAugerRetractDelay.Location = New System.Drawing.Point(72, 448)
        Me.lbAugerRetractDelay.Name = "lbAugerRetractDelay"
        Me.lbAugerRetractDelay.Size = New System.Drawing.Size(200, 23)
        Me.lbAugerRetractDelay.TabIndex = 62
        Me.lbAugerRetractDelay.Text = "Valve Retract Delay (s)"
        '
        'AugerRetractTimeValue
        '
        Me.AugerRetractTimeValue.DecimalPlaces = 2
        Me.AugerRetractTimeValue.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.AugerRetractTimeValue.Location = New System.Drawing.Point(268, 416)
        Me.AugerRetractTimeValue.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.AugerRetractTimeValue.Name = "AugerRetractTimeValue"
        Me.AugerRetractTimeValue.Size = New System.Drawing.Size(72, 27)
        Me.AugerRetractTimeValue.TabIndex = 61
        '
        'lbAugerRetractTime
        '
        Me.lbAugerRetractTime.Location = New System.Drawing.Point(76, 416)
        Me.lbAugerRetractTime.Name = "lbAugerRetractTime"
        Me.lbAugerRetractTime.Size = New System.Drawing.Size(188, 23)
        Me.lbAugerRetractTime.TabIndex = 60
        Me.lbAugerRetractTime.Text = "Valve Retract Time (s)"
        '
        'lbValveTemp
        '
        Me.lbValveTemp.Location = New System.Drawing.Point(76, 544)
        Me.lbValveTemp.Name = "lbValveTemp"
        Me.lbValveTemp.Size = New System.Drawing.Size(160, 23)
        Me.lbValveTemp.TabIndex = 58
        Me.lbValveTemp.Text = "Valve Temperature"
        '
        'valveTempValue
        '
        Me.valveTempValue.DecimalPlaces = 3
        Me.valveTempValue.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.valveTempValue.Location = New System.Drawing.Point(268, 544)
        Me.valveTempValue.Name = "valveTempValue"
        Me.valveTempValue.Size = New System.Drawing.Size(72, 27)
        Me.valveTempValue.TabIndex = 59
        '
        'lbJettingPulseOn
        '
        Me.lbJettingPulseOn.Location = New System.Drawing.Point(76, 416)
        Me.lbJettingPulseOn.Name = "lbJettingPulseOn"
        Me.lbJettingPulseOn.Size = New System.Drawing.Size(188, 23)
        Me.lbJettingPulseOn.TabIndex = 55
        Me.lbJettingPulseOn.Text = "Pulse On Duration(ms)"
        '
        'jettingpulseOnValue
        '
        Me.jettingpulseOnValue.Location = New System.Drawing.Point(268, 416)
        Me.jettingpulseOnValue.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.jettingpulseOnValue.Name = "jettingpulseOnValue"
        Me.jettingpulseOnValue.Size = New System.Drawing.Size(72, 27)
        Me.jettingpulseOnValue.TabIndex = 57
        '
        'lbJettingNoOfDispense
        '
        Me.lbJettingNoOfDispense.Location = New System.Drawing.Point(76, 512)
        Me.lbJettingNoOfDispense.Name = "lbJettingNoOfDispense"
        Me.lbJettingNoOfDispense.Size = New System.Drawing.Size(160, 23)
        Me.lbJettingNoOfDispense.TabIndex = 52
        Me.lbJettingNoOfDispense.Text = "No. Of Dispense"
        '
        'jettingNumOfDispense
        '
        Me.jettingNumOfDispense.Location = New System.Drawing.Point(268, 512)
        Me.jettingNumOfDispense.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.jettingNumOfDispense.Name = "jettingNumOfDispense"
        Me.jettingNumOfDispense.Size = New System.Drawing.Size(72, 27)
        Me.jettingNumOfDispense.TabIndex = 54
        Me.jettingNumOfDispense.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'lbJettingPulseOff
        '
        Me.lbJettingPulseOff.Location = New System.Drawing.Point(76, 480)
        Me.lbJettingPulseOff.Name = "lbJettingPulseOff"
        Me.lbJettingPulseOff.Size = New System.Drawing.Size(188, 23)
        Me.lbJettingPulseOff.TabIndex = 49
        Me.lbJettingPulseOff.Text = "Pulse Off Duration(ms)"
        '
        'jettingPulseOffValue
        '
        Me.jettingPulseOffValue.Location = New System.Drawing.Point(268, 480)
        Me.jettingPulseOffValue.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.jettingPulseOffValue.Name = "jettingPulseOffValue"
        Me.jettingPulseOffValue.Size = New System.Drawing.Size(72, 27)
        Me.jettingPulseOffValue.TabIndex = 51
        '
        'SetupDispenseDuration
        '
        Me.SetupDispenseDuration.DecimalPlaces = 2
        Me.SetupDispenseDuration.Location = New System.Drawing.Point(268, 96)
        Me.SetupDispenseDuration.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.SetupDispenseDuration.Name = "SetupDispenseDuration"
        Me.SetupDispenseDuration.Size = New System.Drawing.Size(72, 27)
        Me.SetupDispenseDuration.TabIndex = 17
        Me.SetupDispenseDuration.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'SetupMaterialAirPressure
        '
        Me.SetupMaterialAirPressure.DecimalPlaces = 1
        Me.SetupMaterialAirPressure.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.SetupMaterialAirPressure.Location = New System.Drawing.Point(268, 128)
        Me.SetupMaterialAirPressure.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.SetupMaterialAirPressure.Name = "SetupMaterialAirPressure"
        Me.SetupMaterialAirPressure.Size = New System.Drawing.Size(72, 27)
        Me.SetupMaterialAirPressure.TabIndex = 17
        Me.SetupMaterialAirPressure.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Tolerance
        '
        Me.Tolerance.DecimalPlaces = 1
        Me.Tolerance.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.Tolerance.Location = New System.Drawing.Point(268, 192)
        Me.Tolerance.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.Tolerance.Name = "Tolerance"
        Me.Tolerance.Size = New System.Drawing.Size(72, 27)
        Me.Tolerance.TabIndex = 17
        Me.Tolerance.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'DesiredWeight
        '
        Me.DesiredWeight.DecimalPlaces = 1
        Me.DesiredWeight.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.DesiredWeight.Location = New System.Drawing.Point(268, 64)
        Me.DesiredWeight.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.DesiredWeight.Name = "DesiredWeight"
        Me.DesiredWeight.Size = New System.Drawing.Size(72, 27)
        Me.DesiredWeight.TabIndex = 17
        Me.DesiredWeight.ThousandsSeparator = True
        Me.DesiredWeight.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'NumberOfAttempts
        '
        Me.NumberOfAttempts.Location = New System.Drawing.Point(268, 256)
        Me.NumberOfAttempts.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumberOfAttempts.Name = "NumberOfAttempts"
        Me.NumberOfAttempts.Size = New System.Drawing.Size(72, 27)
        Me.NumberOfAttempts.TabIndex = 17
        Me.NumberOfAttempts.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(76, 256)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(168, 23)
        Me.Label49.TabIndex = 4
        Me.Label49.Text = "Number of Attempts"
        '
        'SetupRPM
        '
        Me.SetupRPM.DecimalPlaces = 1
        Me.SetupRPM.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.SetupRPM.Location = New System.Drawing.Point(268, 160)
        Me.SetupRPM.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.SetupRPM.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.SetupRPM.Name = "SetupRPM"
        Me.SetupRPM.Size = New System.Drawing.Size(72, 27)
        Me.SetupRPM.TabIndex = 17
        Me.SetupRPM.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'LabelSetupMaterialAirPressure1
        '
        Me.LabelSetupMaterialAirPressure1.Location = New System.Drawing.Point(76, 128)
        Me.LabelSetupMaterialAirPressure1.Name = "LabelSetupMaterialAirPressure1"
        Me.LabelSetupMaterialAirPressure1.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupMaterialAirPressure1.TabIndex = 3
        Me.LabelSetupMaterialAirPressure1.Text = "Dispense Pressure"
        '
        'LabelDesiredWeight2
        '
        Me.LabelDesiredWeight2.Location = New System.Drawing.Point(340, 64)
        Me.LabelDesiredWeight2.Name = "LabelDesiredWeight2"
        Me.LabelDesiredWeight2.Size = New System.Drawing.Size(32, 24)
        Me.LabelDesiredWeight2.TabIndex = 11
        Me.LabelDesiredWeight2.Text = "mg"
        '
        'LabelSetTolerance
        '
        Me.LabelSetTolerance.Location = New System.Drawing.Point(76, 192)
        Me.LabelSetTolerance.Name = "LabelSetTolerance"
        Me.LabelSetTolerance.Size = New System.Drawing.Size(88, 23)
        Me.LabelSetTolerance.TabIndex = 8
        Me.LabelSetTolerance.Text = "Tolerance"
        '
        'LabelSetupDispenseDuration2
        '
        Me.LabelSetupDispenseDuration2.Location = New System.Drawing.Point(340, 96)
        Me.LabelSetupDispenseDuration2.Name = "LabelSetupDispenseDuration2"
        Me.LabelSetupDispenseDuration2.Size = New System.Drawing.Size(32, 24)
        Me.LabelSetupDispenseDuration2.TabIndex = 7
        Me.LabelSetupDispenseDuration2.Text = "ms"
        '
        'LabelSetupRPM
        '
        Me.LabelSetupRPM.Location = New System.Drawing.Point(76, 160)
        Me.LabelSetupRPM.Name = "LabelSetupRPM"
        Me.LabelSetupRPM.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupRPM.TabIndex = 3
        Me.LabelSetupRPM.Text = "Dispense RPM"
        '
        'LabelDesiredWeight1
        '
        Me.LabelDesiredWeight1.Location = New System.Drawing.Point(76, 64)
        Me.LabelDesiredWeight1.Name = "LabelDesiredWeight1"
        Me.LabelDesiredWeight1.Size = New System.Drawing.Size(128, 23)
        Me.LabelDesiredWeight1.TabIndex = 3
        Me.LabelDesiredWeight1.Text = "Desired Weight"
        '
        'LabelSetupDispenseDuration1
        '
        Me.LabelSetupDispenseDuration1.Location = New System.Drawing.Point(76, 96)
        Me.LabelSetupDispenseDuration1.Name = "LabelSetupDispenseDuration1"
        Me.LabelSetupDispenseDuration1.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupDispenseDuration1.TabIndex = 4
        Me.LabelSetupDispenseDuration1.Text = "Dispense Duration"
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(340, 192)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(64, 24)
        Me.Label45.TabIndex = 7
        Me.Label45.Text = "+/- mg"
        '
        'CalibrationType
        '
        Me.CalibrationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CalibrationType.Items.AddRange(New Object() {"By Desired Weight", "By Current Parameters"})
        Me.CalibrationType.Location = New System.Drawing.Point(80, 16)
        Me.CalibrationType.Name = "CalibrationType"
        Me.CalibrationType.Size = New System.Drawing.Size(304, 28)
        Me.CalibrationType.TabIndex = 48
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Location = New System.Drawing.Point(248, 576)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(80, 40)
        Me.ButtonRevert.TabIndex = 6
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(160, 576)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(80, 40)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        '
        'LabelSetupMaterialAirPressure2
        '
        Me.LabelSetupMaterialAirPressure2.Location = New System.Drawing.Point(340, 128)
        Me.LabelSetupMaterialAirPressure2.Name = "LabelSetupMaterialAirPressure2"
        Me.LabelSetupMaterialAirPressure2.Size = New System.Drawing.Size(40, 24)
        Me.LabelSetupMaterialAirPressure2.TabIndex = 6
        Me.LabelSetupMaterialAirPressure2.Text = "Bar"
        '
        'LabelAdjustedDispenseDuration1
        '
        Me.LabelAdjustedDispenseDuration1.Enabled = False
        Me.LabelAdjustedDispenseDuration1.Location = New System.Drawing.Point(76, 384)
        Me.LabelAdjustedDispenseDuration1.Name = "LabelAdjustedDispenseDuration1"
        Me.LabelAdjustedDispenseDuration1.Size = New System.Drawing.Size(160, 23)
        Me.LabelAdjustedDispenseDuration1.TabIndex = 4
        Me.LabelAdjustedDispenseDuration1.Text = "Adjusted Duration"
        Me.LabelAdjustedDispenseDuration1.Visible = False
        '
        'LabelAdjustedMaterialAirPressure1
        '
        Me.LabelAdjustedMaterialAirPressure1.Enabled = False
        Me.LabelAdjustedMaterialAirPressure1.Location = New System.Drawing.Point(76, 384)
        Me.LabelAdjustedMaterialAirPressure1.Name = "LabelAdjustedMaterialAirPressure1"
        Me.LabelAdjustedMaterialAirPressure1.Size = New System.Drawing.Size(160, 23)
        Me.LabelAdjustedMaterialAirPressure1.TabIndex = 3
        Me.LabelAdjustedMaterialAirPressure1.Text = "Adjusted Pressure"
        Me.LabelAdjustedMaterialAirPressure1.Visible = False
        '
        'LabelAdjustedDispenseDuration2
        '
        Me.LabelAdjustedDispenseDuration2.Enabled = False
        Me.LabelAdjustedDispenseDuration2.Location = New System.Drawing.Point(340, 384)
        Me.LabelAdjustedDispenseDuration2.Name = "LabelAdjustedDispenseDuration2"
        Me.LabelAdjustedDispenseDuration2.Size = New System.Drawing.Size(32, 24)
        Me.LabelAdjustedDispenseDuration2.TabIndex = 7
        Me.LabelAdjustedDispenseDuration2.Text = "ms"
        Me.LabelAdjustedDispenseDuration2.Visible = False
        '
        'LabelAdjustedRPM1
        '
        Me.LabelAdjustedRPM1.Enabled = False
        Me.LabelAdjustedRPM1.Location = New System.Drawing.Point(76, 384)
        Me.LabelAdjustedRPM1.Name = "LabelAdjustedRPM1"
        Me.LabelAdjustedRPM1.Size = New System.Drawing.Size(160, 23)
        Me.LabelAdjustedRPM1.TabIndex = 3
        Me.LabelAdjustedRPM1.Text = "Adjusted RPM"
        Me.LabelAdjustedRPM1.Visible = False
        '
        'AdjustedMaterialAirPressure
        '
        Me.AdjustedMaterialAirPressure.DecimalPlaces = 2
        Me.AdjustedMaterialAirPressure.Enabled = False
        Me.AdjustedMaterialAirPressure.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.AdjustedMaterialAirPressure.Location = New System.Drawing.Point(268, 384)
        Me.AdjustedMaterialAirPressure.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AdjustedMaterialAirPressure.Name = "AdjustedMaterialAirPressure"
        Me.AdjustedMaterialAirPressure.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedMaterialAirPressure.TabIndex = 17
        Me.AdjustedMaterialAirPressure.Visible = False
        '
        'AdjustedDispenseDuration
        '
        Me.AdjustedDispenseDuration.Enabled = False
        Me.AdjustedDispenseDuration.Location = New System.Drawing.Point(268, 384)
        Me.AdjustedDispenseDuration.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.AdjustedDispenseDuration.Name = "AdjustedDispenseDuration"
        Me.AdjustedDispenseDuration.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedDispenseDuration.TabIndex = 17
        Me.AdjustedDispenseDuration.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(76, 320)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Retract Speed"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(76, 288)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 23)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Retract Delay"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(340, 288)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 24)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "ms"
        '
        'RetractDelayValue
        '
        Me.RetractDelayValue.Location = New System.Drawing.Point(268, 288)
        Me.RetractDelayValue.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.RetractDelayValue.Name = "RetractDelayValue"
        Me.RetractDelayValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractDelayValue.TabIndex = 17
        Me.RetractDelayValue.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'RetractSpeedValue
        '
        Me.RetractSpeedValue.Location = New System.Drawing.Point(268, 320)
        Me.RetractSpeedValue.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.RetractSpeedValue.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.RetractSpeedValue.Name = "RetractSpeedValue"
        Me.RetractSpeedValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractSpeedValue.TabIndex = 17
        Me.RetractSpeedValue.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(76, 352)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 23)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Retract Height"
        '
        'RetractHeightValue
        '
        Me.RetractHeightValue.Location = New System.Drawing.Point(268, 352)
        Me.RetractHeightValue.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.RetractHeightValue.Name = "RetractHeightValue"
        Me.RetractHeightValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractHeightValue.TabIndex = 17
        Me.RetractHeightValue.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(340, 352)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 24)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "mm"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(340, 320)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(52, 24)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "mm/s"
        '
        'LabelAdjustedMaterialAirPressure2
        '
        Me.LabelAdjustedMaterialAirPressure2.Enabled = False
        Me.LabelAdjustedMaterialAirPressure2.Location = New System.Drawing.Point(340, 384)
        Me.LabelAdjustedMaterialAirPressure2.Name = "LabelAdjustedMaterialAirPressure2"
        Me.LabelAdjustedMaterialAirPressure2.Size = New System.Drawing.Size(40, 24)
        Me.LabelAdjustedMaterialAirPressure2.TabIndex = 6
        Me.LabelAdjustedMaterialAirPressure2.Text = "Bar"
        Me.LabelAdjustedMaterialAirPressure2.Visible = False
        '
        'LabelDurationStep2
        '
        Me.LabelDurationStep2.Location = New System.Drawing.Point(340, 224)
        Me.LabelDurationStep2.Name = "LabelDurationStep2"
        Me.LabelDurationStep2.Size = New System.Drawing.Size(32, 24)
        Me.LabelDurationStep2.TabIndex = 7
        Me.LabelDurationStep2.Text = "ms"
        '
        'LabelRPMStep1
        '
        Me.LabelRPMStep1.Location = New System.Drawing.Point(76, 224)
        Me.LabelRPMStep1.Name = "LabelRPMStep1"
        Me.LabelRPMStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelRPMStep1.TabIndex = 3
        Me.LabelRPMStep1.Text = "RPM Step"
        '
        'LabelPressureStep1
        '
        Me.LabelPressureStep1.Location = New System.Drawing.Point(76, 224)
        Me.LabelPressureStep1.Name = "LabelPressureStep1"
        Me.LabelPressureStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelPressureStep1.TabIndex = 4
        Me.LabelPressureStep1.Text = "Pressure Step"
        '
        'LabelPressureStep2
        '
        Me.LabelPressureStep2.Location = New System.Drawing.Point(340, 224)
        Me.LabelPressureStep2.Name = "LabelPressureStep2"
        Me.LabelPressureStep2.Size = New System.Drawing.Size(40, 24)
        Me.LabelPressureStep2.TabIndex = 6
        Me.LabelPressureStep2.Text = "Bar"
        '
        'LabelDurationStep1
        '
        Me.LabelDurationStep1.Location = New System.Drawing.Point(76, 224)
        Me.LabelDurationStep1.Name = "LabelDurationStep1"
        Me.LabelDurationStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelDurationStep1.TabIndex = 3
        Me.LabelDurationStep1.Text = "Duration Step"
        '
        'MaterialAirPressureStepValue
        '
        Me.MaterialAirPressureStepValue.DecimalPlaces = 2
        Me.MaterialAirPressureStepValue.Location = New System.Drawing.Point(268, 224)
        Me.MaterialAirPressureStepValue.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.MaterialAirPressureStepValue.Name = "MaterialAirPressureStepValue"
        Me.MaterialAirPressureStepValue.Size = New System.Drawing.Size(72, 27)
        Me.MaterialAirPressureStepValue.TabIndex = 17
        Me.MaterialAirPressureStepValue.Value = New Decimal(New Integer() {1, 0, 0, 65536})
        '
        'RPMStepValue
        '
        Me.RPMStepValue.DecimalPlaces = 1
        Me.RPMStepValue.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.RPMStepValue.Location = New System.Drawing.Point(268, 224)
        Me.RPMStepValue.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RPMStepValue.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.RPMStepValue.Name = "RPMStepValue"
        Me.RPMStepValue.Size = New System.Drawing.Size(72, 27)
        Me.RPMStepValue.TabIndex = 17
        Me.RPMStepValue.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'DurationStepValue
        '
        Me.DurationStepValue.Location = New System.Drawing.Point(268, 224)
        Me.DurationStepValue.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.DurationStepValue.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.DurationStepValue.Name = "DurationStepValue"
        Me.DurationStepValue.Size = New System.Drawing.Size(72, 27)
        Me.DurationStepValue.TabIndex = 17
        Me.DurationStepValue.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'AdjustedRPM
        '
        Me.AdjustedRPM.DecimalPlaces = 1
        Me.AdjustedRPM.Enabled = False
        Me.AdjustedRPM.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.AdjustedRPM.Location = New System.Drawing.Point(268, 384)
        Me.AdjustedRPM.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AdjustedRPM.Name = "AdjustedRPM"
        Me.AdjustedRPM.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedRPM.TabIndex = 17
        Me.AdjustedRPM.Visible = False
        '
        'WeighingScaleButton
        '
        Me.WeighingScaleButton.Location = New System.Drawing.Point(400, 16)
        Me.WeighingScaleButton.Name = "WeighingScaleButton"
        Me.WeighingScaleButton.Size = New System.Drawing.Size(75, 50)
        Me.WeighingScaleButton.TabIndex = 50
        Me.WeighingScaleButton.Text = "Scale"
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.Location = New System.Drawing.Point(120, 688)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(128, 40)
        Me.ButtonCalibrate.TabIndex = 6
        Me.ButtonCalibrate.Text = "Calibrate"
        '
        'ButtonTeachCalibrate
        '
        Me.ButtonTeachCalibrate.Location = New System.Drawing.Point(264, 688)
        Me.ButtonTeachCalibrate.Name = "ButtonTeachCalibrate"
        Me.ButtonTeachCalibrate.Size = New System.Drawing.Size(128, 40)
        Me.ButtonTeachCalibrate.TabIndex = 6
        Me.ButtonTeachCalibrate.Text = "Dispense"
        '
        'DispenserType
        '
        Me.DispenserType.Location = New System.Drawing.Point(144, 32)
        Me.DispenserType.Name = "DispenserType"
        Me.DispenserType.Size = New System.Drawing.Size(240, 23)
        Me.DispenserType.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 760)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Result"
        '
        'rtbResult
        '
        Me.rtbResult.Location = New System.Drawing.Point(96, 768)
        Me.rtbResult.Name = "rtbResult"
        Me.rtbResult.ReadOnly = True
        Me.rtbResult.Size = New System.Drawing.Size(360, 112)
        Me.rtbResult.TabIndex = 9
        Me.rtbResult.Text = ""
        '
        'Status
        '
        Me.Status.Location = New System.Drawing.Point(96, 736)
        Me.Status.Name = "Status"
        Me.Status.ReadOnly = True
        Me.Status.Size = New System.Drawing.Size(360, 27)
        Me.Status.TabIndex = 7
        Me.Status.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 736)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Status"
        '
        'gpbDualHead
        '
        Me.gpbDualHead.Controls.Add(Me.chkDualHead)
        Me.gpbDualHead.Location = New System.Drawing.Point(104, 936)
        Me.gpbDualHead.Name = "gpbDualHead"
        Me.gpbDualHead.Size = New System.Drawing.Size(176, 42)
        Me.gpbDualHead.TabIndex = 35
        Me.gpbDualHead.TabStop = False
        Me.gpbDualHead.Text = " System with Single/Dual Head"
        '
        'chkDualHead
        '
        Me.chkDualHead.Location = New System.Drawing.Point(24, 16)
        Me.chkDualHead.Name = "chkDualHead"
        Me.chkDualHead.Size = New System.Drawing.Size(120, 16)
        Me.chkDualHead.TabIndex = 0
        Me.chkDualHead.Text = "Dual Head"
        '
        'VolumeCalibrationSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1152, 888)
        Me.Controls.Add(Me.gpbDualHead)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "VolumeCalibrationSettings"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.BoxSettings.ResumeLayout(False)
        CType(Me.AugerValveRetractDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AugerRetractTimeValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.valveTempValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.jettingpulseOnValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.jettingNumOfDispense, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.jettingPulseOffValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SetupDispenseDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SetupMaterialAirPressure, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tolerance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DesiredWeight, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumberOfAttempts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SetupRPM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AdjustedMaterialAirPressure, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AdjustedDispenseDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractDelayValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractSpeedValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractHeightValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaterialAirPressureStepValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RPMStepValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DurationStepValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AdjustedRPM, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpbDualHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'this is exposed for outer modules to call volume calibration
    Public DispenseDuration, PulseOffDuration, NoOfDispense, ValveTemperature, RPM, AugerRetractTime, AugerRetractDelay, MaterialAirPressure, RetractHeight, RetractSpeed, RetractDelay As Double

    'this is exposed for other modules to read the success/fail status of volume calibration
    Public VolumeCalibrationResult As String
    'this is exposed for other modules to read the current status of volume calibration
    'stopped, running or paused
    Public VolumeCalibrationState As String = "Stopped"
    Public VCRunning As Boolean = False

    'options are "init", "decrease", and "increase"
    Private NextAction As String = "Init"
    Private DispenseDurationStep As Double = 200 '200ms
    Private RPMStep As Double = 5
    Private MaterialAirPressureStep As Double = 0.1
    Private TimeOutDuration As Integer = 15 'in seconds

    Public Delegate Sub VCStatusUpdateDelegate(ByVal status As String)
    Public VCStatusUpdate As VCStatusUpdateDelegate = Nothing
    Public Function ClearCalFile()
        Dim t = New VolumeCalibration.VolumeCalPostHandler
        t.DeleteCalFile()
    End Function

    Friend Sub Weighting_T1_Tick()
        'If Form_Service.NoActionToExecute Then
        '    Form_Service.SetEventCode("1004200")
        '    IDS.Devices.Weighting.GetWeightingError(IDS.__ID)
        'End If
        'If IDS.__ID = "1014201" Or IDS.__ID = "1014202" Then
        '    Form_Service.ResetEventCode()
        'End If
    End Sub

    Public Sub DownloadAdjustedParameters()
        Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
        With IDS.Data.Hardware.Volume.Left
            If CalibrateDuration() Then
                MyDispenserSettings.DownloadJettingParameters(.AdjustedDispenseDuration)
            ElseIf CalibratePressure() Then
                MyDispenserSettings.DownloadMaterialAirPressure(.AdjustedMaterialAirPressure, SuckbackPressure)
            ElseIf CalibrateRPM() Then
                MyDispenserSettings.DownloadAugerRPM(.AdjustedRPM, IDS.Data.Hardware.Dispenser.Left.RetractTime, IDS.Data.Hardware.Dispenser.Left.RetractDelay)
            End If
        End With
    End Sub

    Public Function CalibrateRPM()
        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If HeadType = "Auger Valve" Then Return True
        Return False
    End Function

    Public Function CalibrateDuration()
        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If HeadType = "Jetting Valve" Or HeadType = "Slider Valve" Then Return True
        Return False
    End Function

    Public Function CalibratePressure()
        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If HeadType = "Time Pressure Valve" Or HeadType = "Time Pressure Syringe" Then Return True
        Return False
    End Function

    Public Function DoRetract()
        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If HeadType = "Jetting Valve" Then Return False
        Return True
    End Function

    Private Function CheckState() As Boolean
        If VolumeCalibrationState = "Stopped" Then
            Return False
        ElseIf VolumeCalibrationState = "Paused" Then
            Do
                Sleep(1)
                TraceDoEvents()
                If VolumeCalibrationState = "Stopped" Then
                    Return False
                ElseIf VolumeCalibrationState = "Running" Then
                    Return True
                ElseIf m_Tri.EStopActivated Then
                    Return False
                End If
            Loop Until VolumeCalibrationState <> "Paused" Or m_Tri.EStopActivated
        ElseIf VolumeCalibrationState = "Running" Then
            Return True
        End If
    End Function

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        WeightingScaleForm.Hide()
        RemovePanel(CurrentControl)
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveData()
    End Sub
    Private Sub ShowJettingParam(ByVal toShow As Boolean)
        Me.jettingNumOfDispense.Visible = toShow
        Me.jettingPulseOffValue.Visible = toShow
        Me.jettingpulseOnValue.Visible = toShow
        Me.lbJettingNoOfDispense.Visible = toShow
        Me.lbJettingPulseOff.Visible = toShow
        Me.lbJettingPulseOn.Visible = toShow
        'Me.lbJettingPulseOffTime.Visible = toShow
        'Me.lbJettingPulseOnTime.Visible = toShow
        Me.valveTempValue.Visible = toShow
        Me.lbValveTemp.Visible = toShow
    End Sub
    'all this text can't be helped to toggle on and off the label/valuebox visibility. but if the programmer desires, he can put them into groupboxes 
    Public Sub RevertData(Optional ByVal hideexit As Boolean = False)
        ButtonExit.Visible = Not hideexit
        With IDS.Data.Hardware.Dispenser

            DispenserType.Text = IDS.Data.Hardware.Dispenser.Left.HeadType

            If CalibratePressure() Then
                SetupMaterialAirPressure.Visible = True
                LabelSetupMaterialAirPressure1.Visible = True
                LabelSetupMaterialAirPressure2.Visible = True
                AdjustedMaterialAirPressure.Visible = True
                LabelAdjustedMaterialAirPressure1.Visible = True
                LabelAdjustedMaterialAirPressure2.Visible = True
                LabelPressureStep1.Visible = True
                LabelPressureStep2.Visible = True
                MaterialAirPressureStepValue.Visible = True

                AdjustedDispenseDuration.Visible = False
                LabelAdjustedDispenseDuration1.Visible = False
                LabelAdjustedDispenseDuration2.Visible = False
                LabelDurationStep1.Visible = False
                LabelDurationStep2.Visible = False
                DurationStepValue.Visible = False
                AdjustedRPM.Visible = False
                LabelAdjustedRPM1.Visible = False
                RPMStepValue.Visible = False
                LabelRPMStep1.Visible = False
                SetupRPM.Visible = False
                LabelSetupRPM.Visible = False
                ShowJettingParam(False)
                SetupDispenseDuration.Visible = True
                LabelSetupDispenseDuration2.Visible = True
                LabelSetupDispenseDuration1.Visible = True
                Me.lbAugerRetractTime.Visible = False
                Me.AugerRetractTimeValue.Visible = False
                Me.lbAugerRetractDelay.Visible = False
                Me.AugerValveRetractDelay.Visible = False

            ElseIf CalibrateDuration() Then
                AdjustedDispenseDuration.Visible = True
                LabelAdjustedDispenseDuration1.Visible = True
                LabelAdjustedDispenseDuration2.Visible = True
                LabelDurationStep1.Visible = True
                LabelDurationStep2.Visible = True
                DurationStepValue.Visible = True

                AdjustedMaterialAirPressure.Visible = False
                LabelAdjustedMaterialAirPressure1.Visible = False
                LabelAdjustedMaterialAirPressure2.Visible = False
                LabelPressureStep1.Visible = False
                LabelPressureStep2.Visible = False
                MaterialAirPressureStepValue.Visible = False
                AdjustedRPM.Visible = False
                LabelAdjustedRPM1.Visible = False
                RPMStepValue.Visible = False
                LabelRPMStep1.Visible = False
                SetupRPM.Visible = False
                LabelSetupRPM.Visible = False
                ShowJettingParam(True)
                SetupDispenseDuration.Visible = False
                LabelSetupDispenseDuration2.Visible = False
                LabelSetupDispenseDuration1.Visible = False
                Me.lbAugerRetractTime.Visible = False
                Me.AugerRetractTimeValue.Visible = False
                Me.lbAugerRetractDelay.Visible = False
                Me.AugerValveRetractDelay.Visible = False

            ElseIf CalibrateRPM() Then
                ShowJettingParam(False)
                LabelAdjustedRPM1.Visible = True
                AdjustedRPM.Visible = True
                RPMStepValue.Visible = True
                LabelRPMStep1.Visible = True
                SetupRPM.Visible = True
                LabelSetupRPM.Visible = True

                AdjustedDispenseDuration.Visible = False
                LabelAdjustedDispenseDuration1.Visible = False
                LabelAdjustedDispenseDuration2.Visible = False
                AdjustedMaterialAirPressure.Visible = False
                LabelAdjustedMaterialAirPressure1.Visible = False
                LabelAdjustedMaterialAirPressure2.Visible = False
                LabelDurationStep1.Visible = False
                LabelDurationStep2.Visible = False
                DurationStepValue.Visible = False
                LabelPressureStep1.Visible = False
                LabelPressureStep2.Visible = False
                MaterialAirPressureStepValue.Visible = False
                SetupDispenseDuration.Visible = True
                LabelSetupDispenseDuration2.Visible = True
                LabelSetupDispenseDuration1.Visible = True
                Me.lbAugerRetractTime.Visible = True
                Me.AugerRetractTimeValue.Visible = True
                Me.lbAugerRetractDelay.Visible = True
                Me.AugerValveRetractDelay.Visible = True
            End If

        End With

        With IDS.Data.Hardware.Volume
            If .Left.CalibrationType = "By Desired Weight" Then
                CalibrationType.SelectedIndex = 0
            ElseIf .Left.CalibrationType = "By Current Parameters" Then
                CalibrationType.SelectedIndex = 1
            Else
                CalibrationType.SelectedIndex = 1
            End If
            'SetupMaterialAirPressure.Text = .Left.SetupMaterialAirPressure
            'SetupRPM.Text = .Left.SetupRPM
            'SetupDispenseDuration.Text = .Left.SetupDispenseDuration
            'DesiredWeight.Text = .Left.DesiredWeight
            'Tolerance.Text = .Left.Tolerance
            'NumberOfAttempts.Text = .Left.NumberOfAttempts
            'RetractSpeedValue.Text = .Left.RetractSpeed
            'RetractHeightValue.Text = .Left.RetractHeight
            'RetractDelayValue.Text = .Left.RetractDelay
            'MaterialAirPressureStepValue.Text = .Left.PressureStepValue
            'RPMStepValue.Text = .Left.RPMStepValue
            'DurationStepValue.Text = .Left.DurationStepValue
            'AdjustedDispenseDuration.Text = .Left.AdjustedDispenseDuration
            'AdjustedMaterialAirPressure.Text = .Left.AdjustedMaterialAirPressure
            'AdjustedRPM.Text = .Left.AdjustedRPM
            'Me.AugerRetractTimeValue.Text = IDS.Data.Hardware.Dispenser.Left.RetractTime

            SetupDispenseDuration.Text = .Left.SetupDispenseDuration
            DesiredWeight.Text = .Left.DesiredWeight
            Tolerance.Text = .Left.Tolerance
            NumberOfAttempts.Text = .Left.NumberOfAttempts
            RetractSpeedValue.Text = .Left.RetractSpeed
            RetractHeightValue.Text = .Left.RetractHeight
            RetractDelayValue.Text = .Left.RetractDelay
            MaterialAirPressureStepValue.Text = .Left.PressureStepValue
            RPMStepValue.Text = .Left.RPMStepValue
            DurationStepValue.Text = .Left.DurationStepValue
            AdjustedDispenseDuration.Text = .Left.AdjustedDispenseDuration
            AdjustedMaterialAirPressure.Text = .Left.AdjustedMaterialAirPressure
            AdjustedRPM.Text = .Left.AdjustedRPM
            With IDS.Data.Hardware.Dispenser.Left
                SetupMaterialAirPressure.Text = .MaterialAirPressure
                SetupRPM.Text = .RPM
                Me.AugerRetractTimeValue.Text = .RetractTime
                If CalibrateRPM() Then
                    'Me.RetractDelayValue.Text = .RetractDelay
                    Me.AugerValveRetractDelay.Text = .RetractDelay
                End If
                Me.AugerRetractTimeValue.Text = .RetractTime
                Me.jettingPulseOffValue.Text = .Pause
                Me.jettingpulseOnValue.Text = .Pulse
                Me.valveTempValue.Text = .ValveTemperature
            End With
        End With

    End Sub

    Public Sub SaveData()
        If CalibrateDuration() Then
            With IDS.Data.Hardware.Dispenser.Left
                .Pulse = Me.jettingpulseOnValue.Value
                .Count = Me.jettingNumOfDispense.Value
                .ValveTemperature = Me.valveTempValue.Value
                .Pause = Me.jettingPulseOffValue.Value
                .MaterialAirPressure = Me.SetupMaterialAirPressure.Value
            End With
        ElseIf CalibratePressure() Then
            IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure = Me.SetupMaterialAirPressure.Value
        ElseIf CalibrateRPM() Then
            With IDS.Data.Hardware.Dispenser.Left
                .RPM = Me.SetupRPM.Value
                .MaterialAirPressure = Me.SetupMaterialAirPressure.Value
                .RetractTime = Me.AugerRetractTimeValue.Value
                .RetractDelay = Me.AugerValveRetractDelay.Value
            End With
        End If

        With IDS.Data.Hardware.Volume.Left
            .SetupMaterialAirPressure = SetupMaterialAirPressure.Value
            .SetupRPM = SetupRPM.Value
            .SetupDispenseDuration = SetupDispenseDuration.Value
            .DesiredWeight = DesiredWeight.Value
            .Tolerance = Tolerance.Value
            .CalibrationType = CalibrationType.Text
            .NumberOfAttempts = NumberOfAttempts.Value
            .RetractSpeed = RetractSpeedValue.Value
            .RetractHeight = RetractHeightValue.Value
            .RetractDelay = RetractDelayValue.Value
            .PressureStepValue = MaterialAirPressureStepValue.Value
            .DurationStepValue = Me.DurationStepValue.Value
            .RPMStepValue = Me.RPMStepValue.Value
            .PulseOnDuration = Me.jettingpulseOnValue.Value
            .PulseOffDuration = Me.jettingPulseOffValue.Value
            .JettingNoOfDispense = Me.jettingNumOfDispense.Value
        End With

        IDS.Data.SaveData()
    End Sub

    Public Sub ButtonCalibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click

        If ButtonCalibrate.Text = "Calibrate" Then
            Vision.FrmVision.SwitchCamera("View Camera")
            Vision.FrmVision.ClearDisplay()
            Cursor = Cursors.WaitCursor
            rtbResult.Text = ""
            Status.Text = ""
            'Weighting_Scale.OpenPort()
            Cursor = Cursors.Default
            attemptCount = 0
            VolumeCalibrationState = "Running"
            Me.VCLocalParamSetup()
            rtbResult.Clear()
            rtbResult.Refresh()
            Status.Text = ""
            Status.Refresh()
            'override the global IDS.Data values
            ' MaterialAirPressure = SetupMaterialAirPressure.Value
            'RPM = SetupRPM.Value
            'DispenseDuration = SetupDispenseDuration.Value
            RetractSpeed = RetractSpeedValue.Value
            RetractDelay = RetractDelayValue.Value
            RetractHeight = RetractHeightValue.Value
            m_Tri.SteppingButtons.Enabled = False
            m_Tri.SetMachineRun()
            BoxSettings.Enabled = False
            ButtonCalibrate.Text = "Stop"
            ButtonCalibrate.Refresh()
            With IDS.Data.Hardware.Volume.Left
                If CalibrateDuration() Then
                    .AdjustedDispenseDuration = DispenseDuration
                    Try
                        MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
                    Catch ex As Exception
                        MessageBox.Show("Error occur when download jetting parameters")
                        Return
                    End Try
                ElseIf CalibrateRPM() Then
                    .AdjustedRPM = RPM
                    Try
                        MyDispenserSettings.DownloadAugerRPM(RPM, AugerRetractTime, AugerRetractDelay)
                    Catch ex As Exception
                        MessageBox.Show("Error occur when download auger rpm")
                        Return
                    End Try
                End If
            End With
            Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
            Try
                MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
            Catch ex As Exception
                MessageBox.Show("Error occur when download material air pressure parameters")
                Return
            End Try
            VolumeCalibration(NumberOfAttempts.Value)
            'If DispensingResult = "Stopped" Then
            '    m_Tri.Move_Z(0)
            'End If
            'ButtonCalibrate.Text = "Calibrate"
            'ButtonCalibrate.Enabled = True
            'BoxSettings.Enabled = True
            'm_Tri.ResetCalibrationFlag()
            'm_Tri.SetMachineStop()
            'm_Tri.SteppingButtons.Enabled = True
            'disable temporarily for testing
        ElseIf ButtonCalibrate.Text = "Stop" Then
            ButtonCalibrate.Text = "Stopping"
            ButtonCalibrate.Enabled = False
            Dim getWeight As Boolean = Weighting_Scale.RequestWeightUpdate
            Weighting_Scale.RequestWeightUpdate = False
            m_Tri.TrioStop()
            m_Tri.ResetCalibrationFlag()
            SetServiceSpeed()
            m_Tri.Move_Z(0)
            VolumeCalibrationState = "Stopped"
            NextAction = "Init"
            attemptCount = 0
            ButtonCalibrate.Text = "Calibrate"
            ButtonCalibrate.Enabled = True
            BoxSettings.Enabled = True
            m_Tri.SetMachineStop()
            m_Tri.SteppingButtons.Enabled = True
            If Not getWeight Then
                ButtonCalibrate.Enabled = True
            End If
            'ButtonCalibrate.Text = "Calibrate"
        End If

    End Sub
    Private Function VContainerFullResponse()
        Dim fm As InfoForm = New InfoForm(True)
        fm.SetTitle("Volume Calibration Warning")
        fm.SetMessage("")
        fm.AddNewLine("Volume Calibration container is full. Please clean the container now")
        fm.SetOKButtonText("Clean Now")
        fm.SetCancelButtonText("Ignore")
        If fm.ShowDialog() = DialogResult.OK Then
            fm.SetMessage("")
            fm.AddNewLine("Click on the Done button after change/clean the volume calibration container.")
            fm.SetOKButtonText("Done")
            fm.OkOnly()
            fm.ShowDialog()
            fm.SetOKButtonText("Continue")
            fm.SetCancelButtonText("Abort Process")
            fm.SetMessage("")
            fm.AddNewLine("Click continue button to continue the volume calibration now")
            fm.OKCancelOnly()
            If Not (fm.ShowDialog() = DialogResult.OK) Then
                Return False
            Else

                VolumeCalibrationSetup(True) 'Continue and reset all dispense dot
                Return True
            End If
        Else
            Return False
        End If
    End Function
    Public attemptCount As Integer = 0
    Dim DispensingResult As String
    Dim VCAttemp As Integer = 0
    Public Function VolumeCalibration(ByVal AttemptsLeft As Integer) As Boolean

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        Dim Attempts As Integer = IDS.Data.Hardware.Volume.Left.NumberOfAttempts

        VCAttemp = AttemptsLeft
        If AttemptsLeft > 0 Then
            Dim rtn As Integer = VolumeCalibrationSetup()
            If Not (rtn = 0) Then
                If rtn = VCalStatus.containerFull Then
                    If Not VContainerFullResponse() Then
                        DispensingResult = "Stopped"
                        Status.Text = "Volume calibration failed."
                        NextAction = "Init"
                        VCRunning = False
                        attemptCount = 0
                        ButtonCalibrate.Text = "Calibrate"
                        BoxSettings.Enabled = True
                        Return False
                    End If
                Else
                    MessageBox.Show("Error occur when doing volume calibration setup #" + rtn.ToString())
                    DispensingResult = "Stopped"
                    Return False
                End If

            End If
            attemptCount += 1
            rtbResult.Text = "#" & attemptCount
            If m_Tri.EStopActivated() Or m_Tri.MachineHoming Then Return False

            If CalibrateDuration() Then
                DispensingResult = DispenseAndWeightDuration()
            ElseIf CalibratePressure() Then
                DispensingResult = DispenseAndWeightPressure()

            ElseIf CalibrateRPM() Then
                DispensingResult = DispenseAndWeightRPM()
            End If
            volumeCalPostHandler.SetCurrentDotDispensed()
            'If DispensingResult = "Success" Then
            '    'adjusted pressure/rpm/dispenseduration is saved in WithinTolerance
            '    With IDS.Data.Hardware.Volume.Left
            '        If CalibrateDuration() Then
            '            AdjustedDispenseDuration.Text = DispenseDuration
            '            .AdjustedDispenseDuration = DispenseDuration
            '            Try
            '                MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
            '            Catch ex As Exception
            '                MessageBox.Show("Error occur when download jetting parameters")
            '                Return False
            '            End Try

            '        ElseIf CalibratePressure() Then
            '            AdjustedMaterialAirPressure.Text = MaterialAirPressure
            '            .AdjustedMaterialAirPressure = MaterialAirPressure
            '            Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
            '            Try
            '                MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
            '            Catch ex As Exception
            '                MessageBox.Show("Error occur when download material air pressure parameters")
            '                Return False
            '            End Try

            '        ElseIf CalibrateRPM() Then
            '            AdjustedRPM.Text = RPM
            '            .AdjustedRPM = RPM
            '            Try
            '                MyDispenserSettings.DownloadAugerRPM(RPM, AugerRetractTime, AugerRetractDelay)
            '            Catch ex As Exception
            '                MessageBox.Show("Error occur when download auger rpm")
            '                Return False
            '            End Try

            '        End If
            '        NextAction = "Init"
            '        IDS.Data.SaveData()
            '        attemptCount = 0
            '        Status.Text = "Volume calibration success."
            '        Return True
            '    End With
            'ElseIf DispensingResult = "Failed" Then
            '    Return VolumeCalibration(AttemptsLeft - 1)
            'ElseIf DispensingResult = "Stopped" Then
            '    Status.Text = "Volume calibration stopped."
            '    NextAction = "Init"
            '    VCRunning = False
            '    attemptCount = 0
            '    Return False
            'End If
        Else
            Status.Text = "Volume calibration failed."
            NextAction = "Init"
            VCRunning = False
            attemptCount = 0
            ButtonCalibrate.Text = "Calibrate"
            BoxSettings.Enabled = True
            m_Tri.SteppingButtons.Enabled = True
            MyVolumeCalibrationSettings.VolumeCalibrationState = "FAILED"
            Return False
        End If

    End Function
    Public Function AfterGetWeight(ByVal success As Boolean)
        If success Then
            'adjusted pressure/rpm/dispenseduration is saved in WithinTolerance
            With IDS.Data.Hardware.Volume.Left
                If CalibrateDuration() Then
                    AdjustedDispenseDuration.Text = DispenseDuration
                    .AdjustedDispenseDuration = DispenseDuration
                    Try
                        MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
                    Catch ex As Exception
                        MessageBox.Show("Error occur when download jetting parameters")
                        Return False
                    End Try

                ElseIf CalibratePressure() Then
                    AdjustedMaterialAirPressure.Text = MaterialAirPressure
                    .AdjustedMaterialAirPressure = MaterialAirPressure
                    Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
                    Try
                        MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
                    Catch ex As Exception
                        MessageBox.Show("Error occur when download material air pressure parameters")
                        Return False
                    End Try

                ElseIf CalibrateRPM() Then
                    AdjustedRPM.Text = RPM
                    .AdjustedRPM = RPM
                    Try
                        MyDispenserSettings.DownloadAugerRPM(RPM, AugerRetractTime, AugerRetractDelay)
                    Catch ex As Exception
                        MessageBox.Show("Error occur when download auger rpm")
                        Return False
                    End Try

                End If
                NextAction = "Init"
                IDS.Data.SaveData()
                attemptCount = 0
                Status.Text = "Volume calibration success."
                If DispensingResult = "Stopped" Then
                    m_Tri.Move_Z(0)
                End If
                ButtonCalibrate.Text = "Calibrate"
                ButtonCalibrate.Enabled = True
                BoxSettings.Enabled = True
                m_Tri.ResetCalibrationFlag()
                m_Tri.SetMachineStop()
                m_Tri.SteppingButtons.Enabled = True
                MyVolumeCalibrationSettings.VolumeCalibrationState = "SUCCESS"
                Return True
            End With
        Else
            If VolumeCalibrationState = "Stopped" Then
                NextAction = "Init"
                attemptCount = 0
                ButtonCalibrate.Text = "Calibrate"
                ButtonCalibrate.Enabled = True
                BoxSettings.Enabled = True
                m_Tri.ResetCalibrationFlag()
                m_Tri.SetMachineStop()
                m_Tri.SteppingButtons.Enabled = True
                Exit Function
            End If
            Return VolumeCalibration(VCAttemp - 1)
        End If
        'ElseIf DispensingResult = "Stopped" Then
        '    Status.Text = "Volume calibration stopped."
        '    NextAction = "Init"
        '    VCRunning = False
        '    attemptCount = 0
        '    Return False
        'End If
    End Function
    Public volumeCalPostHandler As VolumeCalibration.VolumeCalPostHandler
    Public vcPostParam As VolumeCalibration.VolumeCalPostHandler.VolumeCalPostHandlerParam
    Enum VCalStatus
        success = 0
        postHandlerFileError = -1
        containerFull = -2
        interrupt = -3
    End Enum
    'Setup param based on what value in the dialog form only but not the file
    Public Function VCLocalParamSetup()
        If CalibrateDuration() Then 'Jetting/Slider valve
            DispenseDuration = Me.jettingpulseOnValue.Value
            PulseOffDuration = Me.jettingPulseOffValue.Value
            NoOfDispense = Me.jettingNumOfDispense.Value
            ValveTemperature = Me.valveTempValue.Value
            MaterialAirPressure = Me.SetupMaterialAirPressure.Value
        ElseIf CalibratePressure() Then
            DispenseDuration = Me.SetupDispenseDuration.Value
            MaterialAirPressure = Me.SetupMaterialAirPressure.Value
        ElseIf CalibrateRPM() Then
            DispenseDuration = Me.SetupDispenseDuration.Value
            MaterialAirPressure = Me.SetupMaterialAirPressure.Value
            RPM = Me.SetupRPM.Value
            AugerRetractTime = Me.AugerRetractTimeValue.Value
            AugerRetractDelay = Me.AugerValveRetractDelay.Value
        End If
    End Function
    Public Function VolumeCalibrationParamSetup()
        'Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            'If ajusted before and not reset, the adjusted value should be used for VC
            If IDS.Data.Hardware.Volume.Left.AdjustedDispenseDuration > 0 Then
                DispenseDuration = IDS.Data.Hardware.Volume.Left.AdjustedDispenseDuration
            Else
                DispenseDuration = IDS.Data.Hardware.Volume.Left.PulseOnDuration
            End If
            PulseOffDuration = IDS.Data.Hardware.Dispenser.Left.Pause
            NoOfDispense = IDS.Data.Hardware.Dispenser.Left.Count
            ValveTemperature = IDS.Data.Hardware.Dispenser.Left.ValveTemperature
            MaterialAirPressure = IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure 'IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
        ElseIf CalibratePressure() Then
            DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
            'If ajusted before and not reset, the adjusted value should be used for VC
            If IDS.Data.Hardware.Volume.Left.AdjustedMaterialAirPressure > 0 Then
                MaterialAirPressure = IDS.Data.Hardware.Volume.Left.AdjustedMaterialAirPressure
            Else
                MaterialAirPressure = IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure 'IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
            End If
        ElseIf CalibrateRPM() Then
            MaterialAirPressure = IDS.Data.Hardware.Dispenser.Left.MaterialAirPressure 'IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
            DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
            AugerRetractTime = IDS.Data.Hardware.Dispenser.Left.RetractTime
            Me.AugerRetractDelay = IDS.Data.Hardware.Dispenser.Left.RetractDelay
            'If ajusted before and not reset, the adjusted value should be used for VC
            If IDS.Data.Hardware.Volume.Left.AdjustedRPM > 0 Then
                RPM = IDS.Data.Hardware.Volume.Left.AdjustedRPM
            Else
                RPM = IDS.Data.Hardware.Volume.Left.SetupRPM
            End If
        End If
    End Function
    Public Function VolumeCalibrationSetup(Optional ByVal resetDispenseDot As Boolean = False) As Integer

        VolumeCalibrationState = "Running"

        SetServiceSpeed()
        With IDS.Data.Hardware.Gantry.WeighingScalePosition
            x = .X
            y = .Y
            z = .Z
        End With
        With IDS.Data.Hardware.Needle
            offset_x = .Left.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
            offset_y = .Left.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
            offset_z = .Left.NeedleCalibrationPosition.Z
        End With
        position(0) = x - offset_x
        position(1) = y - offset_y
        position(2) = z + offset_z
        If volumeCalPostHandler Is Nothing Then
            vcPostParam.topLeftX = position(0)
            vcPostParam.topLeftY = position(1)
            vcPostParam.pitch = 0.01
            vcPostParam.dotDiameter = 3 'mm  'IDS.Data.Hardware.Needle.Left.DotDiameter
            vcPostParam.bottomRightX = IDS.Data.Hardware.Gantry.WeighingScaleBottomRight.X - offset_x
            vcPostParam.bottomRightY = IDS.Data.Hardware.Gantry.WeighingScaleBottomRight.Y - offset_y
            volumeCalPostHandler = New VolumeCalibration.VolumeCalPostHandler(vcPostParam)
        End If
        If Not (volumeCalPostHandler.OldPostFileExist()) Or resetDispenseDot Then
            volumeCalPostHandler.PopulateNewPost()
        End If
        If Not (volumeCalPostHandler.GetAllPopulatedPost()) Then
            Return VCalStatus.postHandlerFileError
        End If
        Dim xTemp, yTemp As Double
        If Not volumeCalPostHandler.GetNextPost(xTemp, yTemp) Then Return VCalStatus.containerFull
        position(0) = xTemp
        position(1) = yTemp
        '1)move to the calibration position
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration

        With IDS.Data.Hardware.Volume.Left
            RetractHeight = .RetractHeight
            RetractSpeed = .RetractSpeed
            RetractDelay = .RetractDelay
            MaterialAirPressureStep = .PressureStepValue
        End With

        'Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        'If CalibrateDuration() Then
        '    DispenseDuration = IDS.Data.Hardware.Volume.Left.AdjustedDispenseDuration
        '    MaterialAirPressure = IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
        'ElseIf CalibratePressure() Then
        '    DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
        '    MaterialAirPressure = IDS.Data.Hardware.Volume.Left.AdjustedMaterialAirPressure
        'ElseIf CalibrateRPM() Then
        '    MaterialAirPressure = IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
        '    DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
        '    RPM = IDS.Data.Hardware.Volume.Left.AdjustedRPM
        'End If

        'disable temporarily for testing
        Return VCalStatus.success

StopCalibration:
        VolumeCalibrationResult = "Volume calibration setup interrupted."
        Return VCalStatus.interrupt

    End Function

    Public Function WithinTolerance(ByVal tolerance As Double, ByVal desired As Double, ByVal measured As Double) As String
        Dim lower_bound, higher_bound As Double
        lower_bound = desired - tolerance
        higher_bound = desired + tolerance
        If higher_bound >= measured And measured >= lower_bound And measured > 0 Then
            Return "Success"
        Else
            'compare how the measured weight relates to the desired weight and tolerance levels if it fails
            'then change the measured 
            If lower_bound > measured Then
                NextAction = "Increase"
            ElseIf higher_bound < measured Then
                NextAction = "Decrease"
            End If
            Return "Failed"
        End If
    End Function

    Public Function DispenseAndWeightPressure() As String

        Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
        MaterialAirPressureStep = IDS.Data.Hardware.Volume.Left.PressureStepValue
        If NextAction = "Decrease" Then
            MaterialAirPressure = MaterialAirPressure - MaterialAirPressureStep
        ElseIf NextAction = "Increase" Then
            MaterialAirPressure = MaterialAirPressure + MaterialAirPressureStep
        End If

        MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
        string_result = DispenseAndWeight1()
        Return string_result

    End Function

    Public Function DispenseAndWeightRPM() As String
        RPMStep = IDS.Data.Hardware.Volume.Left.RPMStepValue
        If NextAction = "Decrease" Then
            RPM = RPM - RPMStep
        ElseIf NextAction = "Increase" Then
            RPM = RPM + RPMStep
        End If

        MyDispenserSettings.DownloadAugerRPM(RPM, AugerRetractTime, AugerRetractDelay)
        string_result = DispenseAndWeight1()
        Return string_result

    End Function

    Public Function DispenseAndWeightDuration() As String

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        DispenseDurationStep = IDS.Data.Hardware.Volume.Left.DurationStepValue
        If NextAction = "Decrease" Then
            DispenseDuration = DispenseDuration - DispenseDurationStep

        ElseIf NextAction = "Increase" Then
            DispenseDuration = DispenseDuration + DispenseDurationStep
        End If

        If HeadType = "Jetting Valve" Then
            'MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
            MyDispenserSettings.DownloadJettingParameters(DispenseDuration, PulseOffDuration, NoOfDispense, ValveTemperature)
        ElseIf HeadType = "Slider Valve" Then
            MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
        End If

        string_result = DispenseAndWeight1()
        Return string_result

    End Function

    Public Function DispenseAndWeight() As String
        Dim errorcode As Integer = 0
        start_time = Now
        Status.Text = "Taring.."
        Weighting_Scale.DoTare()
        If Not CheckState() Then GoTo StopCalibration
        Status.Text = "Taring Done"

        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
        If CheckState() Then
            m_Tri.Move_Z(position(2))
        Else
            GoTo StopCalibration
        End If
        Status.Text = "Dispensing"
        Me.ExportStatus("Volume Calibration Attempt #" + attemptCount.ToString())
        m_Tri.m_TriCtrl.Execute("OP(25,1)")
        m_Tri.m_TriCtrl.Execute("WA(" & DispenseDuration.ToString & ")")
        m_Tri.m_TriCtrl.Execute("OP(25,0)")
        If Me.CalibrateDuration() Then
            Dim tempD As Double = PulseOffDuration * NoOfDispense + DispenseDuration * (NoOfDispense - 1)
            m_Tri.m_TriCtrl.Execute("WA(" & tempD.ToString & ")")
        End If
        m_Tri.m_TriCtrl.Execute("WA(" & RetractDelay.ToString & ")")
        If CalibrateRPM() Then
            m_Tri.m_TriCtrl.Execute("WA(" & AugerRetractTime.ToString & ")")
        End If
        If Not (IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve") Then
            If DoRetract() Then
                m_Tri.Set_Z_Speed(RetractSpeed)
                If CheckState() Then
                    m_Tri.Move_Z(position(2) + RetractHeight)
                Else
                    GoTo StopCalibration
                End If
            End If
        End If
        Status.Text = "Dispensing Done"
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
        If CheckState() Then
            m_Tri.Move_Z(0)
        Else
            GoTo StopCalibration
        End If

        Status.Text = "Waiting weighting scale.."
        '3)read the weighted value
        'need to add some kind of loop here to read until we obtain a valid value
        'wait for taring to be finished before we get weight
        'For i As Integer = 1 To 600
        '    Sleep(5)
        '    TraceDoEvents()
        'Next
        start_time = Now
        Weighting_Scale.GetWeight()
        Status.Text = "Reading weight.."
        Me.ExportStatus("Reading weight..")
        'Console.WriteLine("Read Weight")
        Do
            'Sleep(5)
            TraceDoEvents()
            'DisplayProgressText(Status)
            stop_time = Now
            elapsed_time = stop_time.Subtract(start_time)
        Loop Until Weighting_Scale.ValueUpdated Or elapsed_time.TotalSeconds > TimeOutDuration Or Not CheckState() Or m_Tri.EStopActivated
        Dim reading As String = CStr(Weighting_Scale.WeightReading)
        Weighting_Scale.ResetValues()
        If elapsed_time.TotalSeconds > TimeOutDuration Then
            errorcode = 2
            GoTo StopCalibration
        End If
        'Console.WriteLine("Weight Read")
        If Not CheckState() Then GoTo StopCalibration

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            VolumeCalibrationResult = " Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
        ElseIf CalibratePressure() Then
            VolumeCalibrationResult = " Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
        ElseIf CalibrateRPM() Then
            VolumeCalibrationResult = " RPM of " + RPM.ToString + " gives a result of " + reading + " mg."

        End If
        rtbResult.Text = VolumeCalibrationResult
        With IDS.Data.Hardware.Volume.Left
            Dim str As String = WithinTolerance(.Tolerance, .DesiredWeight, CDbl(reading))
            If str.ToUpper = "SUCCESS" Then
                VolumeCalibrationResult = "Vol. Cal. Success!" + VolumeCalibrationResult
                VolumeCalibrationResult += " Info=" + "Desired Weight:" + .DesiredWeight.ToString("0.0") + "mg" + " Tolerance: " + .Tolerance.ToString("0.0") + "mg "
                Me.ExportStatus(VolumeCalibrationResult)
            Else
                VolumeCalibrationResult = "Vol. Cal. Attemp #" + attemptCount.ToString() + " not success!" + VolumeCalibrationResult
                Me.ExportStatus(VolumeCalibrationResult)
            End If
            Return str
        End With

StopCalibration:
        If errorcode = 1 Then
            VolumeCalibrationResult = "Weighting Scale Taring timeout."
            Me.ExportStatus(VolumeCalibrationResult)
        ElseIf errorcode = 2 Then
            VolumeCalibrationResult = "Weighting Scale reading timeout."
            Me.ExportStatus(VolumeCalibrationResult)
        Else
            VolumeCalibrationResult = "Volume calibration stopped."
        End If
        rtbResult.Text = VolumeCalibrationResult

        Return "Stopped"

WeightTooLow:
        VolumeCalibrationResult = "Volume calibration failed because weight could not be detected."
        rtbResult.Text = VolumeCalibrationResult
        Me.ExportStatus(VolumeCalibrationResult)
        Return "Failed"

    End Function
    Dim VCValue(17) As Double
    Private Sub SetupVCParam()
        VCValue(0) = 3
        VCValue(1) = DispenseDuration
        VCValue(2) = position(2)
        VCValue(3) = position(2) + RetractHeight
        VCValue(4) = RetractDelay
        VCValue(5) = RetractSpeed
        VCValue(6) = IDS.Data.Hardware.Gantry.ServiceZSpeed 'forward speed
        If CalibrateDuration() Then 'Jetting/Slider valve
            VCValue(7) = 1
        ElseIf CalibratePressure() Then
            VCValue(7) = 0
        ElseIf CalibrateRPM() Then
            VCValue(7) = 2
        Else
            VCValue(7) = 0
        End If
        VCValue(8) = PulseOffDuration
        VCValue(9) = NoOfDispense
        VCValue(10) = AugerRetractTime
        VCValue(11) = AugerRetractDelay
        VCValue(12) = 0
        VCValue(13) = 0
        VCValue(14) = 0
        VCValue(15) = 0
        VCValue(16) = 0
        VCValue(17) = 0
    End Sub
    Public Function DispenseAndWeight1() As String
        Dim errorcode As Integer = 0
        start_time = Now
        Status.Text = "Taring.."
        Weighting_Scale.DoTare()
        If Not CheckState() Then GoTo StopCalibration
        Status.Text = "Taring Done"
        SetupVCParam()
        Dim stime As Long
        If (VolumeCalibrationState = "Stopped") Then
            GoTo StopCalibration
        End If
        If m_Tri.m_TriCtrl.SetTable(109, 18, VCValue) Then
            If Not m_Tri.SetCalibrationFlag() Then
                If Not m_Tri.SetCalibrationFlag() Then
                    If Not m_Tri.SetCalibrationFlag() Then
                        MessageBox.Show("Unable to set caibration flag")
                        GoTo StopCalibration
                    End If
                End If
            End If
            m_Tri.GetIDSState()
            If m_Tri.StateContainer(100) = 0 Then
                MessageBox.Show("State Container 100 = 0")
            End If
            If Not (m_Tri.RunTrioMotionProgram("CALIBRATIONS", 3)) Then
                MessageBox.Show("Run Cali program failed")
            End If
            If Not (m_Tri.RunTrioMotionProgram("CALIBRATIONS", 3)) Then
                MessageBox.Show("Run Cali program failed")
                GoTo StopCalibration
            End If
            If Not (m_Tri.RunTrioMotionProgram("CALIBRATIONS", 3)) Then
                MessageBox.Show("Run Cali program failed")
            End If
            'm_Tri.GetIDSState()
            stime = DateTime.Now.Ticks
            Dim table100(0) As Integer
            Dim tableValue As Integer = 1
            While (tableValue = 1 And Not m_Tri.EStopActivated) ' And Not m_Tri.EStopActivated)
                TraceDoEvents()
                If m_Tri.GetVCTable100(table100) Then
                    tableValue = table100(0)
                End If
                If (VolumeCalibrationState = "Stopped") Then
                    GoTo StopCalibration
                End If
            End While
        Else
            GoTo StopCalibration
        End If
        If m_Tri.EStopActivated Then
            GoTo StopCalibration
        End If
        If (DateTime.Now.Ticks - stime) / 10000 < 500 Then
            MessageBox.Show("Calibration got issues")
        End If
        m_Tri.StopTrioMotionProgram("CALIBRATIONS")
        Status.Text = "Waiting weighting scale.."
        start_time = Now
        Weighting_Scale.ReadWeightReturned = New WeightScale.IWeightingScale.ReadWeightDel(AddressOf Me.WeightReturn)
        Status.Text = "Reading weight.."
        Me.ExportStatus("Reading weight..")
        'ButtonCalibrate.Enabled = False
        Weighting_Scale.GetWeight()
        Exit Function
        'Do
        '    TraceDoEvents()
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until Weighting_Scale.ValueUpdated Or elapsed_time.TotalSeconds > TimeOutDuration Or Not CheckState() Or m_Tri.EStopActivated
        'Dim reading As String = CStr(Weighting_Scale.WeightReading)
        'Weighting_Scale.ResetValues()
        'If elapsed_time.TotalSeconds > TimeOutDuration Then
        '    errorcode = 2
        '    GoTo StopCalibration
        'End If
        ''Console.WriteLine("Weight Read")
        'If Not CheckState() Then GoTo StopCalibration

        'Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        'If CalibrateDuration() Then
        '    VolumeCalibrationResult = " Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
        'ElseIf CalibratePressure() Then
        '    VolumeCalibrationResult = " Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
        'ElseIf CalibrateRPM() Then
        '    VolumeCalibrationResult = " RPM of " + RPM.ToString + " gives a result of " + reading + " mg."

        'End If
        'rtbResult.Text = VolumeCalibrationResult
        'With IDS.Data.Hardware.Volume.Left
        '    Dim str As String = WithinTolerance(.Tolerance, .DesiredWeight, CDbl(reading))
        '    If str.ToUpper = "SUCCESS" Then
        '        VolumeCalibrationResult = "Vol. Cal. Success! " + VolumeCalibrationResult
        '        VolumeCalibrationResult += "Weight = " + reading + " mg."
        '        Me.ExportStatus(VolumeCalibrationResult)
        '    Else
        '        VolumeCalibrationResult = "Vol. Cal. Attemp #" + attemptCount.ToString() + " not success! " + reading + " mg."
        '        Me.ExportStatus(VolumeCalibrationResult)
        '    End If
        '    Return str
        'End With

StopCalibration:
        If errorcode = 1 Then
            VolumeCalibrationResult = "Weighting Scale Taring timeout."
            Me.ExportStatus(VolumeCalibrationResult)
        ElseIf errorcode = 2 Then
            VolumeCalibrationResult = "Weighting Scale reading timeout."
            Me.ExportStatus(VolumeCalibrationResult)
        Else
            VolumeCalibrationResult = "Volume calibration stopped."
        End If
        rtbResult.Text = VolumeCalibrationResult

        Return "Stopped"

WeightTooLow:
        VolumeCalibrationResult = "Volume calibration failed because weight could not be detected."
        rtbResult.Text = VolumeCalibrationResult
        Me.ExportStatus(VolumeCalibrationResult)
        Return "Failed"

    End Function
    Private Sub WeightReturn(ByVal weight As Double)
        Dim reading As String = CStr(weight)
        ButtonCalibrate.Enabled = True
        Weighting_Scale.ResetValues()
        Status.Text = "Weight returned:" & reading & " mg"
        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            VolumeCalibrationResult = " Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
        ElseIf CalibratePressure() Then
            VolumeCalibrationResult = " Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
        ElseIf CalibrateRPM() Then
            VolumeCalibrationResult = " RPM = " + RPM.ToString + " Weight = " + reading + " mg."

        End If
        rtbResult.Text = VolumeCalibrationResult
        With IDS.Data.Hardware.Volume.Left
            Dim str As String = WithinTolerance(.Tolerance, .DesiredWeight, CDbl(reading))
            If str.ToUpper = "SUCCESS" Then
                VolumeCalibrationResult = "Volume Calibration Success! " + VolumeCalibrationResult
                'VolumeCalibrationResult += "Weight = " + reading + " mg."
                Me.ExportStatus(VolumeCalibrationResult)
                AfterGetWeight(True)
            Else
                VolumeCalibrationResult = "Vol. Cal. Attemp #" + attemptCount.ToString() + " not success! " + reading + " mg."
                Me.ExportStatus(VolumeCalibrationResult)
                AfterGetWeight(False)
            End If
        End With
    End Sub
    Private Sub DispenseWeightReturn(ByVal weight As Double)

        Dim reading As String = CStr(weight)
        Weighting_Scale.ResetValues()

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            VolumeCalibrationResult = "Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
        ElseIf CalibratePressure() Then
            VolumeCalibrationResult = "Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
        ElseIf CalibrateRPM() Then
            VolumeCalibrationResult = "RPM of " + RPM.ToString + " gives a result of " + reading + " mg."
        Else
            MessageBox.Show("Unknown mode")
        End If

        Status.Text = "Dispensing done"
        ButtonTeachCalibrate.Text = "Dispense"
        BoxSettings.Enabled = True
        m_Tri.SetMachineStop()
        m_Tri.ResetCalibrationFlag()
        rtbResult.Text = VolumeCalibrationResult
        m_Tri.SteppingButtons.Enabled = True
    End Sub
    'Export the status display to any other form
    Public Function ExportStatus(ByVal status As String)
        If Not (Me.VCStatusUpdate Is Nothing) Then
            VCStatusUpdate(status)
        End If
    End Function


    Public Sub DisplayProgressText(ByVal status As Windows.Forms.TextBox)
        If status.Text = "Reading.." Then
            status.Text = "Reading..."
        ElseIf status.Text = "Reading..." Then
            status.Text = "Reading...."
        ElseIf status.Text = "Reading...." Then
            status.Text = "Reading.."
        ElseIf status.Text = "" Then
            status.Text = "Reading.."
        ElseIf status.Text = "Dispensing..." Then
            status.Text = "Reading.."
        ElseIf status.Text = "Taring.." Then
            status.Text = "Taring..."
        ElseIf status.Text = "Taring..." Then
            status.Text = "Taring...."
        ElseIf status.Text = "Taring...." Then
            status.Text = "Taring.."
        End If
    End Sub

    Private Sub CalibrationType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalibrationType.SelectedIndexChanged
        If CalibrationType.Text = "By Desired Weight" Then
            ButtonCalibrate.Enabled = True
            ButtonTeachCalibrate.Enabled = False
        ElseIf CalibrationType.Text = "By Current Parameters" Then
            ButtonCalibrate.Enabled = False
            ButtonTeachCalibrate.Enabled = True
        End If
    End Sub

    Private Sub ButtonTeachCalibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTeachCalibrate.Click

        If ButtonTeachCalibrate.Text = "Stop" Then
            ButtonTeachCalibrate.Enabled = False
            ButtonTeachCalibrate.Text = "Stopping Dispense"
            ButtonTeachCalibrate.Refresh()
            m_Tri.TrioStop()
            If Weighting_Scale.RequestWeightUpdate Then
                ButtonTeachCalibrate.Enabled = True
            End If
            Weighting_Scale.RequestWeightUpdate = False
            VolumeCalibrationState = "Stopped"

        ElseIf ButtonTeachCalibrate.Text = "Dispense" Then
            Vision.FrmVision.SwitchCamera("View Camera")
            Vision.FrmVision.ClearDisplay()
            Cursor = Cursors.WaitCursor
            rtbResult.Text = ""
            rtbResult.Refresh()
            Status.Text = ""
            Status.Refresh()
            Weighting_Scale.OpenPort()
            Cursor = Cursors.Default
            ButtonTeachCalibrate.Text = "Stop"
            ButtonTeachCalibrate.Refresh()
            Me.VCLocalParamSetup()
            Dim rtn As Integer = VolumeCalibrationSetup()
            If Not (rtn = 0) Then
                If rtn = Me.VCalStatus.containerFull Then
                    If Not VContainerFullResponse() Then
                        GoTo StopCalibration
                    End If
                Else
                    MessageBox.Show("Error occur when doing volume calibration setup #" + rtn.ToString())
                    GoTo StopCalibration
                End If

                'GoTo StopCalibration
            End If
            m_Tri.SetMachineRun()
            m_Tri.SteppingButtons.Enabled = False
            BoxSettings.Enabled = False

            Weighting_Scale.DoTare()
            start_time = Now
            Status.Text = "Taring.."

            If VolumeCalibrationState = "Stopped" Then GoTo StopCalibration

            Status.Text = "Tared"
            If Not CheckState() Then GoTo StopCalibration

            Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
            If CalibrateDuration() Then
                MyDispenserSettings.DownloadJettingParameters(DispenseDuration, PulseOffDuration, Me.NoOfDispense, ValveTemperature)
            ElseIf CalibratePressure() Then
                MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
            ElseIf CalibrateRPM() Then
                MyDispenserSettings.DownloadAugerRPM(RPM, AugerRetractTime, AugerRetractDelay)
            End If
            If VolumeCalibrationState = "Stopped" Then GoTo StopCalibration

            SetupVCParam()

            If m_Tri.m_TriCtrl.SetTable(109, 18, VCValue) Then
                m_Tri.SetCalibrationFlag()
                m_Tri.GetIDSState()
                m_Tri.RunTrioMotionProgram("CALIBRATIONS", 3)
                m_Tri.GetIDSState()
                Dim table100 As Integer = 1
                Dim tableValue(0) As Integer
                While (table100 = 1 And Not m_Tri.EStopActivated)
                    If m_Tri.GetVCTable100(tableValue) Then
                        table100 = tableValue(0)
                    End If
                    TraceDoEvents()
                    If (VolumeCalibrationState = "Stopped") Then
                        GoTo StopCalibration
                    End If
                End While
            Else
                GoTo StopCalibration
            End If
            If m_Tri.EStopActivated Then
                Exit Sub
            End If
            volumeCalPostHandler.SetCurrentDotDispensed()
            m_Tri.StopTrioMotionProgram("CALIBRATIONS")
            m_Tri.ResetCalibrationFlag()
            RetractSpeed = RetractSpeedValue.Value
            RetractDelay = RetractDelayValue.Value
            RetractHeight = RetractHeightValue.Value
            start_time = Now
            Weighting_Scale.ReadWeightReturned = New WeightScale.IWeightingScale.ReadWeightDel(AddressOf Me.DispenseWeightReturn)
            Weighting_Scale.GetWeight()
            Status.Text = "Reading.."
        End If
        Exit Sub

StopCalibration:
        BoxSettings.Enabled = True
        Weighting_Scale.RequestWeightUpdate = False
        m_Tri.SetMachineStop()
        m_Tri.ResetCalibrationFlag()
        SetServiceSpeed()
        m_Tri.Move_Z(0)
        rtbResult.Text = "Volume calibration stopped."
        ButtonTeachCalibrate.Text = "Dispense"
        ButtonTeachCalibrate.Enabled = True
        Status.Text = ""
        m_Tri.SteppingButtons.Enabled = True
        m_Tri.StopTrioMotionProgram("CALIBRATIONS")
        Exit Sub

Timeout:
        BoxSettings.Enabled = True
        m_Tri.SetMachineStop()
        rtbResult.Text = "Time out"
        ButtonTeachCalibrate.Text = "Dispense"
        Status.Text = ""
    End Sub


    Private Sub WeighingScaleButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeighingScaleButton.Click
        'Weighting_Scale.Location() = New System.Drawing.Point(0, 0)
        'If Weighting_Scale.Visible Then
        '    Weighting_Scale.Hide()
        'Else
        '    Weighting_Scale.Show()
        'End If
        Cursor = Cursors.WaitCursor
        If WeightingScaleForm.Visible Then
            WeightingScaleForm.Hide()
        Else
            WeightingScaleForm.Show()
        End If
        Cursor = Cursors.Default
    End Sub

End Class
