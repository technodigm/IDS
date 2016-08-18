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
    Friend WithEvents BoxStatus As System.Windows.Forms.GroupBox
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


    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VolumeCalibrationSettings))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.BoxStatus = New System.Windows.Forms.GroupBox
        Me.Status = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label40 = New System.Windows.Forms.Label
        Me.BoxSettings = New System.Windows.Forms.GroupBox
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
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.ButtonTeachCalibrate = New System.Windows.Forms.Button
        Me.WeighingScaleButton = New System.Windows.Forms.Button
        Me.DispenserType = New System.Windows.Forms.Label
        Me.gpbDualHead = New System.Windows.Forms.GroupBox
        Me.chkDualHead = New System.Windows.Forms.CheckBox
        Me.rtbResult = New System.Windows.Forms.RichTextBox
        Me.PanelToBeAdded.SuspendLayout()
        Me.BoxStatus.SuspendLayout()
        Me.BoxSettings.SuspendLayout()
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
        Me.PanelToBeAdded.Controls.Add(Me.BoxStatus)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.Label40)
        Me.PanelToBeAdded.Controls.Add(Me.BoxSettings)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonCalibrate)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonTeachCalibrate)
        Me.PanelToBeAdded.Controls.Add(Me.WeighingScaleButton)
        Me.PanelToBeAdded.Controls.Add(Me.DispenserType)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 0
        '
        'BoxStatus
        '
        Me.BoxStatus.Controls.Add(Me.rtbResult)
        Me.BoxStatus.Controls.Add(Me.Status)
        Me.BoxStatus.Controls.Add(Me.Label1)
        Me.BoxStatus.Controls.Add(Me.Label2)
        Me.BoxStatus.Location = New System.Drawing.Point(16, 664)
        Me.BoxStatus.Name = "BoxStatus"
        Me.BoxStatus.Size = New System.Drawing.Size(480, 208)
        Me.BoxStatus.TabIndex = 36
        Me.BoxStatus.TabStop = False
        Me.BoxStatus.Text = "Status"
        '
        'Status
        '
        Me.Status.Location = New System.Drawing.Point(96, 40)
        Me.Status.Name = "Status"
        Me.Status.ReadOnly = True
        Me.Status.Size = New System.Drawing.Size(360, 27)
        Me.Status.TabIndex = 7
        Me.Status.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Status"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Result"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
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
        Me.Label40.Location = New System.Drawing.Point(8, 80)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(136, 23)
        Me.Label40.TabIndex = 3
        Me.Label40.Text = "Dispenser Type:"
        '
        'BoxSettings
        '
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
        Me.BoxSettings.Location = New System.Drawing.Point(16, 112)
        Me.BoxSettings.Name = "BoxSettings"
        Me.BoxSettings.Size = New System.Drawing.Size(480, 496)
        Me.BoxSettings.TabIndex = 49
        Me.BoxSettings.TabStop = False
        Me.BoxSettings.Text = "Settings"
        '
        'SetupDispenseDuration
        '
        Me.SetupDispenseDuration.Location = New System.Drawing.Point(268, 112)
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
        Me.SetupMaterialAirPressure.Location = New System.Drawing.Point(268, 144)
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
        Me.Tolerance.Location = New System.Drawing.Point(268, 208)
        Me.Tolerance.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.Tolerance.Name = "Tolerance"
        Me.Tolerance.Size = New System.Drawing.Size(72, 27)
        Me.Tolerance.TabIndex = 17
        Me.Tolerance.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'DesiredWeight
        '
        Me.DesiredWeight.DecimalPlaces = 1
        Me.DesiredWeight.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.DesiredWeight.Location = New System.Drawing.Point(268, 80)
        Me.DesiredWeight.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.DesiredWeight.Name = "DesiredWeight"
        Me.DesiredWeight.Size = New System.Drawing.Size(72, 27)
        Me.DesiredWeight.TabIndex = 17
        Me.DesiredWeight.ThousandsSeparator = True
        Me.DesiredWeight.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'NumberOfAttempts
        '
        Me.NumberOfAttempts.Location = New System.Drawing.Point(268, 272)
        Me.NumberOfAttempts.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NumberOfAttempts.Name = "NumberOfAttempts"
        Me.NumberOfAttempts.Size = New System.Drawing.Size(72, 27)
        Me.NumberOfAttempts.TabIndex = 17
        Me.NumberOfAttempts.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(76, 272)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(168, 23)
        Me.Label49.TabIndex = 4
        Me.Label49.Text = "Number of Attempts"
        '
        'SetupRPM
        '
        Me.SetupRPM.DecimalPlaces = 1
        Me.SetupRPM.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.SetupRPM.Location = New System.Drawing.Point(268, 176)
        Me.SetupRPM.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.SetupRPM.Name = "SetupRPM"
        Me.SetupRPM.Size = New System.Drawing.Size(72, 27)
        Me.SetupRPM.TabIndex = 17
        Me.SetupRPM.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'LabelSetupMaterialAirPressure1
        '
        Me.LabelSetupMaterialAirPressure1.Location = New System.Drawing.Point(76, 144)
        Me.LabelSetupMaterialAirPressure1.Name = "LabelSetupMaterialAirPressure1"
        Me.LabelSetupMaterialAirPressure1.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupMaterialAirPressure1.TabIndex = 3
        Me.LabelSetupMaterialAirPressure1.Text = "Dispense Pressure"
        '
        'LabelDesiredWeight2
        '
        Me.LabelDesiredWeight2.Location = New System.Drawing.Point(340, 80)
        Me.LabelDesiredWeight2.Name = "LabelDesiredWeight2"
        Me.LabelDesiredWeight2.Size = New System.Drawing.Size(32, 24)
        Me.LabelDesiredWeight2.TabIndex = 11
        Me.LabelDesiredWeight2.Text = "mg"
        '
        'LabelSetTolerance
        '
        Me.LabelSetTolerance.Location = New System.Drawing.Point(76, 208)
        Me.LabelSetTolerance.Name = "LabelSetTolerance"
        Me.LabelSetTolerance.Size = New System.Drawing.Size(88, 23)
        Me.LabelSetTolerance.TabIndex = 8
        Me.LabelSetTolerance.Text = "Tolerance"
        '
        'LabelSetupDispenseDuration2
        '
        Me.LabelSetupDispenseDuration2.Location = New System.Drawing.Point(340, 112)
        Me.LabelSetupDispenseDuration2.Name = "LabelSetupDispenseDuration2"
        Me.LabelSetupDispenseDuration2.Size = New System.Drawing.Size(32, 24)
        Me.LabelSetupDispenseDuration2.TabIndex = 7
        Me.LabelSetupDispenseDuration2.Text = "ms"
        '
        'LabelSetupRPM
        '
        Me.LabelSetupRPM.Location = New System.Drawing.Point(76, 176)
        Me.LabelSetupRPM.Name = "LabelSetupRPM"
        Me.LabelSetupRPM.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupRPM.TabIndex = 3
        Me.LabelSetupRPM.Text = "Dispense RPM"
        '
        'LabelDesiredWeight1
        '
        Me.LabelDesiredWeight1.Location = New System.Drawing.Point(76, 80)
        Me.LabelDesiredWeight1.Name = "LabelDesiredWeight1"
        Me.LabelDesiredWeight1.Size = New System.Drawing.Size(128, 23)
        Me.LabelDesiredWeight1.TabIndex = 3
        Me.LabelDesiredWeight1.Text = "Desired Weight"
        '
        'LabelSetupDispenseDuration1
        '
        Me.LabelSetupDispenseDuration1.Location = New System.Drawing.Point(76, 112)
        Me.LabelSetupDispenseDuration1.Name = "LabelSetupDispenseDuration1"
        Me.LabelSetupDispenseDuration1.Size = New System.Drawing.Size(160, 23)
        Me.LabelSetupDispenseDuration1.TabIndex = 4
        Me.LabelSetupDispenseDuration1.Text = "Dispense Duration"
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(340, 208)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(64, 24)
        Me.Label45.TabIndex = 7
        Me.Label45.Text = "+/- mg"
        '
        'CalibrationType
        '
        Me.CalibrationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CalibrationType.Items.AddRange(New Object() {"By Desired Weight", "By Current Parameters"})
        Me.CalibrationType.Location = New System.Drawing.Point(100, 32)
        Me.CalibrationType.Name = "CalibrationType"
        Me.CalibrationType.Size = New System.Drawing.Size(280, 28)
        Me.CalibrationType.TabIndex = 48
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Location = New System.Drawing.Point(384, 440)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(80, 40)
        Me.ButtonRevert.TabIndex = 6
        Me.ButtonRevert.Text = "Revert"
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(296, 440)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(80, 40)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        '
        'LabelSetupMaterialAirPressure2
        '
        Me.LabelSetupMaterialAirPressure2.Location = New System.Drawing.Point(340, 144)
        Me.LabelSetupMaterialAirPressure2.Name = "LabelSetupMaterialAirPressure2"
        Me.LabelSetupMaterialAirPressure2.Size = New System.Drawing.Size(40, 24)
        Me.LabelSetupMaterialAirPressure2.TabIndex = 6
        Me.LabelSetupMaterialAirPressure2.Text = "Bar"
        '
        'LabelAdjustedDispenseDuration1
        '
        Me.LabelAdjustedDispenseDuration1.Enabled = False
        Me.LabelAdjustedDispenseDuration1.Location = New System.Drawing.Point(76, 400)
        Me.LabelAdjustedDispenseDuration1.Name = "LabelAdjustedDispenseDuration1"
        Me.LabelAdjustedDispenseDuration1.Size = New System.Drawing.Size(160, 23)
        Me.LabelAdjustedDispenseDuration1.TabIndex = 4
        Me.LabelAdjustedDispenseDuration1.Text = "Adjusted Duration"
        Me.LabelAdjustedDispenseDuration1.Visible = False
        '
        'LabelAdjustedMaterialAirPressure1
        '
        Me.LabelAdjustedMaterialAirPressure1.Enabled = False
        Me.LabelAdjustedMaterialAirPressure1.Location = New System.Drawing.Point(76, 400)
        Me.LabelAdjustedMaterialAirPressure1.Name = "LabelAdjustedMaterialAirPressure1"
        Me.LabelAdjustedMaterialAirPressure1.Size = New System.Drawing.Size(160, 23)
        Me.LabelAdjustedMaterialAirPressure1.TabIndex = 3
        Me.LabelAdjustedMaterialAirPressure1.Text = "Adjusted Pressure"
        Me.LabelAdjustedMaterialAirPressure1.Visible = False
        '
        'LabelAdjustedDispenseDuration2
        '
        Me.LabelAdjustedDispenseDuration2.Enabled = False
        Me.LabelAdjustedDispenseDuration2.Location = New System.Drawing.Point(340, 400)
        Me.LabelAdjustedDispenseDuration2.Name = "LabelAdjustedDispenseDuration2"
        Me.LabelAdjustedDispenseDuration2.Size = New System.Drawing.Size(32, 24)
        Me.LabelAdjustedDispenseDuration2.TabIndex = 7
        Me.LabelAdjustedDispenseDuration2.Text = "ms"
        Me.LabelAdjustedDispenseDuration2.Visible = False
        '
        'LabelAdjustedRPM1
        '
        Me.LabelAdjustedRPM1.Enabled = False
        Me.LabelAdjustedRPM1.Location = New System.Drawing.Point(76, 400)
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
        Me.AdjustedMaterialAirPressure.Location = New System.Drawing.Point(268, 400)
        Me.AdjustedMaterialAirPressure.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AdjustedMaterialAirPressure.Name = "AdjustedMaterialAirPressure"
        Me.AdjustedMaterialAirPressure.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedMaterialAirPressure.TabIndex = 17
        Me.AdjustedMaterialAirPressure.Visible = False
        '
        'AdjustedDispenseDuration
        '
        Me.AdjustedDispenseDuration.Enabled = False
        Me.AdjustedDispenseDuration.Location = New System.Drawing.Point(268, 400)
        Me.AdjustedDispenseDuration.Maximum = New Decimal(New Integer() {5000, 0, 0, 0})
        Me.AdjustedDispenseDuration.Name = "AdjustedDispenseDuration"
        Me.AdjustedDispenseDuration.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedDispenseDuration.TabIndex = 17
        Me.AdjustedDispenseDuration.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(76, 336)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 23)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Retract Speed"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(76, 304)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(160, 23)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "Retract Delay"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(340, 304)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 24)
        Me.Label6.TabIndex = 7
        Me.Label6.Text = "ms"
        '
        'RetractDelayValue
        '
        Me.RetractDelayValue.Location = New System.Drawing.Point(268, 304)
        Me.RetractDelayValue.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.RetractDelayValue.Name = "RetractDelayValue"
        Me.RetractDelayValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractDelayValue.TabIndex = 17
        Me.RetractDelayValue.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'RetractSpeedValue
        '
        Me.RetractSpeedValue.Location = New System.Drawing.Point(268, 336)
        Me.RetractSpeedValue.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.RetractSpeedValue.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.RetractSpeedValue.Name = "RetractSpeedValue"
        Me.RetractSpeedValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractSpeedValue.TabIndex = 17
        Me.RetractSpeedValue.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(76, 368)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 23)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Retract Height"
        '
        'RetractHeightValue
        '
        Me.RetractHeightValue.Location = New System.Drawing.Point(268, 368)
        Me.RetractHeightValue.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.RetractHeightValue.Name = "RetractHeightValue"
        Me.RetractHeightValue.Size = New System.Drawing.Size(72, 27)
        Me.RetractHeightValue.TabIndex = 17
        Me.RetractHeightValue.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(340, 368)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(36, 24)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "mm"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(340, 336)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(52, 24)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "mm/s"
        '
        'LabelAdjustedMaterialAirPressure2
        '
        Me.LabelAdjustedMaterialAirPressure2.Enabled = False
        Me.LabelAdjustedMaterialAirPressure2.Location = New System.Drawing.Point(340, 400)
        Me.LabelAdjustedMaterialAirPressure2.Name = "LabelAdjustedMaterialAirPressure2"
        Me.LabelAdjustedMaterialAirPressure2.Size = New System.Drawing.Size(40, 24)
        Me.LabelAdjustedMaterialAirPressure2.TabIndex = 6
        Me.LabelAdjustedMaterialAirPressure2.Text = "Bar"
        Me.LabelAdjustedMaterialAirPressure2.Visible = False
        '
        'LabelDurationStep2
        '
        Me.LabelDurationStep2.Location = New System.Drawing.Point(340, 240)
        Me.LabelDurationStep2.Name = "LabelDurationStep2"
        Me.LabelDurationStep2.Size = New System.Drawing.Size(32, 24)
        Me.LabelDurationStep2.TabIndex = 7
        Me.LabelDurationStep2.Text = "ms"
        '
        'LabelRPMStep1
        '
        Me.LabelRPMStep1.Location = New System.Drawing.Point(76, 240)
        Me.LabelRPMStep1.Name = "LabelRPMStep1"
        Me.LabelRPMStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelRPMStep1.TabIndex = 3
        Me.LabelRPMStep1.Text = "RPM Step"
        '
        'LabelPressureStep1
        '
        Me.LabelPressureStep1.Location = New System.Drawing.Point(76, 240)
        Me.LabelPressureStep1.Name = "LabelPressureStep1"
        Me.LabelPressureStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelPressureStep1.TabIndex = 4
        Me.LabelPressureStep1.Text = "Pressure Step"
        '
        'LabelPressureStep2
        '
        Me.LabelPressureStep2.Location = New System.Drawing.Point(340, 240)
        Me.LabelPressureStep2.Name = "LabelPressureStep2"
        Me.LabelPressureStep2.Size = New System.Drawing.Size(40, 24)
        Me.LabelPressureStep2.TabIndex = 6
        Me.LabelPressureStep2.Text = "Bar"
        '
        'LabelDurationStep1
        '
        Me.LabelDurationStep1.Location = New System.Drawing.Point(76, 240)
        Me.LabelDurationStep1.Name = "LabelDurationStep1"
        Me.LabelDurationStep1.Size = New System.Drawing.Size(124, 23)
        Me.LabelDurationStep1.TabIndex = 3
        Me.LabelDurationStep1.Text = "Duration Step"
        '
        'MaterialAirPressureStepValue
        '
        Me.MaterialAirPressureStepValue.DecimalPlaces = 2
        Me.MaterialAirPressureStepValue.Location = New System.Drawing.Point(268, 240)
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
        Me.RPMStepValue.Location = New System.Drawing.Point(268, 240)
        Me.RPMStepValue.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.RPMStepValue.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.RPMStepValue.Name = "RPMStepValue"
        Me.RPMStepValue.Size = New System.Drawing.Size(72, 27)
        Me.RPMStepValue.TabIndex = 17
        Me.RPMStepValue.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'DurationStepValue
        '
        Me.DurationStepValue.Location = New System.Drawing.Point(268, 240)
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
        Me.AdjustedRPM.Location = New System.Drawing.Point(268, 400)
        Me.AdjustedRPM.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.AdjustedRPM.Name = "AdjustedRPM"
        Me.AdjustedRPM.Size = New System.Drawing.Size(72, 27)
        Me.AdjustedRPM.TabIndex = 17
        Me.AdjustedRPM.Visible = False
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.Location = New System.Drawing.Point(72, 616)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(176, 40)
        Me.ButtonCalibrate.TabIndex = 6
        Me.ButtonCalibrate.Text = "Calibrate"
        '
        'ButtonTeachCalibrate
        '
        Me.ButtonTeachCalibrate.Location = New System.Drawing.Point(264, 616)
        Me.ButtonTeachCalibrate.Name = "ButtonTeachCalibrate"
        Me.ButtonTeachCalibrate.Size = New System.Drawing.Size(176, 40)
        Me.ButtonTeachCalibrate.TabIndex = 6
        Me.ButtonTeachCalibrate.Text = "Dispense"
        '
        'WeighingScaleButton
        '
        Me.WeighingScaleButton.Location = New System.Drawing.Point(288, 8)
        Me.WeighingScaleButton.Name = "WeighingScaleButton"
        Me.WeighingScaleButton.Size = New System.Drawing.Size(136, 50)
        Me.WeighingScaleButton.TabIndex = 50
        Me.WeighingScaleButton.Text = "Weighing Scale Interface"
        '
        'DispenserType
        '
        Me.DispenserType.Location = New System.Drawing.Point(144, 80)
        Me.DispenserType.Name = "DispenserType"
        Me.DispenserType.Size = New System.Drawing.Size(240, 23)
        Me.DispenserType.TabIndex = 3
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
        'rtbResult
        '
        Me.rtbResult.Location = New System.Drawing.Point(96, 88)
        Me.rtbResult.Name = "rtbResult"
        Me.rtbResult.ReadOnly = True
        Me.rtbResult.Size = New System.Drawing.Size(360, 112)
        Me.rtbResult.TabIndex = 9
        Me.rtbResult.Text = ""
        '
        'VolumeCalibrationSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(888, 878)
        Me.Controls.Add(Me.gpbDualHead)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "VolumeCalibrationSettings"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.BoxStatus.ResumeLayout(False)
        Me.BoxSettings.ResumeLayout(False)
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
    Public DispenseDuration, RPM, MaterialAirPressure, RetractHeight, RetractSpeed, RetractDelay As Double

    'this is exposed for other modules to read the success/fail status of volume calibration
    Public VolumeCalibrationResult As String
    'this is exposed for other modules to read the current status of volume calibration
    'stopped, running or paused
    Public VolumeCalibrationState As String = "Stopped"

    'options are "init", "decrease", and "increase"
    Private NextAction As String = "Init"
    Private DispenseDurationStep As Double = 200 '200ms
    Private RPMStep As Double = 5
    Private MaterialAirPressureStep As Double = 0.1
    Private TimeOutDuration As Integer = 15 'in seconds

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
                MyDispenserSettings.DownloadAugerRPM(.AdjustedRPM)
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

    'all this text can't be helped to toggle on and off the label/valuebox visibility. but if the programmer desires, he can put them into groupboxes 
    Public Sub RevertData()

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

            ElseIf CalibrateRPM() Then
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
            SetupMaterialAirPressure.Text = .Left.SetupMaterialAirPressure
            SetupRPM.Text = .Left.SetupRPM
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
        End With

    End Sub

    Public Sub SaveData()

        'With IDS.Data.Hardware.Dispenser.Left
        '    .MaterialAirPressure = AdjustedMaterialAirPressure.Value
        '    .RPM = RPM
        'End With

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
        End With

        IDS.Data.SaveData()
    End Sub

    Public Sub ButtonCalibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click

        If ButtonCalibrate.Text = "Calibrate" Then
            Cursor = Cursors.WaitCursor
            'Weighting_Scale.OpenPort()
            Cursor = Cursors.Default
            attemptCount = 0
            If Not VolumeCalibrationSetup() Then Exit Sub
            rtbResult.Clear()
            Status.Text = ""

            'override the global IDS.Data values
            MaterialAirPressure = SetupMaterialAirPressure.Value
            RPM = SetupRPM.Value
            DispenseDuration = SetupDispenseDuration.Value
            RetractSpeed = RetractSpeedValue.Value
            RetractDelay = RetractDelayValue.Value
            RetractHeight = RetractHeightValue.Value
            m_Tri.SteppingButtons.Enabled = False
            m_Tri.SetMachineRun()
            BoxSettings.Enabled = False
            ButtonCalibrate.Text = "Stop"
            VolumeCalibration(NumberOfAttempts.Value)
            ButtonCalibrate.Text = "Calibrate"
            BoxSettings.Enabled = True
            m_Tri.SetMachineStop()
            m_Tri.SteppingButtons.Enabled = True
            'disable temporarily for testing
        ElseIf ButtonCalibrate.Text = "Stop" Then
            VolumeCalibrationState = "Stopped"
            ButtonCalibrate.Text = "Calibrate"
        End If

    End Sub
    Private attemptCount As Integer = 0
    Public Function VolumeCalibration(ByVal AttemptsLeft As Integer) As Boolean

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        Dim Attempts As Integer = IDS.Data.Hardware.Volume.Left.NumberOfAttempts
        Dim DispensingResult As String

        If AttemptsLeft > 0 Then
            attemptCount += 1
            rtbResult.Text = "Attempt #" & attemptCount
            If m_Tri.EStopActivated() Or m_Tri.MachineHoming Then Return False

            If CalibrateDuration() Then
                DispensingResult = DispenseAndWeightDuration()
            ElseIf CalibratePressure() Then
                DispensingResult = DispenseAndWeightPressure()
            ElseIf CalibrateRPM() Then
                DispensingResult = DispenseAndWeightRPM()
            End If

            If DispensingResult = "Success" Then
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
                            MyDispenserSettings.DownloadAugerRPM(RPM)
                        Catch ex As Exception
                            MessageBox.Show("Error occur when download auger rpm")
                            Return False
                        End Try

                    End If
                    NextAction = "Init"
                    IDS.Data.SaveData()
                    Status.Text = "Volume calibration success."
                    Return True
                End With
            ElseIf DispensingResult = "Failed" Then
                Return VolumeCalibration(AttemptsLeft - 1)
            ElseIf DispensingResult = "Stopped" Then
                Status.Text = "Volume calibration stopped."
                NextAction = "Init"
                Return False
            End If
        Else
            Status.Text = "Volume calibration failed."
            NextAction = "Init"
            Return False
        End If

    End Function

    Public Function VolumeCalibrationSetup() As Boolean

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

        '1)move to the calibration position
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration

        With IDS.Data.Hardware.Volume.Left
            RetractHeight = .RetractHeight
            RetractSpeed = .RetractSpeed
            RetractDelay = .RetractDelay
            MaterialAirPressureStep = .PressureStepValue
        End With

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            DispenseDuration = IDS.Data.Hardware.Volume.Left.AdjustedDispenseDuration
            MaterialAirPressure = IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
        ElseIf CalibratePressure() Then
            DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
            MaterialAirPressure = IDS.Data.Hardware.Volume.Left.AdjustedMaterialAirPressure
        ElseIf CalibrateRPM() Then
            MaterialAirPressure = IDS.Data.Hardware.Volume.Left.SetupMaterialAirPressure
            DispenseDuration = IDS.Data.Hardware.Volume.Left.SetupDispenseDuration
            RPM = IDS.Data.Hardware.Volume.Left.AdjustedRPM
        End If

        'disable temporarily for testing
        Return True

StopCalibration:
        VolumeCalibrationResult = "Volume calibration setup interrupted."
        Return False

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
        If NextAction = "Decrease" Then
            MaterialAirPressure = MaterialAirPressure - MaterialAirPressureStep
        ElseIf NextAction = "Increase" Then
            MaterialAirPressure = MaterialAirPressure + MaterialAirPressureStep
        End If

        MyDispenserSettings.DownloadMaterialAirPressure(MaterialAirPressure, SuckbackPressure)
        string_result = DispenseAndWeight()
        Return string_result

    End Function

    Public Function DispenseAndWeightRPM() As String

        If NextAction = "Decrease" Then
            RPM = RPM - RPMStep
        ElseIf NextAction = "Increase" Then
            RPM = RPM + RPMStep
        End If

        MyDispenserSettings.DownloadAugerRPM(RPM)
        string_result = DispenseAndWeight()
        Return string_result

    End Function

    Public Function DispenseAndWeightDuration() As String

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If NextAction = "Decrease" Then
            DispenseDuration = DispenseDuration - DispenseDurationStep
        ElseIf NextAction = "Increase" Then
            DispenseDuration = DispenseDuration + DispenseDurationStep
        End If

        If HeadType = "Jetting Valve" Then
            MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
        ElseIf HeadType = "Slider Valve" Then
            MyDispenserSettings.DownloadJettingParameters(DispenseDuration)
        End If

        string_result = DispenseAndWeight()
        Return string_result

    End Function

    Public Function DispenseAndWeight() As String
        Dim errorcode As Integer = 0
        start_time = Now
        Status.Text = "Taring.."
        Weighting_Scale.DoTare()
        Console.WriteLine("Taring")
        'Do
        '    Sleep(5)
        '    'DisplayProgressText(Status)
        '    TraceDoEvents()
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until Not Weighting_Scale.Taring Or elapsed_time.TotalSeconds > TimeOutDuration Or Not CheckState() Or m_Tri.EStopActivated
        'If elapsed_time.TotalSeconds > TimeOutDuration Then
        '    errorcode = 1
        '    GoTo StopCalibration
        'End If
        Console.WriteLine("Tared")
        If Not CheckState() Then GoTo StopCalibration
        Status.Text = "Taring Done"
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
        If CheckState() Then
            m_Tri.Move_Z(position(2))
        Else
            GoTo StopCalibration
        End If
        Status.Text = "Dispensing"
        m_Tri.m_TriCtrl.Execute("OP(25,1)")
        m_Tri.m_TriCtrl.Execute("WA(" & DispenseDuration.ToString & ")")
        m_Tri.m_TriCtrl.Execute("OP(25,0)")
        m_Tri.m_TriCtrl.Execute("WA(" & RetractDelay.ToString & ")")
        If DoRetract() Then
            m_Tri.Set_Z_Speed(RetractSpeed)
            If CheckState() Then
                m_Tri.Move_Z(position(2) + RetractHeight)
            Else
                GoTo StopCalibration
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
        Console.WriteLine("Read Weight")
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
        Console.WriteLine("Weight Read")
        If Not CheckState() Then GoTo StopCalibration

        Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
        If CalibrateDuration() Then
            VolumeCalibrationResult = "Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
        ElseIf CalibratePressure() Then
            VolumeCalibrationResult = "Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
        ElseIf CalibrateRPM() Then
            VolumeCalibrationResult = "RPM of " + RPM.ToString + " gives a result of " + reading + " mg."
        End If

        rtbResult.Text = VolumeCalibrationResult
        With IDS.Data.Hardware.Volume.Left
            Return WithinTolerance(.Tolerance, .DesiredWeight, CDbl(reading))
        End With

StopCalibration:
        If errorcode = 1 Then
            VolumeCalibrationResult = "Weighting Scale Taring timeout."
        ElseIf errorcode = 2 Then
            VolumeCalibrationResult = "Weighting Scale reading timeout."
        Else
            VolumeCalibrationResult = "Volume calibration stopped."
        End If
        rtbResult.Text = VolumeCalibrationResult
        Return "Stopped"

WeightTooLow:
        VolumeCalibrationResult = "Volume calibration failed because weight could not be detected."
        rtbResult.Text = VolumeCalibrationResult
        Return "Failed"

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
            VolumeCalibrationState = "Stopped"
            ButtonTeachCalibrate.Text = "Dispense"
        ElseIf ButtonTeachCalibrate.Text = "Dispense" Then
            Cursor = Cursors.WaitCursor
            Weighting_Scale.OpenPort()
            Cursor = Cursors.Default
            ButtonTeachCalibrate.Text = "Stop"
            If Not VolumeCalibrationSetup() Then GoTo StopCalibration
            m_Tri.SetMachineRun()
            m_Tri.SteppingButtons.Enabled = False
            BoxSettings.Enabled = False

            Weighting_Scale.DoTare()
            start_time = Now
            Status.Text = "Taring.."
            Do
                Sleep(5)
                DisplayProgressText(Status)
                TraceDoEvents()
                stop_time = Now
                elapsed_time = stop_time.Subtract(start_time)
            Loop Until Not Weighting_Scale.Taring Or elapsed_time.TotalSeconds > TimeOutDuration Or Not CheckState() Or m_Tri.EStopActivated
            If elapsed_time.TotalSeconds > TimeOutDuration Then GoTo timeout
            Status.Text = "Tared"
            If Not CheckState() Then GoTo StopCalibration

            Dim SuckbackPressure As Double = IDS.Data.Hardware.Dispenser.Left.SuckbackPressure
            If CalibrateDuration() Then
                MyDispenserSettings.DownloadJettingParameters(SetupDispenseDuration.Value)
            ElseIf CalibratePressure() Then
                MyDispenserSettings.DownloadMaterialAirPressure(SetupMaterialAirPressure.Value, SuckbackPressure)
            ElseIf CalibrateRPM() Then
                MyDispenserSettings.DownloadAugerRPM(SetupRPM.Value)
            End If

            If Not m_Tri.Move_Z(position(2)) Then GoTo StopCalibration
            m_Tri.m_TriCtrl.Execute("OP(25,1)")
            m_Tri.m_TriCtrl.Execute("WA(" & DispenseDuration.ToString & ")")
            m_Tri.m_TriCtrl.Execute("OP(25,0)")
            m_Tri.m_TriCtrl.Execute("WA(" & RetractDelay.ToString & ")")
            If DoRetract() Then
                m_Tri.Set_Z_Speed(RetractSpeed)
                If Not m_Tri.Move_Z(position(2) + RetractHeight) Then GoTo StopCalibration
            End If
            m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
            If Not m_Tri.Move_Z(0) Then GoTo StopCalibration

            RetractSpeed = RetractSpeedValue.Value
            RetractDelay = RetractDelayValue.Value
            RetractHeight = RetractHeightValue.Value

            '3)read the weighted value
            For i As Integer = 1 To 10
                Sleep(5)
                TraceDoEvents()
            Next
            start_time = Now
            Weighting_Scale.GetWeight()
            Status.Text = "Reading.."
            Do
                Sleep(5)
                TraceDoEvents()
                DisplayProgressText(Status)
                stop_time = Now
                elapsed_time = stop_time.Subtract(start_time)
            Loop Until Weighting_Scale.ValueUpdated Or elapsed_time.TotalSeconds > TimeOutDuration Or Not CheckState() Or m_Tri.EStopActivated
            Dim reading As String = CStr(Weighting_Scale.WeightReading)
            Weighting_Scale.ResetValues()
            If elapsed_time.TotalSeconds > TimeOutDuration Then GoTo Timeout
            If Not CheckState() Then GoTo StopCalibration

            Dim HeadType As String = IDS.Data.Hardware.Dispenser.Left.HeadType
            If CalibrateDuration() Then
                VolumeCalibrationResult = "Dispensing duration of " + DispenseDuration.ToString + " ms gives a result of " + reading + " mg."
            ElseIf CalibratePressure() Then
                VolumeCalibrationResult = "Pressure of " + MaterialAirPressure.ToString + " bar gives a result of " + reading + " mg."
            ElseIf CalibrateRPM() Then
                VolumeCalibrationResult = "RPM of " + RPM.ToString + " gives a result of " + reading + " mg."
            End If
            Status.Text = "Dispensing done"
        End If

        BoxSettings.Enabled = True
        m_Tri.SetMachineStop()
        rtbResult.Text = VolumeCalibrationResult
        ButtonTeachCalibrate.Text = "Dispense"
        m_Tri.SteppingButtons.Enabled = True
        Exit Sub

StopCalibration:
        BoxSettings.Enabled = True
        m_Tri.SetMachineStop()
        rtbResult.Text = "Machine prematurely stopped."
        ButtonTeachCalibrate.Text = "Dispense"
        Status.Text = ""
        m_Tri.SteppingButtons.Enabled = True
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
        If WeightingScaleForm.Visible Then
            WeightingScaleForm.Hide()
        Else
            WeightingScaleForm.Show()
        End If

    End Sub

End Class
