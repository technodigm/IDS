Imports System.Windows.Forms

Public Class Heater
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
    Friend WithEvents AxMSComm1 As AxMSCommLib.AxMSComm
    Friend WithEvents CB_PostThermal As System.Windows.Forms.CheckBox
    Friend WithEvents CB_NeedleHeater As System.Windows.Forms.CheckBox
    Friend WithEvents CB_SyringeHeater As System.Windows.Forms.CheckBox
    Friend WithEvents CB_InThermal As System.Windows.Forms.CheckBox
    Friend WithEvents CB_PreThermal As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents NeedleOperationTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents SyringeOperationTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents InThermalOperationTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents PreThermalOperationTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents PostThermalOperationTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents InThermalStandbyTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents PreThermalStandbyTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents PostThermalStandbyTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents SyringeStandbyTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents NeedleStandbyTemp As System.Windows.Forms.NumericUpDown
    Friend WithEvents SyringeThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents NeedleThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents InThermalThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents PreThermalThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents PostThermalThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonStart As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents InThermalTemp As System.Windows.Forms.Label
    Friend WithEvents PreThermalTemp As System.Windows.Forms.Label
    Friend WithEvents SyringeTemp As System.Windows.Forms.Label
    Friend WithEvents NeedleTemp As System.Windows.Forms.Label
    Friend WithEvents PostThermalTemp As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Heater))
        Me.AxMSComm1 = New AxMSCommLib.AxMSComm
        Me.CB_PostThermal = New System.Windows.Forms.CheckBox
        Me.CB_NeedleHeater = New System.Windows.Forms.CheckBox
        Me.CB_SyringeHeater = New System.Windows.Forms.CheckBox
        Me.CB_InThermal = New System.Windows.Forms.CheckBox
        Me.CB_PreThermal = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.NeedleOperationTemp = New System.Windows.Forms.NumericUpDown
        Me.SyringeOperationTemp = New System.Windows.Forms.NumericUpDown
        Me.InThermalOperationTemp = New System.Windows.Forms.NumericUpDown
        Me.PreThermalOperationTemp = New System.Windows.Forms.NumericUpDown
        Me.PostThermalOperationTemp = New System.Windows.Forms.NumericUpDown
        Me.InThermalStandbyTemp = New System.Windows.Forms.NumericUpDown
        Me.PreThermalStandbyTemp = New System.Windows.Forms.NumericUpDown
        Me.PostThermalStandbyTemp = New System.Windows.Forms.NumericUpDown
        Me.SyringeStandbyTemp = New System.Windows.Forms.NumericUpDown
        Me.NeedleStandbyTemp = New System.Windows.Forms.NumericUpDown
        Me.SyringeThreshold = New System.Windows.Forms.NumericUpDown
        Me.NeedleThreshold = New System.Windows.Forms.NumericUpDown
        Me.InThermalThreshold = New System.Windows.Forms.NumericUpDown
        Me.PreThermalThreshold = New System.Windows.Forms.NumericUpDown
        Me.PostThermalThreshold = New System.Windows.Forms.NumericUpDown
        Me.ButtonStart = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.InThermalTemp = New System.Windows.Forms.Label
        Me.PreThermalTemp = New System.Windows.Forms.Label
        Me.SyringeTemp = New System.Windows.Forms.Label
        Me.NeedleTemp = New System.Windows.Forms.Label
        Me.PostThermalTemp = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NeedleOperationTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyringeOperationTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InThermalOperationTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PreThermalOperationTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PostThermalOperationTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InThermalStandbyTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PreThermalStandbyTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PostThermalStandbyTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyringeStandbyTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NeedleStandbyTemp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SyringeThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NeedleThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.InThermalThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PreThermalThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PostThermalThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxMSComm1
        '
        Me.AxMSComm1.Enabled = True
        Me.AxMSComm1.Location = New System.Drawing.Point(0, 0)
        Me.AxMSComm1.Name = "AxMSComm1"
        Me.AxMSComm1.OcxState = CType(resources.GetObject("AxMSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMSComm1.Size = New System.Drawing.Size(38, 38)
        Me.AxMSComm1.TabIndex = 0
        '
        'CB_PostThermal
        '
        Me.CB_PostThermal.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_PostThermal.Location = New System.Drawing.Point(8, 296)
        Me.CB_PostThermal.Name = "CB_PostThermal"
        Me.CB_PostThermal.Size = New System.Drawing.Size(132, 24)
        Me.CB_PostThermal.TabIndex = 46
        Me.CB_PostThermal.Text = "Post Thermal"
        '
        'CB_NeedleHeater
        '
        Me.CB_NeedleHeater.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_NeedleHeater.Location = New System.Drawing.Point(8, 104)
        Me.CB_NeedleHeater.Name = "CB_NeedleHeater"
        Me.CB_NeedleHeater.Size = New System.Drawing.Size(140, 24)
        Me.CB_NeedleHeater.TabIndex = 47
        Me.CB_NeedleHeater.Text = "Needle Heater"
        '
        'CB_SyringeHeater
        '
        Me.CB_SyringeHeater.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_SyringeHeater.Location = New System.Drawing.Point(8, 152)
        Me.CB_SyringeHeater.Name = "CB_SyringeHeater"
        Me.CB_SyringeHeater.Size = New System.Drawing.Size(148, 24)
        Me.CB_SyringeHeater.TabIndex = 48
        Me.CB_SyringeHeater.Text = "Syringe Heater"
        '
        'CB_InThermal
        '
        Me.CB_InThermal.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_InThermal.Location = New System.Drawing.Point(8, 248)
        Me.CB_InThermal.Name = "CB_InThermal"
        Me.CB_InThermal.Size = New System.Drawing.Size(116, 24)
        Me.CB_InThermal.TabIndex = 45
        Me.CB_InThermal.Text = "In Thermal"
        '
        'CB_PreThermal
        '
        Me.CB_PreThermal.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_PreThermal.Location = New System.Drawing.Point(8, 200)
        Me.CB_PreThermal.Name = "CB_PreThermal"
        Me.CB_PreThermal.Size = New System.Drawing.Size(124, 24)
        Me.CB_PreThermal.TabIndex = 44
        Me.CB_PreThermal.Text = "Pre Thermal"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(200, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(320, 32)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "Temperature Units: Degree Celcius"
        '
        'Button5
        '
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(8, 352)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(152, 48)
        Me.Button5.TabIndex = 42
        Me.Button5.Text = "Test Communications"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(160, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 24)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "Operation Temp"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(296, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 24)
        Me.Label3.TabIndex = 43
        Me.Label3.Text = "Standby Temp"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(432, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 24)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Threshold"
        '
        'NeedleOperationTemp
        '
        Me.NeedleOperationTemp.Location = New System.Drawing.Point(168, 104)
        Me.NeedleOperationTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NeedleOperationTemp.Name = "NeedleOperationTemp"
        Me.NeedleOperationTemp.Size = New System.Drawing.Size(104, 27)
        Me.NeedleOperationTemp.TabIndex = 49
        Me.NeedleOperationTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NeedleOperationTemp.Visible = False
        '
        'SyringeOperationTemp
        '
        Me.SyringeOperationTemp.Location = New System.Drawing.Point(168, 152)
        Me.SyringeOperationTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.SyringeOperationTemp.Name = "SyringeOperationTemp"
        Me.SyringeOperationTemp.Size = New System.Drawing.Size(104, 27)
        Me.SyringeOperationTemp.TabIndex = 49
        Me.SyringeOperationTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.SyringeOperationTemp.Visible = False
        '
        'InThermalOperationTemp
        '
        Me.InThermalOperationTemp.Location = New System.Drawing.Point(168, 248)
        Me.InThermalOperationTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.InThermalOperationTemp.Name = "InThermalOperationTemp"
        Me.InThermalOperationTemp.Size = New System.Drawing.Size(104, 27)
        Me.InThermalOperationTemp.TabIndex = 49
        Me.InThermalOperationTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.InThermalOperationTemp.Visible = False
        '
        'PreThermalOperationTemp
        '
        Me.PreThermalOperationTemp.Location = New System.Drawing.Point(168, 200)
        Me.PreThermalOperationTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PreThermalOperationTemp.Name = "PreThermalOperationTemp"
        Me.PreThermalOperationTemp.Size = New System.Drawing.Size(104, 27)
        Me.PreThermalOperationTemp.TabIndex = 49
        Me.PreThermalOperationTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PreThermalOperationTemp.Visible = False
        '
        'PostThermalOperationTemp
        '
        Me.PostThermalOperationTemp.Location = New System.Drawing.Point(168, 296)
        Me.PostThermalOperationTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PostThermalOperationTemp.Name = "PostThermalOperationTemp"
        Me.PostThermalOperationTemp.Size = New System.Drawing.Size(104, 27)
        Me.PostThermalOperationTemp.TabIndex = 49
        Me.PostThermalOperationTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PostThermalOperationTemp.Visible = False
        '
        'InThermalStandbyTemp
        '
        Me.InThermalStandbyTemp.Location = New System.Drawing.Point(304, 248)
        Me.InThermalStandbyTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.InThermalStandbyTemp.Name = "InThermalStandbyTemp"
        Me.InThermalStandbyTemp.Size = New System.Drawing.Size(104, 27)
        Me.InThermalStandbyTemp.TabIndex = 49
        Me.InThermalStandbyTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.InThermalStandbyTemp.Visible = False
        '
        'PreThermalStandbyTemp
        '
        Me.PreThermalStandbyTemp.Location = New System.Drawing.Point(304, 200)
        Me.PreThermalStandbyTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PreThermalStandbyTemp.Name = "PreThermalStandbyTemp"
        Me.PreThermalStandbyTemp.Size = New System.Drawing.Size(104, 27)
        Me.PreThermalStandbyTemp.TabIndex = 49
        Me.PreThermalStandbyTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PreThermalStandbyTemp.Visible = False
        '
        'PostThermalStandbyTemp
        '
        Me.PostThermalStandbyTemp.Location = New System.Drawing.Point(304, 296)
        Me.PostThermalStandbyTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PostThermalStandbyTemp.Name = "PostThermalStandbyTemp"
        Me.PostThermalStandbyTemp.Size = New System.Drawing.Size(104, 27)
        Me.PostThermalStandbyTemp.TabIndex = 49
        Me.PostThermalStandbyTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.PostThermalStandbyTemp.Visible = False
        '
        'SyringeStandbyTemp
        '
        Me.SyringeStandbyTemp.Location = New System.Drawing.Point(304, 152)
        Me.SyringeStandbyTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.SyringeStandbyTemp.Name = "SyringeStandbyTemp"
        Me.SyringeStandbyTemp.Size = New System.Drawing.Size(104, 27)
        Me.SyringeStandbyTemp.TabIndex = 49
        Me.SyringeStandbyTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.SyringeStandbyTemp.Visible = False
        '
        'NeedleStandbyTemp
        '
        Me.NeedleStandbyTemp.Location = New System.Drawing.Point(304, 104)
        Me.NeedleStandbyTemp.Minimum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NeedleStandbyTemp.Name = "NeedleStandbyTemp"
        Me.NeedleStandbyTemp.Size = New System.Drawing.Size(104, 27)
        Me.NeedleStandbyTemp.TabIndex = 49
        Me.NeedleStandbyTemp.Value = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NeedleStandbyTemp.Visible = False
        '
        'SyringeThreshold
        '
        Me.SyringeThreshold.Location = New System.Drawing.Point(440, 152)
        Me.SyringeThreshold.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.SyringeThreshold.Name = "SyringeThreshold"
        Me.SyringeThreshold.Size = New System.Drawing.Size(104, 27)
        Me.SyringeThreshold.TabIndex = 49
        Me.SyringeThreshold.Visible = False
        '
        'NeedleThreshold
        '
        Me.NeedleThreshold.Location = New System.Drawing.Point(440, 104)
        Me.NeedleThreshold.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.NeedleThreshold.Name = "NeedleThreshold"
        Me.NeedleThreshold.Size = New System.Drawing.Size(104, 27)
        Me.NeedleThreshold.TabIndex = 49
        Me.NeedleThreshold.Visible = False
        '
        'InThermalThreshold
        '
        Me.InThermalThreshold.Location = New System.Drawing.Point(440, 248)
        Me.InThermalThreshold.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.InThermalThreshold.Name = "InThermalThreshold"
        Me.InThermalThreshold.Size = New System.Drawing.Size(104, 27)
        Me.InThermalThreshold.TabIndex = 49
        Me.InThermalThreshold.Visible = False
        '
        'PreThermalThreshold
        '
        Me.PreThermalThreshold.Location = New System.Drawing.Point(440, 200)
        Me.PreThermalThreshold.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.PreThermalThreshold.Name = "PreThermalThreshold"
        Me.PreThermalThreshold.Size = New System.Drawing.Size(104, 27)
        Me.PreThermalThreshold.TabIndex = 49
        Me.PreThermalThreshold.Visible = False
        '
        'PostThermalThreshold
        '
        Me.PostThermalThreshold.Location = New System.Drawing.Point(440, 296)
        Me.PostThermalThreshold.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.PostThermalThreshold.Name = "PostThermalThreshold"
        Me.PostThermalThreshold.Size = New System.Drawing.Size(104, 27)
        Me.PostThermalThreshold.TabIndex = 49
        Me.PostThermalThreshold.Visible = False
        '
        'ButtonStart
        '
        Me.ButtonStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStart.Location = New System.Drawing.Point(176, 352)
        Me.ButtonStart.Name = "ButtonStart"
        Me.ButtonStart.Size = New System.Drawing.Size(120, 48)
        Me.ButtonStart.TabIndex = 42
        Me.ButtonStart.Text = "Start Reading"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 512)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 23)
        Me.Label9.TabIndex = 53
        Me.Label9.Text = "In Thermal"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 480)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(104, 23)
        Me.Label10.TabIndex = 52
        Me.Label10.Text = "Pre Thermal"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 448)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 51
        Me.Label5.Text = "Syringe Heater"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 416)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 23)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Needle Heater"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 544)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 23)
        Me.Label6.TabIndex = 53
        Me.Label6.Text = "Post Thermal"
        '
        'InThermalTemp
        '
        Me.InThermalTemp.BackColor = System.Drawing.SystemColors.Control
        Me.InThermalTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.InThermalTemp.Location = New System.Drawing.Point(168, 512)
        Me.InThermalTemp.Name = "InThermalTemp"
        Me.InThermalTemp.Size = New System.Drawing.Size(32, 23)
        Me.InThermalTemp.TabIndex = 53
        Me.InThermalTemp.Visible = False
        '
        'PreThermalTemp
        '
        Me.PreThermalTemp.Location = New System.Drawing.Point(168, 480)
        Me.PreThermalTemp.Name = "PreThermalTemp"
        Me.PreThermalTemp.Size = New System.Drawing.Size(32, 23)
        Me.PreThermalTemp.TabIndex = 52
        Me.PreThermalTemp.Visible = False
        '
        'SyringeTemp
        '
        Me.SyringeTemp.Location = New System.Drawing.Point(168, 448)
        Me.SyringeTemp.Name = "SyringeTemp"
        Me.SyringeTemp.Size = New System.Drawing.Size(32, 23)
        Me.SyringeTemp.TabIndex = 51
        Me.SyringeTemp.Visible = False
        '
        'NeedleTemp
        '
        Me.NeedleTemp.Location = New System.Drawing.Point(168, 416)
        Me.NeedleTemp.Name = "NeedleTemp"
        Me.NeedleTemp.Size = New System.Drawing.Size(32, 23)
        Me.NeedleTemp.TabIndex = 50
        Me.NeedleTemp.Visible = False
        '
        'PostThermalTemp
        '
        Me.PostThermalTemp.BackColor = System.Drawing.SystemColors.Control
        Me.PostThermalTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PostThermalTemp.Location = New System.Drawing.Point(168, 544)
        Me.PostThermalTemp.Name = "PostThermalTemp"
        Me.PostThermalTemp.Size = New System.Drawing.Size(32, 23)
        Me.PostThermalTemp.TabIndex = 53
        Me.PostThermalTemp.Visible = False
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(320, 352)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(120, 48)
        Me.Button1.TabIndex = 42
        Me.Button1.Text = "Update Settings"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(464, 352)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(120, 48)
        Me.Button2.TabIndex = 42
        Me.Button2.Text = "Standby Mode"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(280, 488)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(144, 27)
        Me.TextBox1.TabIndex = 54
        Me.TextBox1.Text = ""
        '
        'Heater
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(600, 590)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.NeedleOperationTemp)
        Me.Controls.Add(Me.CB_PostThermal)
        Me.Controls.Add(Me.CB_NeedleHeater)
        Me.Controls.Add(Me.CB_SyringeHeater)
        Me.Controls.Add(Me.CB_InThermal)
        Me.Controls.Add(Me.CB_PreThermal)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.AxMSComm1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.SyringeOperationTemp)
        Me.Controls.Add(Me.InThermalOperationTemp)
        Me.Controls.Add(Me.PreThermalOperationTemp)
        Me.Controls.Add(Me.PostThermalOperationTemp)
        Me.Controls.Add(Me.InThermalStandbyTemp)
        Me.Controls.Add(Me.PreThermalStandbyTemp)
        Me.Controls.Add(Me.PostThermalStandbyTemp)
        Me.Controls.Add(Me.SyringeStandbyTemp)
        Me.Controls.Add(Me.NeedleStandbyTemp)
        Me.Controls.Add(Me.SyringeThreshold)
        Me.Controls.Add(Me.NeedleThreshold)
        Me.Controls.Add(Me.InThermalThreshold)
        Me.Controls.Add(Me.PreThermalThreshold)
        Me.Controls.Add(Me.PostThermalThreshold)
        Me.Controls.Add(Me.ButtonStart)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.InThermalTemp)
        Me.Controls.Add(Me.PreThermalTemp)
        Me.Controls.Add(Me.SyringeTemp)
        Me.Controls.Add(Me.NeedleTemp)
        Me.Controls.Add(Me.PostThermalTemp)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "Heater"
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NeedleOperationTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyringeOperationTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InThermalOperationTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PreThermalOperationTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PostThermalOperationTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InThermalStandbyTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PreThermalStandbyTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PostThermalStandbyTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyringeStandbyTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NeedleStandbyTemp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SyringeThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NeedleThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.InThermalThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PreThermalThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PostThermalThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim command As String
    Dim byte_result As Byte()
    Dim raw_result As Object
    Dim byte_num, BCC As Byte
    Dim output_str As String
    Dim node_num As Char
    Public Thermal_Comm_Error As Boolean() = {False, False, False, False, False}
    Public Temperature_Alarm As Boolean() = {False, False, False, False, False}
    Public Temperature As Double() = {0, 0, 0, 0, 0}
    Dim node_num_int, echoback_count As Integer
    Dim k As Integer = 1 'node number when reading repeatedly

    Public Sub OpenPort()
        If AxMSComm1.PortOpen = True Then
            Exit Sub
        End If
        Try
            With AxMSComm1
                .CommPort = HeaterPort
                .Settings = "9600,E,7,2"
                .RThreshold = 27 'the threshold limit for receiving characters
                .InBufferSize = 27
                .InputLen = 27
                .InputMode = MSCommLib.InputModeConstants.comInputModeBinary
                .InBufferCount = 0
                If Not .PortOpen Then
                    .PortOpen = True
                End If
            End With
        Catch ex As System.Runtime.InteropServices.COMException
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Sub GetHeaterError(ByRef ID As String, ByVal HeaterEnabled As Boolean())

        'Dim index As Integer
        'If counter2 = 0 Then 'if counter 2 is ticking then stop checking, use current values only to read through the whole error log
        '    For z As Integer = 0 To HeaterEnabled.Length - 1
        '        If HeaterEnabled(z) = True And Thermal_Comm_Error(z) = True And heater_on_or_off(z) = True Then
        '            ID = Convert.ToInt32(ID) + z + 11
        '            ReDim Preserve offenders(counter3)
        '            offenders(counter3) = Convert.ToInt32(ID)
        '            ID = Convert.ToInt32(ID) - z - 11
        '            counter3 += 1
        '        End If
        '    Next
        '    If Not offenders Is Nothing Then
        '        If offenders(0) = 0 Then
        '            index = offenders.Length + 1
        '        Else
        '            index = offenders.Length
        '        End If
        '    Else
        '        index = 0
        '    End If
        '    For y As Integer = 0 To HeaterEnabled.Length - 1
        '        If HeaterEnabled(y) = True And Temperature_Alarm(y) = True And heater_on_or_off(y) = True Then
        '            ID = Convert.ToInt32(ID) + y + 1
        '            ReDim Preserve offenders(index)
        '            offenders(index) = Convert.ToInt32(ID)
        '            ID = Convert.ToInt32(ID) - y - 1
        '            counter3 += 1
        '            index += 1
        '        End If
        '    Next
        'End If

        'If offenders Is Nothing Then
        '    Exit Sub
        'ElseIf offenders.Length = 1 And offenders(counter2) = 0 Then
        '    Exit Sub
        'Else
        '    If offenders(counter2) = 0 Then
        '        counter2 += 1
        '        Exit Sub
        '    End If
        '    ID = offenders(counter2)
        '    counter2 += 1
        'End If

        'If counter2 = offenders.Length Then 'reset when out of index! i.e. offenders is length 2, so should stop when counter2 is 2 or offenders(2) = error
        '    counter2 = 0
        '    counter3 = 0
        '    ReDim offenders(0)
        'End If

    End Sub

    Private Sub Heater_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OpenPort()
    End Sub

    Public Sub echoback_test(ByVal node_num As Integer)

        'i.e.             output_str = Chr(2) & "01000" & "080106" & Chr(3) & Chr(BCC)
        'the 06 in 080106 is the send data. you should see 06 in dec or 54 in char when receiving the echo.
        'STX,NodeNo,Sub,SID,&add,VerType,&Add,&BitPos+NoOfElements,&ETX,&BCC      (should response 25 bytes)

        output_str = Chr(2)
        BCC = CByte(&H30)

        If node_num = 1 Then
            BCC = BCC Xor CByte(&H31)
            output_str = output_str & "01000"
        ElseIf node_num = 2 Then
            BCC = BCC Xor CByte(&H32)
            output_str = output_str & "02000"
        ElseIf node_num = 3 Then
            BCC = BCC Xor CByte(&H33)
            output_str = output_str & "03000"
        ElseIf node_num = 4 Then
            BCC = BCC Xor CByte(&H34)
            output_str = output_str & "04000"
        ElseIf node_num = 5 Then
            BCC = BCC Xor CByte(&H35)
            output_str = output_str & "05000"
        End If

        BCC = BCC Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H38) Xor CByte(&H30) Xor CByte(&H31) Xor CByte(&H30) Xor CByte(&H36) Xor CByte(&H3)
        output_str = output_str & "080106" & Chr(3) & Chr(BCC)

        AxMSComm1.Output = output_str
    End Sub

    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function
    Public Sub echoback()
        AxMSComm1.InBufferCount = 0
        command = "echoback test"
        For i As Integer = 1 To 5
            echoback_test(i)
            Sleep(100)
        Next
    End Sub


    Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm
        'read received data here

        With AxMSComm1
            Select Case .CommEvent
                Case 2
                Case Else
                    .RThreshold = 0

                    Dim error_flag As Boolean = False
                    Dim temp_or_alarm As String = ""
                    Dim error_check As Integer() = {48, 48, 48, 48, 48, 49, 48, 49, 48, 48, 48, 48}
                    Dim i, error_byte_int1, error_byte_int2, error_byte_int3, error_byte_int4 As Integer

                    'check echoback
                    If command = "Echoback" Then

                        raw_result = AxMSComm1.Input
                        byte_result = CType(raw_result, Byte())

                        'mrc, src, response code, 0-23 bytes of data (in this case its 6)
                        'response code in chr is NULL STX 01 0000 0801 0000 06 ETX BCC NULL (21 bytes)
                        byte_num = 0
                        If byte_result(0) = 2 Then
                            node_num = Convert.ToChar(byte_result(2))
                        ElseIf byte_result(1) = 2 Then
                            node_num = Convert.ToChar(byte_result(3))
                        End If

                        node_num_int = Integer.Parse(node_num)
                        error_byte_int1 = CInt(byte_result(byte_num + 12))
                        error_byte_int2 = CInt(byte_result(byte_num + 13))
                        error_byte_int3 = CInt(byte_result(byte_num + 14))
                        error_byte_int4 = CInt(byte_result(byte_num + 15))
                        If error_byte_int1 = 48 And error_byte_int2 = 48 And error_byte_int3 = 48 And error_byte_int4 = 48 Then
                            If echoback_count = 5 Then
                                echoback_count = 1
                                command = "Finished"
                            Else
                                echoback_count += 1
                            End If
                            MsgBox("Heater " & node_num & " is working.")
                        Else
                            MsgBox("Heater " & node_num & " is not working.")
                        End If
                    End If

                    'read temperature
                    If AxMSComm1.InBufferSize > 25 Then

                        raw_result = AxMSComm1.Input
                        byte_result = CType(raw_result, Byte())

                        If byte_result(0) = 2 Then 'start byte is 2, STX 
                            i = 0
                        ElseIf byte_result(1) = 2 Then
                            i = 1
                        End If

                        If byte_result(23 + i) = 3 Then 'end byte is 3, ETX

                            node_num = Convert.ToChar(byte_result(2 + i))  'first check to see node number

                            For k As Integer = 0 To 11 'second check to see if match pattern
                                If Not byte_result(3 + i + k) = error_check(k) Then
                                    error_flag = True 'communication fail error!
                                    Thermal_Comm_Error(Val(node_num)) = True
                                End If
                            Next

                            If byte_result(16 + i) = 50 Then   'WARNING there is a possibility that it is not 50 or 51!
                                'heater_on_or_off(node_num) = True '50/true for run
                                temp_or_alarm = "alarm"
                            ElseIf byte_result(16 + i) = 51 Then
                                temp_or_alarm = "alarm"
                                'heater_on_or_off(node_num) = False '51/false for stop
                            ElseIf byte_result(16 + i) = 48 Then
                                temp_or_alarm = "temp"
                            Else
                                Thermal_Comm_Error(Val(node_num)) = True
                            End If

                            If temp_or_alarm = "temp" And error_flag = False Then 'third check to see whether read temp or read alarm
                                Temperature(Val(node_num) - 1) = C_D(Convert.ToChar(byte_result(22 + i))) + C_D(Convert.ToChar(byte_result(21 + i))) * 16 + C_D(Convert.ToChar(byte_result(20 + i))) * 256
                            ElseIf temp_or_alarm = "alarm" And error_flag = False Then
                                Dim error_byte As Integer = byte_result(19)
                                If error_byte = 50 Or error_byte = 51 Or error_byte = 54 Or error_byte = 55 Or error_byte = 65 Or error_byte = 66 Or error_byte = 69 Or error_byte = 70 Then '0 1 2 3 4 -> 1 2 3 4 5
                                    Temperature_Alarm(Val(node_num)) = True
                                Else
                                    Temperature_Alarm(Val(node_num)) = False
                                End If
                            End If
                        End If

                        NeedleTemp.Text = Temperature(0).ToString
                        SyringeTemp.Text = Temperature(1).ToString
                        PreThermalTemp.Text = Temperature(2).ToString
                        InThermalTemp.Text = Temperature(3).ToString
                        PostThermalTemp.Text = Temperature(4).ToString

                        If ButtonStart.Text = "Reading.." Then
                            k += 1
                            If k = 6 Then k = 1
                            TextBox1.Text = "k is :" + k.ToString
                            Get_Temperature(k)
                        End If
                    End If

                    .RThreshold = 1
            End Select
        End With

    End Sub

    Private Function C_D(ByVal letter As String) As Double
        Dim actual_value As Double
        If letter = "0" Or letter = "1" Or letter = "2" Or letter = "3" Or letter = "4" Or letter = "5" Or letter = "6" Or letter = "7" Or letter = "8" Or letter = "9" Then
            actual_value = Val(letter)
        ElseIf letter = "A" Then : actual_value = 10 : ElseIf letter = "B" Then : actual_value = 11
        ElseIf letter = "C" Then : actual_value = 12 : ElseIf letter = "D" Then : actual_value = 13
        ElseIf letter = "E" Then : actual_value = 14 : ElseIf letter = "F" Then : actual_value = 15
        End If
        Return actual_value
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        echoback()
        command = "Echoback"
    End Sub

    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click
        'Get_Temperature(1)
        If ButtonStart.Text = "Start Reading" Then
            ButtonStart.Text = "Reading.."
            k = 1
            Get_Temperature(k)
        Else
            ButtonStart.Text = "Start Reading"
        End If

    End Sub

    Public Sub Get_Temperature(ByVal NodeNumber As Integer)

        BCC = &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H31 Xor &H43 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H3
        'Dim tezt As Byte = &H30 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H44
        'Dim bezt As Byte = &H30 Xor &H35 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H35 Xor &H32 Xor &H38

        'STX,NodeNo,Sub,SID,&add,VerType,&Add,&BitPos+NoOfElements,&ETX,&BCC      (should response 25 bytes)

        output_str = Chr(2)

        If NodeNumber = 1 Then
            BCC = BCC Xor CByte(&H31)
            output_str = output_str & "01000"
        ElseIf NodeNumber = 2 Then
            BCC = BCC Xor CByte(&H32)
            output_str = output_str & "02000"
        ElseIf NodeNumber = 3 Then
            BCC = BCC Xor CByte(&H33)
            output_str = output_str & "03000"
        ElseIf NodeNumber = 4 Then
            BCC = BCC Xor CByte(&H34)
            output_str = output_str & "04000"
        ElseIf NodeNumber = 5 Then
            BCC = BCC Xor CByte(&H35)
            output_str = output_str & "05000"
        End If

        output_str = output_str & "0101C0" & "0000" & "000001" & Chr(3) & Chr(BCC)
        AxMSComm1.Output = output_str
        Sleep(50)
        Application.DoEvents()

    End Sub

    Private Sub CB_NeedleHeater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_NeedleHeater.Click
        If CB_NeedleHeater.Checked Then
            NeedleOperationTemp.Visible = True
            NeedleStandbyTemp.Visible = True
            NeedleThreshold.Visible = True
            NeedleTemp.Visible = True
        Else
            NeedleOperationTemp.Visible = False
            NeedleStandbyTemp.Visible = False
            NeedleThreshold.Visible = False
            NeedleTemp.Visible = False
        End If
    End Sub

    Private Sub CB_SyringeHeater_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_SyringeHeater.Click
        If CB_SyringeHeater.Checked Then
            SyringeOperationTemp.Visible = True
            SyringeStandbyTemp.Visible = True
            SyringeThreshold.Visible = True
            SyringeTemp.Visible = True
        Else
            SyringeOperationTemp.Visible = False
            SyringeStandbyTemp.Visible = False
            SyringeThreshold.Visible = False
            SyringeTemp.Visible = False
        End If
    End Sub

    Private Sub CB_PreThermal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_PreThermal.Click

        If CB_PreThermal.Checked Then
            PreThermalOperationTemp.Visible = True
            PreThermalStandbyTemp.Visible = True
            PreThermalThreshold.Visible = True
            PreThermalTemp.Visible = True
        Else
            PreThermalOperationTemp.Visible = False
            PreThermalStandbyTemp.Visible = False
            PreThermalThreshold.Visible = False
            PreThermalTemp.Visible = False
        End If
    End Sub

    Private Sub CB_InThermal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_InThermal.Click

        If CB_InThermal.Checked Then
            InThermalOperationTemp.Visible = True
            InThermalStandbyTemp.Visible = True
            InThermalThreshold.Visible = True
            InThermalTemp.Visible = True
        Else
            InThermalOperationTemp.Visible = False
            InThermalStandbyTemp.Visible = False
            InThermalThreshold.Visible = False
            InThermalTemp.Visible = False
        End If
    End Sub

    Private Sub CB_PostThermal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_PostThermal.Click

        If CB_PostThermal.Checked Then
            PostThermalOperationTemp.Visible = True
            PostThermalStandbyTemp.Visible = True
            PostThermalThreshold.Visible = True
            PostThermalTemp.Visible = True
        Else
            PostThermalOperationTemp.Visible = False
            PostThermalStandbyTemp.Visible = False
            PostThermalThreshold.Visible = False
            PostThermalTemp.Visible = False
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If CB_NeedleHeater.Checked Then
            Set_Temperature(1, CByte(NeedleOperationTemp.Value))
        End If
        If CB_SyringeHeater.Checked Then
            Set_Temperature(2, CByte(SyringeOperationTemp.Value))
        End If
        If CB_PreThermal.Checked Then
            Set_Temperature(3, CByte(PreThermalOperationTemp.Value))
        End If
        If CB_InThermal.Checked Then
            Set_Temperature(4, CByte(InThermalOperationTemp.Value))
        End If
        If CB_PostThermal.Checked Then
            Set_Temperature(5, CByte(PostThermalOperationTemp.Value))
        End If

    End Sub

    Public Sub Set_Temperature(ByVal NodeNumber As Integer, ByVal Set_Value As Byte)
        Dim output_str As String
        Dim CommandValue As String

        CommandValue = "000000"

        BCC = &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H32 Xor &H43 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H33 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H3 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 'commandvalue "000000"
        If Convert.ToByte(Fix(Set_Value / 16)) = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Convert.ToByte(Fix(Set_Value / 16)) = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If Set_Value Mod 16 = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Set_Value Mod 16 = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Set_Value Mod 16 = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Set_Value Mod 16 = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Set_Value Mod 16 = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Set_Value Mod 16 = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Set_Value Mod 16 = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Set_Value Mod 16 = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Set_Value Mod 16 = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Set_Value Mod 16 = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Set_Value Mod 16 = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Set_Value Mod 16 = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Set_Value Mod 16 = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Set_Value Mod 16 = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Set_Value Mod 16 = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Set_Value Mod 16 = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If NodeNumber = 1 Then
            BCC = ((BCC Xor CByte(&H30) Xor CByte(&H31)))
            output_str = Chr(2) & "01000" & "0102C1" & "0003" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 2 Then
            BCC = ((BCC Xor CByte(&H30) Xor CByte(&H32)))
            output_str = Chr(2) & "02000" & "0102C1" & "0003" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 3 Then
            BCC = ((BCC Xor CByte(&H30) Xor CByte(&H33)))
            output_str = Chr(2) & "03000" & "0102C1" & "0003" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 4 Then
            BCC = ((BCC Xor CByte(&H30) Xor CByte(&H34)))
            output_str = Chr(2) & "04000" & "0102C1" & "0003" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 5 Then
            BCC = ((BCC Xor CByte(&H30) Xor CByte(&H35)))
            output_str = Chr(2) & "05000" & "0102C1" & "0003" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        End If

        AxMSComm1.Output = output_str

    End Sub

    Public Sub On_Off_thermal(ByVal NodeNumber As Integer, ByVal On_Off As Boolean)

        Dim On_Off_text As String
        Dim BCC As Byte = &H30 Xor &H30 Xor &H30 Xor &H33 Xor &H30 Xor &H30 Xor &H35 Xor &H30 Xor &H31 Xor &H30 Xor &H3

        If On_Off = True Then
            On_Off_text = "0100"
            BCC = BCC Xor CByte(&H30)
        Else
            On_Off_text = "0101"
            BCC = BCC Xor CByte(&H31)
        End If

        If NodeNumber = -1 Then
            'whats this for
            'BCC = BCC Xor CByte(&H58) Xor CByte(&H58)
            'output_str = Chr(2) & "XX000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 1 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H31)
            output_str = Chr(2) & "01000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 2 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H32)
            output_str = Chr(2) & "02000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 3 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H33)
            output_str = Chr(2) & "03000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 4 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H34)
            output_str = Chr(2) & "04000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 5 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H35)
            output_str = Chr(2) & "05000" & "3005" & On_Off_text & Chr(3) & Chr(BCC)
        End If

        AxMSComm1.Output = output_str

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim standby As Boolean = False

        If Button2.Text = "Standby Mode" Then
            Button2.Text = "Operating Mode"
            standby = True
        Else
            Button2.Text = "Standby Mode"
            standby = False
        End If

        If CB_NeedleHeater.Checked Then
            On_Off_thermal(1, standby)
        End If
        If CB_SyringeHeater.Checked Then
            On_Off_thermal(2, standby)
        End If
        If CB_PreThermal.Checked Then
            On_Off_thermal(3, standby)
        End If
        If CB_InThermal.Checked Then
            On_Off_thermal(4, standby)
        End If
        If CB_PostThermal.Checked Then
            On_Off_thermal(5, standby)
        End If
    End Sub


    Public Sub Set_Alarm(ByVal NodeNumber As Integer, ByVal Alarm_Value As Byte)
        Dim CommandText As String
        Dim CommandValue As String

        CommandValue = "000000"

        'kr nov24 change XOR &38 which is 8 in the C1 0008 to whatever mode you want.

        Dim BCC As Byte = &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H32 Xor &H43 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H38 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H3
        BCC = BCC Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) 'commandvalue "000000"
        If Convert.ToByte(Fix(Alarm_Value / 16)) = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If Alarm_Value Mod 16 = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Alarm_Value Mod 16 = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Alarm_Value Mod 16 = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Alarm_Value Mod 16 = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Alarm_Value Mod 16 = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Alarm_Value Mod 16 = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Alarm_Value Mod 16 = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Alarm_Value Mod 16 = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Alarm_Value Mod 16 = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Alarm_Value Mod 16 = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Alarm_Value Mod 16 = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Alarm_Value Mod 16 = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Alarm_Value Mod 16 = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Alarm_Value Mod 16 = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Alarm_Value Mod 16 = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Alarm_Value Mod 16 = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If NodeNumber = 1 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H31)
            CommandText = Chr(2) & "01000" & "0102C1" & "0008" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 2 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H32)
            CommandText = Chr(2) & "02000" & "0102C1" & "0008" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 3 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H33)
            CommandText = Chr(2) & "03000" & "0102C1" & "0008" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 4 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H34)
            CommandText = Chr(2) & "04000" & "0102C1" & "0008" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 5 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H35)
            CommandText = Chr(2) & "05000" & "0102C1" & "0008" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        End If

        'krr

        AxMSComm1.Output = CommandText

        Sleep(100)                       '*************************************** To set for lower limit

        CommandValue = "000000"

        BCC = &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H30 Xor &H32 Xor &H43 Xor &H31 Xor &H30 Xor &H30 Xor &H30 Xor &H39 Xor _
                            &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H30 Xor &H31 Xor &H3
        BCC = BCC Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) Xor CByte(&H30) 'commandvalue "000000"
        If Convert.ToByte(Fix(Alarm_Value / 16)) = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Convert.ToByte(Fix(Alarm_Value / 16)) = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If Alarm_Value Mod 16 = 0 Then
            BCC = BCC Xor CByte(&H30) : CommandValue = CommandValue & "0"
        ElseIf Alarm_Value Mod 16 = 1 Then
            BCC = BCC Xor CByte(&H31) : CommandValue = CommandValue & "1"
        ElseIf Alarm_Value Mod 16 = 2 Then
            BCC = BCC Xor CByte(&H32) : CommandValue = CommandValue & "2"
        ElseIf Alarm_Value Mod 16 = 3 Then
            BCC = BCC Xor CByte(&H33) : CommandValue = CommandValue & "3"
        ElseIf Alarm_Value Mod 16 = 4 Then
            BCC = BCC Xor CByte(&H34) : CommandValue = CommandValue & "4"
        ElseIf Alarm_Value Mod 16 = 5 Then
            BCC = BCC Xor CByte(&H35) : CommandValue = CommandValue & "5"
        ElseIf Alarm_Value Mod 16 = 6 Then
            BCC = BCC Xor CByte(&H36) : CommandValue = CommandValue & "6"
        ElseIf Alarm_Value Mod 16 = 7 Then
            BCC = BCC Xor CByte(&H37) : CommandValue = CommandValue & "7"
        ElseIf Alarm_Value Mod 16 = 8 Then
            BCC = BCC Xor CByte(&H38) : CommandValue = CommandValue & "8"
        ElseIf Alarm_Value Mod 16 = 9 Then
            BCC = BCC Xor CByte(&H39) : CommandValue = CommandValue & "9"
        ElseIf Alarm_Value Mod 16 = 10 Then
            BCC = BCC Xor CByte(&H41) : CommandValue = CommandValue & "A"
        ElseIf Alarm_Value Mod 16 = 11 Then
            BCC = BCC Xor CByte(&H42) : CommandValue = CommandValue & "B"
        ElseIf Alarm_Value Mod 16 = 12 Then
            BCC = BCC Xor CByte(&H43) : CommandValue = CommandValue & "C"
        ElseIf Alarm_Value Mod 16 = 13 Then
            BCC = BCC Xor CByte(&H44) : CommandValue = CommandValue & "D"
        ElseIf Alarm_Value Mod 16 = 14 Then
            BCC = BCC Xor CByte(&H45) : CommandValue = CommandValue & "E"
        ElseIf Alarm_Value Mod 16 = 15 Then
            BCC = BCC Xor CByte(&H46) : CommandValue = CommandValue & "F"
        End If

        If NodeNumber = 1 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H31)
            CommandText = Chr(2) & "01000" & "0102C1" & "0009" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 2 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H32)
            CommandText = Chr(2) & "02000" & "0102C1" & "0009" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 3 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H33)
            CommandText = Chr(2) & "03000" & "0102C1" & "0009" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 4 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H34)
            CommandText = Chr(2) & "04000" & "0102C1" & "0009" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        ElseIf NodeNumber = 5 Then
            BCC = BCC Xor CByte(&H30) Xor CByte(&H35)
            CommandText = Chr(2) & "05000" & "0102C1" & "0009" & "000001" & CommandValue & Chr(3) & Chr(BCC)
        End If

        AxMSComm1.Output = CommandText

        Sleep(100)
    End Sub

End Class
