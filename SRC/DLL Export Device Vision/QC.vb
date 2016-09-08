Imports System.IO
Imports System.Math
Public Class QC
    Inherits System.Windows.Forms.Form

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Xue Wen                                                                                     '
    '   Another method, Not to fire "ValueChanged" event while constructing Form's components.      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Initializing As Boolean = True

#Region " Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Initializing = True
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ValueBinarized As System.Windows.Forms.NumericUpDown
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ValueClose As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ValueOpen As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ValueRoughness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ValueCompactness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents RadioButton_BlackDot As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_WhiteDot As System.Windows.Forms.RadioButton
    Friend WithEvents ValueMinArea As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueMaxArea As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Ok As System.Windows.Forms.Button
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents ValueDotBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Button_Reset As System.Windows.Forms.Button
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents ValueTolerance As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents binarizedimage As System.Windows.Forms.PictureBox
    Friend WithEvents closeoperationimage As System.Windows.Forms.PictureBox
    Friend WithEvents openoperationimage As System.Windows.Forms.PictureBox
    Friend WithEvents originalimage As System.Windows.Forms.PictureBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents ValueDetectedDiameter As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents ValueDesiredDiameter As System.Windows.Forms.NumericUpDown
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbLocal As System.Windows.Forms.RadioButton
    Friend WithEvents rbGlobal As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ValueBinarized = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button_Test = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ValueMinArea = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.ValueMaxArea = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.ValueClose = New System.Windows.Forms.NumericUpDown
        Me.Label4 = New System.Windows.Forms.Label
        Me.ValueOpen = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.ValueRoughness = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.ValueCompactness = New System.Windows.Forms.NumericUpDown
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.ValueDotBrightness = New System.Windows.Forms.NumericUpDown
        Me.Label16 = New System.Windows.Forms.Label
        Me.RadioButton_BlackDot = New System.Windows.Forms.RadioButton
        Me.RadioButton_WhiteDot = New System.Windows.Forms.RadioButton
        Me.Button_Ok = New System.Windows.Forms.Button
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rbGlobal = New System.Windows.Forms.RadioButton
        Me.rbLocal = New System.Windows.Forms.RadioButton
        Me.Label12 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button_Reset = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.binarizedimage = New System.Windows.Forms.PictureBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.ValueTolerance = New System.Windows.Forms.NumericUpDown
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.closeoperationimage = New System.Windows.Forms.PictureBox
        Me.openoperationimage = New System.Windows.Forms.PictureBox
        Me.originalimage = New System.Windows.Forms.PictureBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.ValueDetectedDiameter = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.ValueDesiredDiameter = New System.Windows.Forms.NumericUpDown
        CType(Me.ValueBinarized, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueMinArea, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueMaxArea, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueClose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueOpen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueRoughness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueCompactness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueDotBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.ValueTolerance, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueDesiredDiameter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ValueBinarized
        '
        Me.ValueBinarized.Location = New System.Drawing.Point(240, 112)
        Me.ValueBinarized.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBinarized.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueBinarized.Name = "ValueBinarized"
        Me.ValueBinarized.Size = New System.Drawing.Size(64, 27)
        Me.ValueBinarized.TabIndex = 0
        Me.ValueBinarized.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Binarized Threshold"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button_Test
        '
        Me.Button_Test.Location = New System.Drawing.Point(1112, 160)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(152, 136)
        Me.Button_Test.TabIndex = 2
        Me.Button_Test.Text = "Test"
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'ValueMinArea
        '
        Me.ValueMinArea.DecimalPlaces = 3
        Me.ValueMinArea.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.ValueMinArea.Location = New System.Drawing.Point(440, 112)
        Me.ValueMinArea.Maximum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.ValueMinArea.Minimum = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.ValueMinArea.Name = "ValueMinArea"
        Me.ValueMinArea.Size = New System.Drawing.Size(64, 27)
        Me.ValueMinArea.TabIndex = 4
        Me.ValueMinArea.Value = New Decimal(New Integer() {1, 0, 0, 65536})
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(320, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Min Radius"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueMaxArea
        '
        Me.ValueMaxArea.DecimalPlaces = 3
        Me.ValueMaxArea.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.ValueMaxArea.Location = New System.Drawing.Point(440, 80)
        Me.ValueMaxArea.Maximum = New Decimal(New Integer() {8, 0, 0, 0})
        Me.ValueMaxArea.Minimum = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.ValueMaxArea.Name = "ValueMaxArea"
        Me.ValueMaxArea.Size = New System.Drawing.Size(64, 27)
        Me.ValueMaxArea.TabIndex = 6
        Me.ValueMaxArea.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(320, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 23)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Max Radius"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueClose
        '
        Me.ValueClose.Location = New System.Drawing.Point(240, 144)
        Me.ValueClose.Maximum = New Decimal(New Integer() {9, 0, 0, 0})
        Me.ValueClose.Name = "ValueClose"
        Me.ValueClose.Size = New System.Drawing.Size(64, 27)
        Me.ValueClose.TabIndex = 10
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(232, 23)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Number of Close Operations"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueOpen
        '
        Me.ValueOpen.Location = New System.Drawing.Point(240, 176)
        Me.ValueOpen.Maximum = New Decimal(New Integer() {9, 0, 0, 0})
        Me.ValueOpen.Name = "ValueOpen"
        Me.ValueOpen.Size = New System.Drawing.Size(64, 27)
        Me.ValueOpen.TabIndex = 8
        Me.ValueOpen.Value = New Decimal(New Integer() {4, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 176)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(232, 23)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Number of Open Operations"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueRoughness
        '
        Me.ValueRoughness.DecimalPlaces = 1
        Me.ValueRoughness.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueRoughness.Location = New System.Drawing.Point(440, 144)
        Me.ValueRoughness.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.ValueRoughness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueRoughness.Name = "ValueRoughness"
        Me.ValueRoughness.Size = New System.Drawing.Size(64, 27)
        Me.ValueRoughness.TabIndex = 13
        Me.ValueRoughness.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(320, 144)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 23)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Roughness"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueCompactness
        '
        Me.ValueCompactness.DecimalPlaces = 1
        Me.ValueCompactness.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueCompactness.Location = New System.Drawing.Point(440, 176)
        Me.ValueCompactness.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.ValueCompactness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueCompactness.Name = "ValueCompactness"
        Me.ValueCompactness.Size = New System.Drawing.Size(64, 27)
        Me.ValueCompactness.TabIndex = 15
        Me.ValueCompactness.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(320, 176)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 23)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Compactness"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(496, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 23)
        Me.Label11.TabIndex = 47
        Me.Label11.Text = "mm"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(496, 80)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 23)
        Me.Label10.TabIndex = 46
        Me.Label10.Text = "mm "
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ValueDotBrightness
        '
        Me.ValueDotBrightness.Location = New System.Drawing.Point(240, 80)
        Me.ValueDotBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueDotBrightness.Name = "ValueDotBrightness"
        Me.ValueDotBrightness.Size = New System.Drawing.Size(64, 27)
        Me.ValueDotBrightness.TabIndex = 41
        Me.ValueDotBrightness.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(16, 80)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(96, 23)
        Me.Label16.TabIndex = 40
        Me.Label16.Text = "Brightness"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RadioButton_BlackDot
        '
        Me.RadioButton_BlackDot.Checked = True
        Me.RadioButton_BlackDot.Location = New System.Drawing.Point(24, 208)
        Me.RadioButton_BlackDot.Name = "RadioButton_BlackDot"
        Me.RadioButton_BlackDot.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton_BlackDot.TabIndex = 18
        Me.RadioButton_BlackDot.TabStop = True
        Me.RadioButton_BlackDot.Text = "Black Dot"
        '
        'RadioButton_WhiteDot
        '
        Me.RadioButton_WhiteDot.Location = New System.Drawing.Point(104, 208)
        Me.RadioButton_WhiteDot.Name = "RadioButton_WhiteDot"
        Me.RadioButton_WhiteDot.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton_WhiteDot.TabIndex = 19
        Me.RadioButton_WhiteDot.Text = "White Dot"
        '
        'Button_Ok
        '
        Me.Button_Ok.Location = New System.Drawing.Point(1112, 300)
        Me.Button_Ok.Name = "Button_Ok"
        Me.Button_Ok.Size = New System.Drawing.Size(75, 24)
        Me.Button_Ok.TabIndex = 19
        Me.Button_Ok.Text = "Ok"
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Location = New System.Drawing.Point(1192, 300)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(72, 24)
        Me.Button_Cancel.TabIndex = 20
        Me.Button_Cancel.Text = "Cancel"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(488, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 23)
        Me.Label8.TabIndex = 66
        Me.Label8.Text = "mm"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(880, 376)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(152, 24)
        Me.Label9.TabIndex = 67
        Me.Label9.Text = "Dot Size Diameter"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.ValueBinarized)
        Me.Panel2.Controls.Add(Me.ValueClose)
        Me.Panel2.Controls.Add(Me.ValueDotBrightness)
        Me.Panel2.Controls.Add(Me.ValueOpen)
        Me.Panel2.Controls.Add(Me.ValueMaxArea)
        Me.Panel2.Controls.Add(Me.RadioButton_WhiteDot)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.ValueCompactness)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.ValueRoughness)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.ValueMinArea)
        Me.Panel2.Controls.Add(Me.RadioButton_BlackDot)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Location = New System.Drawing.Point(0, 40)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(552, 288)
        Me.Panel2.TabIndex = 25
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbGlobal)
        Me.GroupBox1.Controls.Add(Me.rbLocal)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 16)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(520, 48)
        Me.GroupBox1.TabIndex = 49
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "QC Mode"
        '
        'rbGlobal
        '
        Me.rbGlobal.Location = New System.Drawing.Point(160, 16)
        Me.rbGlobal.Name = "rbGlobal"
        Me.rbGlobal.TabIndex = 1
        Me.rbGlobal.Text = "Global"
        '
        'rbLocal
        '
        Me.rbLocal.Location = New System.Drawing.Point(24, 16)
        Me.rbLocal.Name = "rbLocal"
        Me.rbLocal.TabIndex = 0
        Me.rbLocal.Text = "Local"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(0, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 24)
        Me.Label12.TabIndex = 48
        Me.Label12.Text = "Filter Settings"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button_Reset)
        Me.GroupBox2.Controls.Add(Me.Button_Test)
        Me.GroupBox2.Controls.Add(Me.Button_Ok)
        Me.GroupBox2.Controls.Add(Me.Button_Cancel)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Panel2)
        Me.GroupBox2.Controls.Add(Me.Panel3)
        Me.GroupBox2.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1285, 400)
        Me.GroupBox2.TabIndex = 62
        Me.GroupBox2.TabStop = False
        '
        'Button_Reset
        '
        Me.Button_Reset.Location = New System.Drawing.Point(1112, 40)
        Me.Button_Reset.Name = "Button_Reset"
        Me.Button_Reset.Size = New System.Drawing.Size(152, 120)
        Me.Button_Reset.TabIndex = 62
        Me.Button_Reset.Text = "Reset"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(8, 288)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(536, 27)
        Me.TextBox1.TabIndex = 18
        Me.TextBox1.Text = ""
        Me.TextBox1.Visible = False
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.binarizedimage)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.Label13)
        Me.Panel3.Controls.Add(Me.Label14)
        Me.Panel3.Controls.Add(Me.ValueTolerance)
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.Label17)
        Me.Panel3.Controls.Add(Me.closeoperationimage)
        Me.Panel3.Controls.Add(Me.openoperationimage)
        Me.Panel3.Controls.Add(Me.originalimage)
        Me.Panel3.Controls.Add(Me.Label18)
        Me.Panel3.Controls.Add(Me.Label19)
        Me.Panel3.Controls.Add(Me.Label21)
        Me.Panel3.Controls.Add(Me.Label20)
        Me.Panel3.Controls.Add(Me.ValueDetectedDiameter)
        Me.Panel3.Controls.Add(Me.Label22)
        Me.Panel3.Controls.Add(Me.Label23)
        Me.Panel3.Controls.Add(Me.ValueDesiredDiameter)
        Me.Panel3.Location = New System.Drawing.Point(552, 40)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(560, 288)
        Me.Panel3.TabIndex = 69
        '
        'binarizedimage
        '
        Me.binarizedimage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.binarizedimage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.binarizedimage.Location = New System.Drawing.Point(140, 136)
        Me.binarizedimage.Name = "binarizedimage"
        Me.binarizedimage.Size = New System.Drawing.Size(140, 150)
        Me.binarizedimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.binarizedimage.TabIndex = 69
        Me.binarizedimage.TabStop = False
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(0, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 24)
        Me.Label13.TabIndex = 48
        Me.Label13.Text = "Results"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(120, 23)
        Me.Label14.TabIndex = 7
        Me.Label14.Text = "Tolerance   +/-"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueTolerance
        '
        Me.ValueTolerance.DecimalPlaces = 2
        Me.ValueTolerance.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueTolerance.Location = New System.Drawing.Point(152, 72)
        Me.ValueTolerance.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.ValueTolerance.Name = "ValueTolerance"
        Me.ValueTolerance.Size = New System.Drawing.Size(64, 27)
        Me.ValueTolerance.TabIndex = 6
        Me.ValueTolerance.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(208, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(56, 23)
        Me.Label15.TabIndex = 46
        Me.Label15.Text = "mm "
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(272, 32)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(152, 23)
        Me.Label17.TabIndex = 7
        Me.Label17.Text = "Detected Diameter"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'closeoperationimage
        '
        Me.closeoperationimage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.closeoperationimage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.closeoperationimage.Location = New System.Drawing.Point(280, 136)
        Me.closeoperationimage.Name = "closeoperationimage"
        Me.closeoperationimage.Size = New System.Drawing.Size(140, 150)
        Me.closeoperationimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.closeoperationimage.TabIndex = 69
        Me.closeoperationimage.TabStop = False
        '
        'openoperationimage
        '
        Me.openoperationimage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.openoperationimage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.openoperationimage.Location = New System.Drawing.Point(420, 136)
        Me.openoperationimage.Name = "openoperationimage"
        Me.openoperationimage.Size = New System.Drawing.Size(140, 150)
        Me.openoperationimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.openoperationimage.TabIndex = 69
        Me.openoperationimage.TabStop = False
        '
        'originalimage
        '
        Me.originalimage.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.originalimage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.originalimage.Location = New System.Drawing.Point(0, 136)
        Me.originalimage.Name = "originalimage"
        Me.originalimage.Size = New System.Drawing.Size(140, 150)
        Me.originalimage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.originalimage.TabIndex = 69
        Me.originalimage.TabStop = False
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(0, 112)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(144, 23)
        Me.Label18.TabIndex = 46
        Me.Label18.Text = "Original"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(144, 112)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(136, 23)
        Me.Label19.TabIndex = 46
        Me.Label19.Text = "After Binarizing"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(280, 112)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(144, 23)
        Me.Label21.TabIndex = 46
        Me.Label21.Text = "After Closing"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(424, 112)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(136, 23)
        Me.Label20.TabIndex = 46
        Me.Label20.Text = "After Opening"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ValueDetectedDiameter
        '
        Me.ValueDetectedDiameter.Location = New System.Drawing.Point(424, 32)
        Me.ValueDetectedDiameter.Name = "ValueDetectedDiameter"
        Me.ValueDetectedDiameter.ReadOnly = True
        Me.ValueDetectedDiameter.Size = New System.Drawing.Size(64, 27)
        Me.ValueDetectedDiameter.TabIndex = 68
        Me.ValueDetectedDiameter.Text = "0"
        Me.ValueDetectedDiameter.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(216, 32)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(40, 23)
        Me.Label22.TabIndex = 66
        Me.Label22.Text = "mm"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(8, 32)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(144, 23)
        Me.Label23.TabIndex = 7
        Me.Label23.Text = "Desired Diameter"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueDesiredDiameter
        '
        Me.ValueDesiredDiameter.DecimalPlaces = 2
        Me.ValueDesiredDiameter.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueDesiredDiameter.Location = New System.Drawing.Point(152, 32)
        Me.ValueDesiredDiameter.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.ValueDesiredDiameter.Name = "ValueDesiredDiameter"
        Me.ValueDesiredDiameter.Size = New System.Drawing.Size(64, 27)
        Me.ValueDesiredDiameter.TabIndex = 6
        Me.ValueDesiredDiameter.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'QC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(1284, 328)
        Me.Controls.Add(Me.GroupBox2)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Location = New System.Drawing.Point(0, 33)
        Me.Name = "QC"
        Me.Text = "QC"
        Me.TopMost = True
        CType(Me.ValueBinarized, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueMinArea, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueMaxArea, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueClose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueOpen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueRoughness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueCompactness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueDotBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.ValueTolerance, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueDesiredDiameter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'developer tips
    '"1) Set Roughness to 1 if perfect circle is choosen. " 
    '"2) Set Compactness to 1 if solid circle is choosen "
    '"3) Adjust minimum size of the dot" 
    '"4) Adjust Binarized Threshold/Open/Close to remove the noise at the dot." 
    '"5) Adjust brightness to contrast the dot with background if the dot is blur"
    Public Delegate Sub FormCloseDelegate()
    Public FormCloseEvent As FormCloseDelegate = Nothing
    Private Shared m_instance As QC
    Private m_isGlobalQC As Boolean = False
    Public Shared ReadOnly Property Instance() As QC
        Get
            If m_instance Is Nothing Then
                m_instance = New QC
            Else
                If m_instance.IsDisposed Then
                    m_instance = New QC
                End If
            End If
            Return m_instance
        End Get
    End Property

    Public Structure QCParam
        Public _BlackDot As Boolean
        Public _Binarized As Integer
        Public _MaxArea As Double
        Public _MinArea As Double
        Public _Close As Integer
        Public _Open As Integer
        Public _Roughness As Double
        Public _Compactness As Double
        Public _MRegionX As Decimal
        Public _MRegionY As Decimal
        Public _MROIx As Decimal
        Public _MROIy As Decimal
        Public _Brightness As Integer

        'kr added
        Public _Diameter As Double
        Public _Tolerance As Double
    End Structure

    Public QC_success As String 'for spc logging
    Public QC_failure As String 'for dot size popup check

#Region "Global Variables"
    Dim timer As Boolean = False
    Shared PointX1a, PointY1a, PointX2a, PointY2a, PointX3a, PointY3a, PointX4a, PointY4a, PointX5a, PointY5a As Double
    Dim P1x, P1y, P2x, P2y, P3x, P3y, P4x, P4y, P5x, P5y As Double
    Dim MQC_X, MQC_Y, MQC_Dia As Double
    Dim MQCRegionX As Double = 768 / 2
    Dim MQCRegionY As Double = 576 / 2
    Dim MQCRoix As Double = 100
    Dim MQCRoiY As Double = 100
    Dim BlackDot As Boolean = True
    Dim Status As Integer = 0   '0=yet done, 1=done, 2=cancel '3=reset
    Dim MainEdge As Integer = 1
#End Region

    Function GetQCStatus() As Integer
        TraceDoEvents()
        Return Status
    End Function
    Function GetIsGlobalQC() As Boolean
        Return Me.m_isGlobalQC
    End Function
    Function SetAllowGlobalMode(ByVal allow As Boolean)
        If allow Then
            Me.rbGlobal.Enabled = True
            Me.rbGlobal.Checked = False
            Me.rbLocal.Checked = True
            Me.rbLocal.Enabled = True
        Else
            Me.rbGlobal.Enabled = False
            Me.rbGlobal.Checked = False
            Me.rbLocal.Checked = True
            Me.rbLocal.Enabled = True
        End If
    End Function
    Function SetGlobalMode(ByVal isGlobal As Boolean)
        If isGlobal Then
            Me.rbLocal.Checked = False
            Me.rbLocal.Enabled = False
            Me.rbGlobal.Enabled = True
            Me.rbGlobal.Checked = True
        Else
            Me.rbLocal.Checked = True
            Me.rbLocal.Enabled = True
            Me.rbGlobal.Enabled = False
            Me.rbGlobal.Checked = False
        End If
       
    End Function
    Sub SetQCBrightness(ByVal brightness As Decimal)
        ValueDotBrightness.Value = brightness
    End Sub
    Sub ResetQCStatus()
        Status = 0
    End Sub
    Function GetQCParameters(ByRef QCParam As QCParam)
        QCParam._Brightness = ValueDotBrightness.Value
        QCParam._BlackDot = RadioButton_BlackDot.Checked
        QCParam._Binarized = ValueBinarized.Value
        QCParam._MaxArea = ValueMaxArea.Value
        QCParam._MinArea = ValueMinArea.Value
        QCParam._Close = ValueClose.Value
        QCParam._Open = ValueOpen.Value
        QCParam._Roughness = ValueRoughness.Value
        QCParam._Compactness = ValueCompactness.Value
        QCParam._MRegionX = MQCRegionX
        QCParam._MRegionY = MQCRegionY
        QCParam._MROIx = MQCRoix
        QCParam._MROIy = MQCRoiY
        QCParam._Diameter = CDbl(ValueDesiredDiameter.Value)
        QCParam._Tolerance = ValueTolerance.Value
    End Function
#Region "Production"
    Function ChipEdgeOutput(ByVal PointX1, ByVal PointY1, ByVal PointX2, ByVal PointY2, ByVal PointX3, ByVal PointY3, ByVal PointX4, ByVal PointY4, ByVal MEdge)
        PointX1a = PointX1
        PointX2a = PointX2
        PointX3a = PointX3
        PointX4a = PointX4
        PointY1a = PointY1
        PointY2a = PointY2
        PointY3a = PointY3
        PointY4a = PointY4
        MainEdge = MEdge
    End Function
    Function DotOutput(ByVal QC_X As Double, ByVal QC_Y As Double, ByVal QC_Dia As Double) 'mm
        MQC_X = QC_X
        MQC_Y = QC_Y
        MQC_Dia = QC_Dia
    End Function
    Dim QC As Boolean = False
    Sub QCResults(ByVal Results As Boolean)
        If Results = True Then
            QC = True
        Else
            QC = False
        End If
    End Sub
#End Region
    Sub GetVariables(ByVal DisplayCenterXPosition As Double, ByVal DisplayCenterYPosition As Double, ByVal MRoix As Double, ByVal MRoiy As Double, ByVal QC_X As Double, ByVal QC_Y As Double, ByVal QC_Dia As Double)
        MQCRegionX = DisplayCenterXPosition
        MQCRegionY = DisplayCenterYPosition
        MQCRoix = MRoix
        MQCRoiY = MRoiy
        MQC_X = QC_X
        MQC_Y = QC_Y
        MQC_Dia = QC_Dia
    End Sub
    Function ChipEdgePointsCoordinate(ByVal PointX1, ByVal PointY1, ByVal PointX2, ByVal PointY2, ByVal PointX3, ByVal PointY3, ByVal PointX4, ByVal PointY4, ByVal PointX5, ByVal PointY5)
        P1x = PointX1
        P2x = PointX2
        P3x = PointX3
        P4x = PointX4
        P5x = PointX5
        P1y = PointY1
        P2y = PointY2
        P3y = PointY3
        P4y = PointY4
        P5y = PointY5
    End Function
    Function CheckFileName(ByVal FileName) As Boolean
        'This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
        'It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
        'It keeps going until it returns a filename that does not exist.
        'You could just create a filename from the ID field but that means writing the record - and it still might exist!
        'N.B. Requires strSaveToPath variable to be available - and containing the path to save to
        Dim Counter
        Dim Flag
        Dim strTempFileName
        Dim FileExt As String
        Dim NewFullPath
        Dim objFSO As Object

        objFSO = CreateObject("Scripting.FileSystemObject")
        Counter = 0
        FileExt = "txt" 'GetExt(FileName)
        strTempFileName = "Reject" 'StripExt(FileName)
        NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\QC" & "\" & FileName & "." & FileExt
        Flag = False

        If objFSO.FileExists(NewFullPath) = True Then
            Dim FileName_checked = Mid(NewFullPath, InStrRev(NewFullPath, "\") + 1)
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub QC_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True
        FrmVision.modelregionDrawing()
        Status = 2
    End Sub
    Private Sub Button_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Test.Click
        If timer = False Then
            timer = True
            Timer1.Start()
            Button_Test.Text = "Stop"
        Else
            Timer1.Stop()
            timer = False
            Button_Test.Text = "Test"
            FrmVision.DisplayIndicator()
            FrmVision.modelregionDrawing()
        End If
    End Sub
    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        Status = 1  'for SJ to check if QC done
        DotResetVariables()
        Dim QCParam As QCParam
        GetQCParameters(QCParam)
        Me.Visible = False
        Me.m_isGlobalQC = Me.rbGlobal.Checked
        FrmVision.DisplayIndicator()
        If Not (FormCloseEvent Is Nothing) Then
            FormCloseEvent()
        End If
    End Sub
    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Status = 2 'for SJ to check if QC canceled
        DotResetVariables()
        Me.Visible = False
        FrmVision.DisplayIndicator()
        If Not (FormCloseEvent Is Nothing) Then
            FormCloseEvent()
        End If
    End Sub
    Private Sub RadioButton_WhiteDot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_WhiteDot.CheckedChanged
        If RadioButton_WhiteDot.Checked = True Then
            BlackDot = False
        Else
            BlackDot = True
        End If
    End Sub
    Private Sub RadioButton_BlackDot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_BlackDot.CheckedChanged
        If RadioButton_BlackDot.Checked = True Then
            BlackDot = True
        Else
            BlackDot = False
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Timer1.Stop()
            Me.DialogResult = DialogResult.OK
            FrmVision.ClearDisplay()
            FrmVision.modelregionDrawing()
            FrmVision.DotQC(BlackDot, ValueBinarized.Value, ValueClose.Value, ValueOpen.Value, ValueMinArea.Value, ValueMaxArea.Value, ValueRoughness.Value, ValueCompactness.Value)

            Dim orig As New FileStream("C:\IDS\SRC\DLL Export Device Vision\QC\test1.bmp", FileMode.Open, FileAccess.Read)
            Dim binarized As New FileStream("C:\IDS\SRC\DLL Export Device Vision\QC\test2.bmp", FileMode.Open, FileAccess.Read)
            Dim close As New FileStream("C:\IDS\SRC\DLL Export Device Vision\QC\test3.bmp", FileMode.Open, FileAccess.Read)
            Dim open As New FileStream("C:\IDS\SRC\DLL Export Device Vision\QC\test4.bmp", FileMode.Open, FileAccess.Read)
            originalimage.Image = Image.FromStream(orig)
            binarizedimage.Image = Image.FromStream(binarized)
            closeoperationimage.Image = Image.FromStream(close)
            openoperationimage.Image = Image.FromStream(open)
            binarized.Close()
            orig.Close()
            close.Close()
            open.Close()

            If timer Then
                Timer1.Start()
            Else
                FrmVision.modelregionDrawing()
            End If

        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub


#Region "Reset"
    Sub SetQCReset()
        DotResetVariables()
        FrmVision.DisplayIndicator()
        FrmVision.ResetChipEdgePoint()  'clear 5points
    End Sub
    Sub DotResetVariables()
        'Reset all variables
        ValueDetectedDiameter.Text = 0
        Button_Test.Text = "Test"
        Timer1.Stop()
        timer = False
        FrmVision.Reset()

        originalimage.Image = Nothing
        binarizedimage.Image = Nothing
        closeoperationimage.Image = Nothing
        openoperationimage.Image = Nothing
    End Sub
#End Region

    Private Sub Button_Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Reset.Click

        Panel2.BringToFront()
        Button_Test.Enabled = True
        Button_Ok.Enabled = True

        DotResetVariables()

        FrmVision.ClearDisplay()
        FrmVision.Reset()
        FrmVision.ModelRegionSetting(True)
        FrmVision.modelregionDrawing()

    End Sub

    Private Sub ValueDotBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueDotBrightness.TextChanged
        If ValueDotBrightness.Focused = True Then  'if you remove this an error will occur during initializatio
            FrmVision.SetBrightness(ValueDotBrightness.Value)
        End If
    End Sub

    Private Sub rbGlobal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbGlobal.Click
        rbLocal.Checked = False
    End Sub

    Private Sub rbLocal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbLocal.Click
        rbGlobal.Checked = False
    End Sub
End Class