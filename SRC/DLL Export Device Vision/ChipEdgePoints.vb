Imports System.IO

Public Class ChipEdgePoints
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

        Initializing = False
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
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_Comments As System.Windows.Forms.TextBox
    Friend WithEvents RadioButton_TwoEdges As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_OneEdge As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Dot As System.Windows.Forms.RadioButton
    Friend WithEvents CheckBox_ChipRec_Enable As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_EdgeClearance As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents ValueRot As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValuePos As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueSize As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox_PosY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PosX As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_SizeX As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox_SizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ValueContrast As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents ValueBrightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueThreshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents RadioButton_Outside_In As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Inside_out As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_RRot As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RPosY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RPosX As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RSizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Score As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RSizeX As System.Windows.Forms.TextBox
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button_Ok As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents ValueROI As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Horizontal As System.Windows.Forms.RadioButton
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_CCW As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_CW As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox_Vertical_Horizontal As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Vertical As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox_DispenseModel As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Reset As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button_Cancel = New System.Windows.Forms.Button
        Me.Button_Ok = New System.Windows.Forms.Button
        Me.TextBox_Comments = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.ValueRot = New System.Windows.Forms.NumericUpDown
        Me.ValuePos = New System.Windows.Forms.NumericUpDown
        Me.ValueSize = New System.Windows.Forms.NumericUpDown
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.TextBox_PosY = New System.Windows.Forms.TextBox
        Me.TextBox_PosX = New System.Windows.Forms.TextBox
        Me.TextBox_SizeX = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox_SizeY = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.ValueROI = New System.Windows.Forms.NumericUpDown
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.RadioButton_CCW = New System.Windows.Forms.RadioButton
        Me.RadioButton_CW = New System.Windows.Forms.RadioButton
        Me.Label39 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.ValueContrast = New System.Windows.Forms.NumericUpDown
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.ValueBrightness = New System.Windows.Forms.NumericUpDown
        Me.ValueThreshold = New System.Windows.Forms.NumericUpDown
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.RadioButton_Outside_In = New System.Windows.Forms.RadioButton
        Me.RadioButton_Inside_out = New System.Windows.Forms.RadioButton
        Me.GroupBox_Vertical_Horizontal = New System.Windows.Forms.GroupBox
        Me.RadioButton_Horizontal = New System.Windows.Forms.RadioButton
        Me.RadioButton_Vertical = New System.Windows.Forms.RadioButton
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox_EdgeClearance = New System.Windows.Forms.TextBox
        Me.CheckBox_ChipRec_Enable = New System.Windows.Forms.CheckBox
        Me.GroupBox_DispenseModel = New System.Windows.Forms.GroupBox
        Me.RadioButton_TwoEdges = New System.Windows.Forms.RadioButton
        Me.RadioButton_OneEdge = New System.Windows.Forms.RadioButton
        Me.RadioButton_Dot = New System.Windows.Forms.RadioButton
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.TextBox_RRot = New System.Windows.Forms.TextBox
        Me.TextBox_RPosY = New System.Windows.Forms.TextBox
        Me.TextBox_RPosX = New System.Windows.Forms.TextBox
        Me.TextBox_RSizeY = New System.Windows.Forms.TextBox
        Me.Label38 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.Button_Test = New System.Windows.Forms.Button
        Me.Label26 = New System.Windows.Forms.Label
        Me.TextBox_Score = New System.Windows.Forms.TextBox
        Me.TextBox_RSizeX = New System.Windows.Forms.TextBox
        Me.Button_Reset = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.ValueRot, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValuePos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_Settings.SuspendLayout()
        CType(Me.ValueROI, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        CType(Me.ValueContrast, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueThreshold, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox_Vertical_Horizontal.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox_DispenseModel.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Location = New System.Drawing.Point(8, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1264, 320)
        Me.Panel1.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button_Cancel)
        Me.GroupBox1.Controls.Add(Me.Button_Ok)
        Me.GroupBox1.Controls.Add(Me.TextBox_Comments)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.GroupBox_Settings)
        Me.GroupBox1.Controls.Add(Me.GroupBox5)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1248, 312)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ChipEdge"
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Location = New System.Drawing.Point(1080, 280)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.TabIndex = 11
        Me.Button_Cancel.Text = "Cancel"
        '
        'Button_Ok
        '
        Me.Button_Ok.Enabled = False
        Me.Button_Ok.Location = New System.Drawing.Point(1160, 280)
        Me.Button_Ok.Name = "Button_Ok"
        Me.Button_Ok.TabIndex = 10
        Me.Button_Ok.Text = "Ok"
        '
        'TextBox_Comments
        '
        Me.TextBox_Comments.Location = New System.Drawing.Point(8, 280)
        Me.TextBox_Comments.Name = "TextBox_Comments"
        Me.TextBox_Comments.Size = New System.Drawing.Size(1064, 20)
        Me.TextBox_Comments.TabIndex = 9
        Me.TextBox_Comments.Text = "Define settings for Chip Recognition."
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label18)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.Label19)
        Me.GroupBox2.Controls.Add(Me.ValueRot)
        Me.GroupBox2.Controls.Add(Me.ValuePos)
        Me.GroupBox2.Controls.Add(Me.ValueSize)
        Me.GroupBox2.Controls.Add(Me.Label17)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label14)
        Me.GroupBox2.Controls.Add(Me.Label13)
        Me.GroupBox2.Controls.Add(Me.Label12)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.TextBox_PosY)
        Me.GroupBox2.Controls.Add(Me.TextBox_PosX)
        Me.GroupBox2.Controls.Add(Me.TextBox_SizeX)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.TextBox_SizeY)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Button3)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 256)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Chip's Size & Position"
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(128, 136)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(24, 23)
        Me.Label18.TabIndex = 33
        Me.Label18.Text = "±"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(248, 88)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(24, 23)
        Me.Label20.TabIndex = 38
        Me.Label20.Text = "%"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(248, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(24, 23)
        Me.Label19.TabIndex = 37
        Me.Label19.Text = "%"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ValueRot
        '
        Me.ValueRot.DecimalPlaces = 1
        Me.ValueRot.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueRot.Location = New System.Drawing.Point(88, 136)
        Me.ValueRot.Maximum = New Decimal(New Integer() {99, 0, 0, 0})
        Me.ValueRot.Name = "ValueRot"
        Me.ValueRot.Size = New System.Drawing.Size(40, 20)
        Me.ValueRot.TabIndex = 36
        Me.ValueRot.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueRot.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'ValuePos
        '
        Me.ValuePos.DecimalPlaces = 1
        Me.ValuePos.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValuePos.Location = New System.Drawing.Point(204, 88)
        Me.ValuePos.Minimum = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValuePos.Name = "ValuePos"
        Me.ValuePos.Size = New System.Drawing.Size(44, 20)
        Me.ValuePos.TabIndex = 35
        Me.ValuePos.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValuePos.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'ValueSize
        '
        Me.ValueSize.DecimalPlaces = 1
        Me.ValueSize.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueSize.Location = New System.Drawing.Point(204, 32)
        Me.ValueSize.Minimum = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueSize.Name = "ValueSize"
        Me.ValueSize.Size = New System.Drawing.Size(44, 20)
        Me.ValueSize.TabIndex = 34
        Me.ValueSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueSize.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(180, 88)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(24, 23)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "±"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(180, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(24, 23)
        Me.Label16.TabIndex = 31
        Me.Label16.Text = "±"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(148, 104)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 23)
        Me.Label7.TabIndex = 30
        Me.Label7.Text = "mm,"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(148, 80)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 23)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "mm,"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(66, 104)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(16, 24)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "Y"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(66, 48)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(16, 24)
        Me.Label14.TabIndex = 27
        Me.Label14.Text = "Y"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(66, 80)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(16, 24)
        Me.Label13.TabIndex = 26
        Me.Label13.Text = "X"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(66, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(16, 24)
        Me.Label12.TabIndex = 25
        Me.Label12.Text = "X"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 136)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 24)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Rotation:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 80)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(56, 24)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "Position:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(8, 24)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 24)
        Me.Label9.TabIndex = 22
        Me.Label9.Text = "Size:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(148, 136)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 23)
        Me.Label8.TabIndex = 21
        Me.Label8.Text = "deg,"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_PosY
        '
        Me.TextBox_PosY.Location = New System.Drawing.Point(84, 104)
        Me.TextBox_PosY.Name = "TextBox_PosY"
        Me.TextBox_PosY.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_PosY.TabIndex = 18
        Me.TextBox_PosY.Text = "0"
        Me.TextBox_PosY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_PosX
        '
        Me.TextBox_PosX.Location = New System.Drawing.Point(84, 80)
        Me.TextBox_PosX.Name = "TextBox_PosX"
        Me.TextBox_PosX.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_PosX.TabIndex = 16
        Me.TextBox_PosX.Text = "0"
        Me.TextBox_PosX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_SizeX
        '
        Me.TextBox_SizeX.Location = New System.Drawing.Point(84, 24)
        Me.TextBox_SizeX.Name = "TextBox_SizeX"
        Me.TextBox_SizeX.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_SizeX.TabIndex = 14
        Me.TextBox_SizeX.Text = "0"
        Me.TextBox_SizeX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(148, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 23)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "mm,"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_SizeY
        '
        Me.TextBox_SizeY.Location = New System.Drawing.Point(84, 48)
        Me.TextBox_SizeY.Name = "TextBox_SizeY"
        Me.TextBox_SizeY.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_SizeY.TabIndex = 12
        Me.TextBox_SizeY.Text = "0"
        Me.TextBox_SizeY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(148, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 23)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "mm,"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(208, 136)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(56, 23)
        Me.Button3.TabIndex = 42
        Me.Button3.Text = "HideROI"
        '
        'GroupBox_Settings
        '
        Me.GroupBox_Settings.Controls.Add(Me.ValueROI)
        Me.GroupBox_Settings.Controls.Add(Me.GroupBox7)
        Me.GroupBox_Settings.Controls.Add(Me.Label39)
        Me.GroupBox_Settings.Controls.Add(Me.Label21)
        Me.GroupBox_Settings.Controls.Add(Me.ValueContrast)
        Me.GroupBox_Settings.Controls.Add(Me.Label25)
        Me.GroupBox_Settings.Controls.Add(Me.Label24)
        Me.GroupBox_Settings.Controls.Add(Me.Label23)
        Me.GroupBox_Settings.Controls.Add(Me.ValueBrightness)
        Me.GroupBox_Settings.Controls.Add(Me.ValueThreshold)
        Me.GroupBox_Settings.Controls.Add(Me.GroupBox3)
        Me.GroupBox_Settings.Controls.Add(Me.GroupBox_Vertical_Horizontal)
        Me.GroupBox_Settings.Controls.Add(Me.GroupBox4)
        Me.GroupBox_Settings.Controls.Add(Me.GroupBox_DispenseModel)
        Me.GroupBox_Settings.Location = New System.Drawing.Point(296, 16)
        Me.GroupBox_Settings.Name = "GroupBox_Settings"
        Me.GroupBox_Settings.Size = New System.Drawing.Size(504, 256)
        Me.GroupBox_Settings.TabIndex = 2
        Me.GroupBox_Settings.TabStop = False
        Me.GroupBox_Settings.Text = "Settings"
        '
        'ValueROI
        '
        Me.ValueROI.DecimalPlaces = 1
        Me.ValueROI.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueROI.Location = New System.Drawing.Point(224, 96)
        Me.ValueROI.Maximum = New Decimal(New Integer() {3, 0, 0, 0})
        Me.ValueROI.Minimum = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.ValueROI.Name = "ValueROI"
        Me.ValueROI.Size = New System.Drawing.Size(48, 20)
        Me.ValueROI.TabIndex = 44
        Me.ValueROI.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueROI.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.RadioButton_CCW)
        Me.GroupBox7.Controls.Add(Me.RadioButton_CW)
        Me.GroupBox7.Location = New System.Drawing.Point(344, 72)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(128, 48)
        Me.GroupBox7.TabIndex = 48
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Direction:"
        '
        'RadioButton_CCW
        '
        Me.RadioButton_CCW.Location = New System.Drawing.Point(64, 16)
        Me.RadioButton_CCW.Name = "RadioButton_CCW"
        Me.RadioButton_CCW.Size = New System.Drawing.Size(56, 24)
        Me.RadioButton_CCW.TabIndex = 67
        Me.RadioButton_CCW.Text = "CCW"
        '
        'RadioButton_CW
        '
        Me.RadioButton_CW.Checked = True
        Me.RadioButton_CW.Location = New System.Drawing.Point(8, 16)
        Me.RadioButton_CW.Name = "RadioButton_CW"
        Me.RadioButton_CW.TabIndex = 66
        Me.RadioButton_CW.TabStop = True
        Me.RadioButton_CW.Text = "CW"
        '
        'Label39
        '
        Me.Label39.Location = New System.Drawing.Point(272, 96)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(32, 23)
        Me.Label39.TabIndex = 45
        Me.Label39.Text = "mm"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(216, 72)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(64, 23)
        Me.Label21.TabIndex = 43
        Me.Label21.Text = "ROI size:"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ValueContrast
        '
        Me.ValueContrast.Location = New System.Drawing.Point(144, 96)
        Me.ValueContrast.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueContrast.Name = "ValueContrast"
        Me.ValueContrast.ReadOnly = True
        Me.ValueContrast.Size = New System.Drawing.Size(48, 20)
        Me.ValueContrast.TabIndex = 41
        Me.ValueContrast.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueContrast.Value = New Decimal(New Integer() {55, 0, 0, 0})
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(136, 72)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(64, 23)
        Me.Label25.TabIndex = 40
        Me.Label25.Text = "Contrast:"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(72, 72)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(64, 23)
        Me.Label24.TabIndex = 39
        Me.Label24.Text = "Brightness:"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(8, 72)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(64, 23)
        Me.Label23.TabIndex = 38
        Me.Label23.Text = "Threshold:"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ValueBrightness
        '
        Me.ValueBrightness.Location = New System.Drawing.Point(80, 96)
        Me.ValueBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.ValueBrightness.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueBrightness.Name = "ValueBrightness"
        Me.ValueBrightness.Size = New System.Drawing.Size(48, 20)
        Me.ValueBrightness.TabIndex = 37
        Me.ValueBrightness.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueBrightness.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'ValueThreshold
        '
        Me.ValueThreshold.Location = New System.Drawing.Point(16, 96)
        Me.ValueThreshold.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.ValueThreshold.Name = "ValueThreshold"
        Me.ValueThreshold.Size = New System.Drawing.Size(48, 20)
        Me.ValueThreshold.TabIndex = 36
        Me.ValueThreshold.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ValueThreshold.Value = New Decimal(New Integer() {15, 0, 0, 0})
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.RadioButton_Outside_In)
        Me.GroupBox3.Controls.Add(Me.RadioButton_Inside_out)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(192, 48)
        Me.GroupBox3.TabIndex = 46
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Foreground/Background:"
        '
        'RadioButton_Outside_In
        '
        Me.RadioButton_Outside_In.Location = New System.Drawing.Point(96, 16)
        Me.RadioButton_Outside_In.Name = "RadioButton_Outside_In"
        Me.RadioButton_Outside_In.Size = New System.Drawing.Size(80, 24)
        Me.RadioButton_Outside_In.TabIndex = 14
        Me.RadioButton_Outside_In.Text = "Outside_In"
        '
        'RadioButton_Inside_out
        '
        Me.RadioButton_Inside_out.Checked = True
        Me.RadioButton_Inside_out.Location = New System.Drawing.Point(8, 16)
        Me.RadioButton_Inside_out.Name = "RadioButton_Inside_out"
        Me.RadioButton_Inside_out.Size = New System.Drawing.Size(78, 24)
        Me.RadioButton_Inside_out.TabIndex = 13
        Me.RadioButton_Inside_out.TabStop = True
        Me.RadioButton_Inside_out.Text = "Inside_Out"
        '
        'GroupBox_Vertical_Horizontal
        '
        Me.GroupBox_Vertical_Horizontal.Controls.Add(Me.RadioButton_Horizontal)
        Me.GroupBox_Vertical_Horizontal.Controls.Add(Me.RadioButton_Vertical)
        Me.GroupBox_Vertical_Horizontal.Location = New System.Drawing.Point(224, 16)
        Me.GroupBox_Vertical_Horizontal.Name = "GroupBox_Vertical_Horizontal"
        Me.GroupBox_Vertical_Horizontal.Size = New System.Drawing.Size(248, 48)
        Me.GroupBox_Vertical_Horizontal.TabIndex = 47
        Me.GroupBox_Vertical_Horizontal.TabStop = False
        Me.GroupBox_Vertical_Horizontal.Text = "Vertical/Horizontal:"
        '
        'RadioButton_Horizontal
        '
        Me.RadioButton_Horizontal.Location = New System.Drawing.Point(160, 16)
        Me.RadioButton_Horizontal.Name = "RadioButton_Horizontal"
        Me.RadioButton_Horizontal.Size = New System.Drawing.Size(80, 24)
        Me.RadioButton_Horizontal.TabIndex = 67
        Me.RadioButton_Horizontal.Text = "Horizontal"
        '
        'RadioButton_Vertical
        '
        Me.RadioButton_Vertical.Checked = True
        Me.RadioButton_Vertical.Location = New System.Drawing.Point(16, 16)
        Me.RadioButton_Vertical.Name = "RadioButton_Vertical"
        Me.RadioButton_Vertical.Size = New System.Drawing.Size(72, 24)
        Me.RadioButton_Vertical.TabIndex = 66
        Me.RadioButton_Vertical.TabStop = True
        Me.RadioButton_Vertical.Text = "Vertical"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.TextBox_EdgeClearance)
        Me.GroupBox4.Controls.Add(Me.CheckBox_ChipRec_Enable)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 128)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(192, 72)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Chip's Edge Clearance"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(144, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "mm"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Edge Clearance:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_EdgeClearance
        '
        Me.TextBox_EdgeClearance.Location = New System.Drawing.Point(80, 40)
        Me.TextBox_EdgeClearance.Name = "TextBox_EdgeClearance"
        Me.TextBox_EdgeClearance.Size = New System.Drawing.Size(56, 20)
        Me.TextBox_EdgeClearance.TabIndex = 4
        Me.TextBox_EdgeClearance.Text = "1.000"
        Me.TextBox_EdgeClearance.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox_ChipRec_Enable
        '
        Me.CheckBox_ChipRec_Enable.Checked = True
        Me.CheckBox_ChipRec_Enable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_ChipRec_Enable.Location = New System.Drawing.Point(8, 16)
        Me.CheckBox_ChipRec_Enable.Name = "CheckBox_ChipRec_Enable"
        Me.CheckBox_ChipRec_Enable.Size = New System.Drawing.Size(160, 24)
        Me.CheckBox_ChipRec_Enable.TabIndex = 4
        Me.CheckBox_ChipRec_Enable.Text = "Enable Chip Recognition"
        '
        'GroupBox_DispenseModel
        '
        Me.GroupBox_DispenseModel.Controls.Add(Me.RadioButton_TwoEdges)
        Me.GroupBox_DispenseModel.Controls.Add(Me.RadioButton_OneEdge)
        Me.GroupBox_DispenseModel.Controls.Add(Me.RadioButton_Dot)
        Me.GroupBox_DispenseModel.Location = New System.Drawing.Point(224, 128)
        Me.GroupBox_DispenseModel.Name = "GroupBox_DispenseModel"
        Me.GroupBox_DispenseModel.Size = New System.Drawing.Size(248, 48)
        Me.GroupBox_DispenseModel.TabIndex = 62
        Me.GroupBox_DispenseModel.TabStop = False
        Me.GroupBox_DispenseModel.Text = "Dispense Model:"
        '
        'RadioButton_TwoEdges
        '
        Me.RadioButton_TwoEdges.Checked = True
        Me.RadioButton_TwoEdges.Location = New System.Drawing.Point(164, 16)
        Me.RadioButton_TwoEdges.Name = "RadioButton_TwoEdges"
        Me.RadioButton_TwoEdges.Size = New System.Drawing.Size(80, 24)
        Me.RadioButton_TwoEdges.TabIndex = 7
        Me.RadioButton_TwoEdges.TabStop = True
        Me.RadioButton_TwoEdges.Text = "Two Edges"
        '
        'RadioButton_OneEdge
        '
        Me.RadioButton_OneEdge.Location = New System.Drawing.Point(80, 16)
        Me.RadioButton_OneEdge.Name = "RadioButton_OneEdge"
        Me.RadioButton_OneEdge.Size = New System.Drawing.Size(74, 24)
        Me.RadioButton_OneEdge.TabIndex = 6
        Me.RadioButton_OneEdge.Text = "One Edge"
        '
        'RadioButton_Dot
        '
        Me.RadioButton_Dot.Location = New System.Drawing.Point(16, 16)
        Me.RadioButton_Dot.Name = "RadioButton_Dot"
        Me.RadioButton_Dot.Size = New System.Drawing.Size(48, 24)
        Me.RadioButton_Dot.TabIndex = 5
        Me.RadioButton_Dot.Text = "Dot"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Button1)
        Me.GroupBox5.Controls.Add(Me.RichTextBox1)
        Me.GroupBox5.Controls.Add(Me.TextBox_RRot)
        Me.GroupBox5.Controls.Add(Me.TextBox_RPosY)
        Me.GroupBox5.Controls.Add(Me.TextBox_RPosX)
        Me.GroupBox5.Controls.Add(Me.TextBox_RSizeY)
        Me.GroupBox5.Controls.Add(Me.Label38)
        Me.GroupBox5.Controls.Add(Me.Label27)
        Me.GroupBox5.Controls.Add(Me.Label28)
        Me.GroupBox5.Controls.Add(Me.Label29)
        Me.GroupBox5.Controls.Add(Me.Label30)
        Me.GroupBox5.Controls.Add(Me.Label31)
        Me.GroupBox5.Controls.Add(Me.Label32)
        Me.GroupBox5.Controls.Add(Me.Label33)
        Me.GroupBox5.Controls.Add(Me.Label34)
        Me.GroupBox5.Controls.Add(Me.Label35)
        Me.GroupBox5.Controls.Add(Me.Label36)
        Me.GroupBox5.Controls.Add(Me.Label37)
        Me.GroupBox5.Controls.Add(Me.Button_Test)
        Me.GroupBox5.Controls.Add(Me.Label26)
        Me.GroupBox5.Controls.Add(Me.TextBox_Score)
        Me.GroupBox5.Controls.Add(Me.TextBox_RSizeX)
        Me.GroupBox5.Controls.Add(Me.Button_Reset)
        Me.GroupBox5.Location = New System.Drawing.Point(808, 16)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(432, 256)
        Me.GroupBox5.TabIndex = 61
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Results"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(80, 192)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 67
        Me.Button1.Text = "Button1"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.RichTextBox1.Location = New System.Drawing.Point(296, 16)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(128, 232)
        Me.RichTextBox1.TabIndex = 66
        Me.RichTextBox1.Text = "Tips:"
        '
        'TextBox_RRot
        '
        Me.TextBox_RRot.Location = New System.Drawing.Point(88, 112)
        Me.TextBox_RRot.Name = "TextBox_RRot"
        Me.TextBox_RRot.ReadOnly = True
        Me.TextBox_RRot.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_RRot.TabIndex = 65
        Me.TextBox_RRot.Text = "0"
        Me.TextBox_RRot.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RPosY
        '
        Me.TextBox_RPosY.Location = New System.Drawing.Point(88, 88)
        Me.TextBox_RPosY.Name = "TextBox_RPosY"
        Me.TextBox_RPosY.ReadOnly = True
        Me.TextBox_RPosY.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_RPosY.TabIndex = 64
        Me.TextBox_RPosY.Text = "0"
        Me.TextBox_RPosY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RPosX
        '
        Me.TextBox_RPosX.Location = New System.Drawing.Point(88, 64)
        Me.TextBox_RPosX.Name = "TextBox_RPosX"
        Me.TextBox_RPosX.ReadOnly = True
        Me.TextBox_RPosX.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_RPosX.TabIndex = 63
        Me.TextBox_RPosX.Text = "0"
        Me.TextBox_RPosX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RSizeY
        '
        Me.TextBox_RSizeY.Location = New System.Drawing.Point(88, 40)
        Me.TextBox_RSizeY.Name = "TextBox_RSizeY"
        Me.TextBox_RSizeY.ReadOnly = True
        Me.TextBox_RSizeY.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_RSizeY.TabIndex = 62
        Me.TextBox_RSizeY.Text = "0"
        Me.TextBox_RSizeY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(136, 40)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(32, 23)
        Me.Label38.TabIndex = 44
        Me.Label38.Text = "mm"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(136, 88)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(32, 23)
        Me.Label27.TabIndex = 55
        Me.Label27.Text = "mm"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(136, 64)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(32, 23)
        Me.Label28.TabIndex = 54
        Me.Label28.Text = "mm"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(64, 88)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(24, 24)
        Me.Label29.TabIndex = 53
        Me.Label29.Text = "Y="
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(64, 40)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(24, 24)
        Me.Label30.TabIndex = 52
        Me.Label30.Text = "Y="
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label31
        '
        Me.Label31.Location = New System.Drawing.Point(64, 64)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(24, 24)
        Me.Label31.TabIndex = 51
        Me.Label31.Text = "X="
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label32
        '
        Me.Label32.Location = New System.Drawing.Point(64, 16)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(24, 24)
        Me.Label32.TabIndex = 50
        Me.Label32.Text = "X="
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(8, 112)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(56, 16)
        Me.Label33.TabIndex = 49
        Me.Label33.Text = "Rotation:"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label34
        '
        Me.Label34.Location = New System.Drawing.Point(8, 64)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(56, 24)
        Me.Label34.TabIndex = 48
        Me.Label34.Text = "Position:"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(8, 16)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(56, 24)
        Me.Label35.TabIndex = 47
        Me.Label35.Text = "Size:"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label36
        '
        Me.Label36.Location = New System.Drawing.Point(136, 112)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(32, 16)
        Me.Label36.TabIndex = 46
        Me.Label36.Text = "deg"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(136, 16)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(32, 23)
        Me.Label37.TabIndex = 45
        Me.Label37.Text = "mm"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button_Test
        '
        Me.Button_Test.Enabled = False
        Me.Button_Test.Location = New System.Drawing.Point(184, 42)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(104, 208)
        Me.Button_Test.TabIndex = 11
        Me.Button_Test.Text = "Test"
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(8, 144)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(56, 24)
        Me.Label26.TabIndex = 42
        Me.Label26.Text = "Score:"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Score
        '
        Me.TextBox_Score.Location = New System.Drawing.Point(88, 144)
        Me.TextBox_Score.Name = "TextBox_Score"
        Me.TextBox_Score.ReadOnly = True
        Me.TextBox_Score.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_Score.TabIndex = 43
        Me.TextBox_Score.Text = "0"
        Me.TextBox_Score.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RSizeX
        '
        Me.TextBox_RSizeX.Location = New System.Drawing.Point(88, 16)
        Me.TextBox_RSizeX.Name = "TextBox_RSizeX"
        Me.TextBox_RSizeX.ReadOnly = True
        Me.TextBox_RSizeX.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_RSizeX.TabIndex = 61
        Me.TextBox_RSizeX.Text = "0"
        Me.TextBox_RSizeX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_Reset
        '
        Me.Button_Reset.Location = New System.Drawing.Point(184, 16)
        Me.Button_Reset.Name = "Button_Reset"
        Me.Button_Reset.Size = New System.Drawing.Size(104, 23)
        Me.Button_Reset.TabIndex = 62
        Me.Button_Reset.Text = "Reset"
        '
        'ChipEdgePoints
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1285, 320)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "ChipEdgePoints"
        Me.Text = "ChipEdgePoints"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.ValueRot, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValuePos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_Settings.ResumeLayout(False)
        CType(Me.ValueROI, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        CType(Me.ValueContrast, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueThreshold, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox_Vertical_Horizontal.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox_DispenseModel.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Shared m_instance As ChipEdgePoints

    Public Shared ReadOnly Property Instance() As ChipEdgePoints
        Get
            If m_instance Is Nothing Then
                m_instance = New ChipEdgePoints
            Else
                If m_instance.IsDisposed Then
                    m_instance = New ChipEdgePoints
                End If
            End If
            Return m_instance
        End Get

    End Property

    Public Structure ChipEdgeParam
        Public _SizeX As Double
        Public _SizeY As Double
        Public _PosX As Double
        Public _PosY As Double
        Public _Size As Double
        Public _Pos As Double
        Public _Rot As Double
        Public _Inside_out As Boolean
        Public _Cw_CCw As Boolean
        Public _Threshold As Double
        Public _ROI As Double
        Public _Brightness As Integer
        Public _Vertical As Boolean
        Public _MainEdge As Integer
        Public _PointX1 As Double
        Public _PointY1 As Double
        Public _PointX2 As Double
        Public _PointY2 As Double
        Public _PointX3 As Double
        Public _PointY3 As Double
        Public _PointX4 As Double
        Public _PointY4 As Double
        Public _PointX5 As Double
        Public _PointY5 As Double
        Public _DispenseModel As Integer
        Public _EdgeClearance As Double
        Public _CheckBox_ChipRec_Enable As Boolean
    End Structure

    Dim Folder_CE As String = "C:\IDS\SRC\DLL Export Device Vision\ChipEdgePoint\"
#Region "GlobalVariable"
    Dim Clickno As Integer = 1
    Dim Test_Click As Boolean = False
    Dim Inside_out As Boolean = True
    Dim MainEdge As Integer = 1
    Dim P1x, P1y, P2x, P2y, P3x, P3y, P4x, P4y, P5x, P5y As Double
    Dim Vertical As Boolean
    Dim PointX1a, PointY1a, PointX2a, PointY2a, PointX3a, PointY3a, PointX4a, PointY4a, PointX5a, PointY5a As Double
    Dim CW_CCW As Boolean = True
    Dim DispenseModel As Integer = 2
#End Region
    Dim Status As Integer = 0 '0=yet done, 1=done, 2=cancel, 3=reset
    Sub ResetChipEdgeStatus()
        Status = 0
    End Sub
    Function GetChipEdgeStatus() As Integer
        Return Status
    End Function

    Function GetChipEdgeParameters(ByRef edgeParam As ChipEdgeParam)
        edgeParam._SizeX = TextBox_SizeX.Text
        edgeParam._SizeY = TextBox_SizeY.Text
        edgeParam._PosX = TextBox_PosX.Text
        edgeParam._PosY = TextBox_PosY.Text
        edgeParam._Size = ValueSize.Value
        edgeParam._Pos = ValuePos.Value
        edgeParam._Rot = ValueRot.Value
        edgeParam._Inside_out = Inside_out
        edgeParam._Cw_CCw = CW_CCW
        edgeParam._Threshold = ValueThreshold.Value
        edgeParam._ROI = ValueROI.Value
        edgeParam._Brightness = ValueBrightness.Value
        edgeParam._Vertical = Vertical
        edgeParam._MainEdge = MainEdge
        edgeParam._PointX1 = P1x
        edgeParam._PointY1 = P1y
        edgeParam._PointX2 = P2x
        edgeParam._PointY2 = P2y
        edgeParam._PointX3 = P3x
        edgeParam._PointY3 = P3y
        edgeParam._PointX4 = P4x
        edgeParam._PointY4 = P4y
        edgeParam._PointX5 = P5x
        edgeParam._PointY5 = P5y
        edgeParam._DispenseModel = DispenseModel
        edgeParam._EdgeClearance = TextBox_EdgeClearance.Text
        edgeParam._CheckBox_ChipRec_Enable = CheckBox_ChipRec_Enable.Checked
    End Function
    Function GetChipEdgeParameters_bak(ByRef _SizeX As Double, ByRef _SizeY As Double, ByRef _PosX As Double, ByRef _PosY As Double, ByRef _Size As Double, ByRef _Pos As Double, ByRef _Rot As Double, ByRef _Inside_out As Boolean, ByVal _Cw_CCw As Boolean, ByRef _Threshold As Double, ByRef _ROI As Double, ByRef _Brightness As Integer, ByRef _Vertical As Boolean, ByRef _MainEdge As Integer, ByRef _PointX1 As Double, ByRef _PointY1 As Double, ByRef _PointX2 As Double, ByRef _PointY2 As Double, ByRef _PointX3 As Double, ByRef _PointY3 As Double, ByRef _PointX4 As Double, ByRef _PointY4 As Double, ByRef _PointX5 As Double, ByRef _PointY5 As Double, ByRef _DispenseModel As Integer, ByRef _EdgeClearance As Double, ByRef _CheckBox_ChipRec_Enable As Boolean)
        _SizeX = TextBox_SizeX.Text
        _SizeY = TextBox_SizeY.Text
        _PosX = TextBox_PosX.Text
        _PosY = TextBox_PosY.Text
        _Size = ValueSize.Value
        _Pos = ValuePos.Value
        _Rot = ValueRot.Value
        _Inside_out = Inside_out
        _Cw_CCw = CW_CCW
        _Threshold = ValueThreshold.Value
        _ROI = ValueROI.Value
        _Brightness = ValueBrightness.Value
        _Vertical = Vertical
        _MainEdge = MainEdge
        _PointX1 = P1x
        _PointY1 = P1y
        _PointX2 = P2x
        _PointY2 = P2y
        _PointX3 = P3x
        _PointY3 = P3y
        _PointX4 = P4x
        _PointY4 = P4y
        _PointX5 = P5x
        _PointY5 = P5y
        _DispenseModel = DispenseModel
        _EdgeClearance = TextBox_EdgeClearance.Text
        _CheckBox_ChipRec_Enable = CheckBox_ChipRec_Enable.Checked
    End Function
    Sub ResetVariables()
        'Reset all variables
        TextBox_SizeX.Text = 0
        TextBox_SizeY.Text = 0
        TextBox_PosX.Text = 0
        TextBox_PosY.Text = 0
        ValueSize.Text = 5
        ValuePos.Text = 5
        ValueRot.Text = 5
        RadioButton_Inside_out.Checked = True
        RadioButton_Outside_In.Checked = False
        RadioButton_CW.Checked = True
        ValueThreshold.Value = 15
        ValueROI.Value = 2
        ValueBrightness.Value = 6
        RadioButton_Vertical.Checked = True
        RadioButton_Horizontal.Checked = False
        MainEdge = 1
        RadioButton_TwoEdges.Checked = True
        TextBox_EdgeClearance.Text = 1.0
        CheckBox_ChipRec_Enable.Checked = True

        TextBox_RSizeX.Text = 0
        TextBox_RSizeY.Text = 0
        TextBox_RPosX.Text = 0
        TextBox_RPosY.Text = 0
        TextBox_RRot.Text = 0
        TextBox_Score.Text = 0
        FrmVision.ResetChipEdgePoint()

        PointX1a = 0
        PointY1a = 0
        PointX2a = 0
        PointY2a = 0
        PointX3a = 0
        PointY3a = 0
        PointX4a = 0
        PointY4a = 0
        PointX5a = 0
        PointY5a = 0
        P1x = 0
        P2x = 0
        P3x = 0
        P4x = 0
        P5x = 0
        P1y = 0
        P2y = 0
        P3y = 0
        P4y = 0
        P5y = 0
        Clickno = 1
        Test_Click = False

        GroupBox_Vertical_Horizontal.Enabled = True
        Timer1.Stop()
        Button_Test.Text = "Test"
        Button_Ok.Enabled = False
        FrmVision.DisableChipEdgeDrawing()
    End Sub
    Function ChipEdgeOutput(ByRef PointX1, ByRef PointY1, ByRef PointX2, ByRef PointY2, ByRef PointX3, ByRef PointY3, ByRef PointX4, ByRef PointY4, ByRef MEdge)
        'Production Output
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
    Function Vertical_Horizontal() As Boolean
        Return Vertical
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
        NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\RejectPoint" & "\" & FileName & "." & FileExt
        Flag = False

        'Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = True Then
            'Flag = True
            Dim FileName_checked = Mid(NewFullPath, InStrRev(NewFullPath, "\") + 1)
            Return True
        Else
            'Counter = Counter + 1
            Return False
            'NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\RejectPoint" & "\" & strTempFileName & Counter & "." & FileExt
        End If
        'Loop
    End Function
#Region "Production"
    Function ChipEdgePointExe() As Boolean
        FrmVision.ClearDisplay()
        Dim result As Boolean = FrmVision.MeasurementPoint(ValueContrast.Value, ValueThreshold.Value, ValueRot.Value, Inside_out, Vertical, ValueROI.Value, 1)
        If result = False Then
            ResetVariables()
            Return False
        Else
            FrmVision.SearchRegionPoints(1, ValueROI.Value)
            FrmVision.ChipPointOutPut(1)
            'MsgBox(PointX1a.ToString & " " & PointY1a.ToString & " " & PointX2a.ToString & " " & PointY2a.ToString & " " & PointX3a.ToString & " " & PointY3a.ToString & " " & PointX4a.ToString & " " & PointY4a.ToString & " " & MainEdge.ToString & " " & TextBox_EdgeClearance.Text.ToString)
            ResetVariables()
            Return True
        End If
    End Function
    Function GetVariables(ByVal FrmInside_out As Boolean, ByVal FrmVertical As Boolean, ByVal FrmCw_CCw As Boolean, ByVal FrmDispenseModel As Integer)
        Inside_out = FrmInside_out
        Vertical = FrmVertical
        CW_CCW = FrmCw_CCw
        DispenseModel = FrmDispenseModel
    End Function
#End Region

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        FrmVision.ClearDisplay()
        FrmVision.MeasurementPoint(ValueContrast.Value, ValueThreshold.Value, ValueRot.Value, Inside_out, Vertical, ValueROI.Value, 1)
        If Clickno = 1 Then
            'FrmVision.SearchRegionResultsPoints(1)
            FrmVision.SearchRegionPoints(1, ValueROI.Value)
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Clickno = 1 Then
            Clickno = 0
        ElseIf Clickno = 0 Then
            Clickno = 1
        End If
    End Sub
    Private Sub Button_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Test.Click
        If Test_Click = False Then
            FrmVision.SetBrightness(ValueBrightness.Value)
            Timer1.Start()
            Button_Test.Text = "Stop"
            Test_Click = True
        ElseIf Test_Click = True Then
            Timer1.Stop()
            Test_Click = False
            Button_Test.Text = "Test"
            FrmVision.ClearDisplay()
        End If
    End Sub
    Private Sub Button_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Cancel.Click
        Status = 2 'For SJ to check if CE is cancel
        FrmVision.ResetChipEdgePoint()
        ResetVariables()
        FrmVision.DisplayIndicator()
        FrmVision.DisableChipEdgeDrawing()
        Me.Visible = False
    End Sub
    Private Sub ChipEdgePoints_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True
        RichTextBox1.Text = "Tips:" & Chr(13) & Chr(10) & _
        "1) Choose vertical or horizontal " & Chr(13) & Chr(10) & _
        "2) Click 5 points according to vertical or horizontal edge " & Chr(13) & Chr(10) & _
        "3) Adjust size of ROI" & Chr(13) & Chr(10) & _
        "4) Adjust edge threshold"

        'Reset All value
    End Sub
    Private Sub RadioButton_Inside_out_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Inside_out.CheckedChanged
        Inside_out = True
    End Sub
    Private Sub RadioButton_Outside_In_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Outside_In.CheckedChanged
        Inside_out = False
    End Sub
    Private Sub RadioButton_Verical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Vertical.CheckedChanged
        Vertical = True
    End Sub
    Private Sub RadioButton_Horizontal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Horizontal.CheckedChanged
        Vertical = False
    End Sub
    Private Sub RadioButton_CW_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_CW.CheckedChanged
        If RadioButton_CW.Checked = True Then
            CW_CCW = True
        End If
    End Sub
    Private Sub RadioButton_CCW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_CCW.CheckedChanged
        If RadioButton_CCW.Checked = True Then
            CW_CCW = False
        End If
    End Sub
    Private Sub RadioButton_TwoEdges_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_TwoEdges.CheckedChanged
        DispenseModel = 2
    End Sub
    Private Sub RadioButton_OneEdge_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_OneEdge.CheckedChanged
        DispenseModel = 1
    End Sub
    Private Sub RadioButton_Dot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Dot.CheckedChanged
        DispenseModel = 0
    End Sub
    Private Sub CheckBox_ChipRec_Enable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox_ChipRec_Enable.CheckedChanged
        If CheckBox_ChipRec_Enable.Checked = False Then
            'GroupBox_Settings.Enabled = False
            GroupBox_DispenseModel.Enabled = False
            GroupBox_Vertical_Horizontal.Enabled = False
            GroupBox7.Enabled = False
            GroupBox3.Enabled = False
            GroupBox5.Enabled = False
            FrmVision.ClearDisplay()
            FrmVision.ClearanceRectangle()
            FrmVision.SearchRegionPoints(1, ValueROI.Value)
            'FrmVision.Chipedgepoint()
            '---------------

            RadioButton_Inside_out.Checked = True
            RadioButton_Outside_In.Checked = False
            RadioButton_CW.Checked = True
            ValueThreshold.Value = 15
            ValueROI.Value = 2
            ValueBrightness.Value = 6
            RadioButton_Vertical.Checked = True
            RadioButton_Horizontal.Checked = False

            TextBox_RSizeX.Text = 0
            TextBox_RSizeY.Text = 0
            TextBox_RPosX.Text = 0
            TextBox_RPosY.Text = 0
            TextBox_RRot.Text = 0
            TextBox_Score.Text = 0

            PointX1a = 0
            PointY1a = 0
            PointX2a = 0
            PointY2a = 0
            PointX3a = 0
            PointY3a = 0
            PointX4a = 0
            PointY4a = 0
            PointX5a = 0
            PointY5a = 0

            Clickno = 1
            Test_Click = False

            Timer1.Stop()
            Button_Test.Text = "Test"
        Else
            GroupBox_DispenseModel.Enabled = True
            GroupBox_Vertical_Horizontal.Enabled = True
            GroupBox7.Enabled = True
            GroupBox3.Enabled = True
            GroupBox5.Enabled = True
        End If
    End Sub
    Private Sub Button_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Ok.Click
        Timer1.Stop()
        Status = 1 'for SJ to check if CE done
        FrmVision.DisplayIndicator()
        FrmVision.ChipPointOutPut(1) 'rearrange points for output
        Dim vParam As DLL_Export_Device_Vision.ChipEdgePoints.ChipEdgeParam
        FrmVision.GetChipEdgeParameters(vParam)
        FrmVision.DisplayIndicator()
        Me.Visible = False
    End Sub

    Private Sub Button_Reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Reset.Click
        Status = 3 'reset being pressed
        FrmVision.ResetChipEdgePoint()
        ResetVariables()
        FrmVision.EnableChipEdgeDrawing()
        FrmVision.DisplayIndicator()
    End Sub

    Private Sub ValueBrightness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueBrightness.TextChanged
        If ValueBrightness.Focused = True Then  'Xu Long 20/02/2012
            CloseTimer1()
            If (ValueBrightness.Value >= 1) And (ValueBrightness.Value <= 255) Then
                FrmVision.SetBrightness(ValueBrightness.Value)
            Else
                ValueBrightness.Value = 128
                FrmVision.SetBrightness(ValueBrightness.Value)
            End If
            OnTimer1()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FrmVision.ClearDisplay()
        FrmVision.MeasurementPoint(ValueContrast.Value, ValueThreshold.Value, ValueRot.Value, Inside_out, Vertical, ValueROI.Value, 1)
        If Clickno = 1 Then
            'FrmVision.SearchRegionResultsPoints(1)
            FrmVision.SearchRegionPoints(1, ValueROI.Value)
        End If
    End Sub

#Region "Xue Wen"
    '''''''''''''''''''''''''''''
    '   Setting validation.     '
    '''''''''''''''''''''''''''''
    Private Sub TextBox_EdgeClearance_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_EdgeClearance.TextChanged
        CloseTimer1()

        If Not IsNumeric(TextBox_EdgeClearance.Text) Then
            TextBox_EdgeClearance.Text = "0.001"
        ElseIf (TextBox_EdgeClearance.Text = "") Then
            TextBox_EdgeClearance.Text = "0.001"
        Else
            If (CDbl(TextBox_EdgeClearance.Text) < 0) Then
                TextBox_EdgeClearance.Text = "0.001"
            End If
        End If

        OnTimer1()
    End Sub

    Private Sub ValueThreshold_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueThreshold.ValueChanged
        CloseTimer1()

        If (CStr(ValueThreshold.Value) = "") Then
            ValueThreshold.Value = 15
        Else
            If Not ((ValueThreshold.Value >= 1) And (ValueThreshold.Value <= 100)) Then
                ValueThreshold.Value = 15
            End If
        End If

        OnTimer1()
    End Sub

    Private Sub ValueROI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueROI.ValueChanged
        CloseTimer1()

        If (CStr(ValueROI.Value) = "") Then
            ValueROI.Value = 2.0
        Else
            If Not ((ValueROI.Value >= 0.1) And (ValueROI.Value <= 3)) Then
                ValueROI.Value = 2.0
            End If
        End If

        OnTimer1()
    End Sub

    Private Sub CloseTimer1()
        If (Test_Click = True) Then
            Timer1.Enabled = False
        End If
    End Sub

    Private Sub OnTimer1()
        If (Test_Click = True) Then
            Timer1.Enabled = True
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''
    '   Chip size position validation.      '
    '''''''''''''''''''''''''''''''''''''''''
    Private Sub TextBox_SizeX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_SizeX.TextChanged
        CloseTimer1()

        If Not IsNumeric(TextBox_SizeX.Text) Then
            TextBox_SizeX.Text = "0"
        ElseIf (TextBox_SizeX.Text = "") Then
            TextBox_SizeX.Text = "0"
        ElseIf (CDbl(TextBox_SizeX.Text) < 0) Then
            TextBox_SizeX.Text = CStr(CDbl(TextBox_SizeX.Text) * (-1))
        End If

        OnTimer1()
    End Sub

    Private Sub TextBox_SizeY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_SizeY.TextChanged
        CloseTimer1()

        If Not IsNumeric(TextBox_SizeY.Text) Then
            TextBox_SizeY.Text = "0"
        ElseIf (TextBox_SizeY.Text = "") Then
            TextBox_SizeY.Text = "0"
        ElseIf (CDbl(TextBox_SizeY.Text) < 0) Then
            TextBox_SizeY.Text = CStr(CDbl(TextBox_SizeY.Text) * (-1))
        End If

        OnTimer1()
    End Sub

    Private Sub TextBox_PosX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PosX.TextChanged

    End Sub

    Private Sub TextBox_PosY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox_PosY.TextChanged

    End Sub

    Private Sub ValueRot_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueRot.ValueChanged

    End Sub

    Private Sub ValueSize_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueSize.ValueChanged

    End Sub

    Private Sub ValuePos_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValuePos.ValueChanged

    End Sub
#End Region

End Class
