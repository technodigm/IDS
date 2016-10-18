Imports System.io

Public Class FiducialForm
    Inherits System.Windows.Forms.Form

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Xue Wen                                                                                     '
    '   Another method, Not to fire "ValueChanged" event while constructing Form's components.      '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Initializing As Boolean = True
    Public Delegate Sub FormCloseDelegate()
    Public FormCloseEvent As FormCloseDelegate = Nothing
    Public isEdit As Boolean = False

#Region " Windows Form Designer generated code "
    Dim frm As FormVision

    Public Sub New()

        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Dim FormMain As New FormExport
        'FormMain.ShowDialog()
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
    Friend WithEvents Button_Fid_Ok As System.Windows.Forms.Button
    Friend WithEvents Button_Fid_Cancel As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox8 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox12 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox13 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox14 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox15 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox16 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox17 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel_BlackOnWhite As System.Windows.Forms.Panel
    Friend WithEvents Panel_WhiteOnBlack As System.Windows.Forms.Panel
    Friend WithEvents PictureBox18 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox19 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox20 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_Score As System.Windows.Forms.TextBox
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Size As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Teach As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button_Load As System.Windows.Forms.Button
    Friend WithEvents Button_SettingOK As System.Windows.Forms.Button
    Friend WithEvents ValueX As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueY As System.Windows.Forms.NumericUpDown
    Friend WithEvents valuedetecteddiameter As System.Windows.Forms.NumericUpDown
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents TextBox_RPosY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RPosX As System.Windows.Forms.TextBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ValueThickness As System.Windows.Forms.NumericUpDown
    Friend WithEvents ValueInDia As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button_Diameter As System.Windows.Forms.Button
    Friend WithEvents Button_Distance As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox_X As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Y As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Dia As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Dist As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_EndFi As System.Windows.Forms.Button
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents BrightnessValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents tbModelInfo As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FiducialForm))
        Me.Button_Fid_Ok = New System.Windows.Forms.Button
        Me.Button_Fid_Cancel = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Panel_BlackOnWhite = New System.Windows.Forms.Panel
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox6 = New System.Windows.Forms.PictureBox
        Me.PictureBox7 = New System.Windows.Forms.PictureBox
        Me.PictureBox9 = New System.Windows.Forms.PictureBox
        Me.PictureBox8 = New System.Windows.Forms.PictureBox
        Me.Panel_WhiteOnBlack = New System.Windows.Forms.Panel
        Me.PictureBox18 = New System.Windows.Forms.PictureBox
        Me.PictureBox11 = New System.Windows.Forms.PictureBox
        Me.PictureBox12 = New System.Windows.Forms.PictureBox
        Me.PictureBox13 = New System.Windows.Forms.PictureBox
        Me.PictureBox14 = New System.Windows.Forms.PictureBox
        Me.PictureBox15 = New System.Windows.Forms.PictureBox
        Me.PictureBox16 = New System.Windows.Forms.PictureBox
        Me.PictureBox17 = New System.Windows.Forms.PictureBox
        Me.PictureBox19 = New System.Windows.Forms.PictureBox
        Me.PictureBox10 = New System.Windows.Forms.PictureBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox_Size = New System.Windows.Forms.GroupBox
        Me.BrightnessValue = New System.Windows.Forms.NumericUpDown
        Me.Label15 = New System.Windows.Forms.Label
        Me.ValueThickness = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.ValueInDia = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.valuedetecteddiameter = New System.Windows.Forms.NumericUpDown
        Me.ValueY = New System.Windows.Forms.NumericUpDown
        Me.ValueX = New System.Windows.Forms.NumericUpDown
        Me.Button_SettingOK = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button_Test = New System.Windows.Forms.Button
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Label46 = New System.Windows.Forms.Label
        Me.Label44 = New System.Windows.Forms.Label
        Me.Label43 = New System.Windows.Forms.Label
        Me.Label42 = New System.Windows.Forms.Label
        Me.TextBox_RPosX = New System.Windows.Forms.TextBox
        Me.TextBox_RPosY = New System.Windows.Forms.TextBox
        Me.TextBox_Score = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.TextBox_Dia = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TextBox_X = New System.Windows.Forms.TextBox
        Me.TextBox_Y = New System.Windows.Forms.TextBox
        Me.Button_Diameter = New System.Windows.Forms.Button
        Me.Button_Distance = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.TextBox_Dist = New System.Windows.Forms.TextBox
        Me.Button_Teach = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.PictureBox20 = New System.Windows.Forms.PictureBox
        Me.Button_EndFi = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.tbModelInfo = New System.Windows.Forms.TextBox
        Me.Button_Load = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.Panel_BlackOnWhite.SuspendLayout()
        Me.Panel_WhiteOnBlack.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox_Size.SuspendLayout()
        CType(Me.BrightnessValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueThickness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueInDia, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.valuedetecteddiameter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ValueX, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button_Fid_Ok
        '
        Me.Button_Fid_Ok.Location = New System.Drawing.Point(1184, 304)
        Me.Button_Fid_Ok.Name = "Button_Fid_Ok"
        Me.Button_Fid_Ok.Size = New System.Drawing.Size(80, 25)
        Me.Button_Fid_Ok.TabIndex = 0
        Me.Button_Fid_Ok.Text = "Ok"
        '
        'Button_Fid_Cancel
        '
        Me.Button_Fid_Cancel.Location = New System.Drawing.Point(1080, 304)
        Me.Button_Fid_Cancel.Name = "Button_Fid_Cancel"
        Me.Button_Fid_Cancel.Size = New System.Drawing.Size(100, 25)
        Me.Button_Fid_Cancel.TabIndex = 1
        Me.Button_Fid_Cancel.Text = "Cancel"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'Panel_BlackOnWhite
        '
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox3)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox1)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox2)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox4)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox5)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox6)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox7)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox9)
        Me.Panel_BlackOnWhite.Controls.Add(Me.PictureBox8)
        Me.Panel_BlackOnWhite.Location = New System.Drawing.Point(-6, -6)
        Me.Panel_BlackOnWhite.Name = "Panel_BlackOnWhite"
        Me.Panel_BlackOnWhite.Size = New System.Drawing.Size(203, 200)
        Me.Panel_BlackOnWhite.TabIndex = 9
        '
        'PictureBox3
        '
        Me.PictureBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(136, 8)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox3.TabIndex = 11
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(72, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox2.TabIndex = 10
        Me.PictureBox2.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(8, 72)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox4.TabIndex = 10
        Me.PictureBox4.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(72, 72)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox5.TabIndex = 10
        Me.PictureBox5.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(136, 72)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox6.TabIndex = 10
        Me.PictureBox6.TabStop = False
        '
        'PictureBox7
        '
        Me.PictureBox7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(8, 136)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox7.TabIndex = 10
        Me.PictureBox7.TabStop = False
        '
        'PictureBox9
        '
        Me.PictureBox9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox9.Enabled = False
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(136, 136)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox9.TabIndex = 11
        Me.PictureBox9.TabStop = False
        '
        'PictureBox8
        '
        Me.PictureBox8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox8.Enabled = False
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(72, 136)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox8.TabIndex = 10
        Me.PictureBox8.TabStop = False
        '
        'Panel_WhiteOnBlack
        '
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox18)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox11)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox12)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox13)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox14)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox15)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox16)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox17)
        Me.Panel_WhiteOnBlack.Controls.Add(Me.PictureBox19)
        Me.Panel_WhiteOnBlack.Location = New System.Drawing.Point(-6, -5)
        Me.Panel_WhiteOnBlack.Name = "Panel_WhiteOnBlack"
        Me.Panel_WhiteOnBlack.Size = New System.Drawing.Size(198, 200)
        Me.Panel_WhiteOnBlack.TabIndex = 11
        '
        'PictureBox18
        '
        Me.PictureBox18.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox18.Enabled = False
        Me.PictureBox18.Image = CType(resources.GetObject("PictureBox18.Image"), System.Drawing.Image)
        Me.PictureBox18.Location = New System.Drawing.Point(72, 136)
        Me.PictureBox18.Name = "PictureBox18"
        Me.PictureBox18.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox18.TabIndex = 12
        Me.PictureBox18.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(8, 8)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox11.TabIndex = 2
        Me.PictureBox11.TabStop = False
        '
        'PictureBox12
        '
        Me.PictureBox12.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox12.Image = CType(resources.GetObject("PictureBox12.Image"), System.Drawing.Image)
        Me.PictureBox12.Location = New System.Drawing.Point(72, 8)
        Me.PictureBox12.Name = "PictureBox12"
        Me.PictureBox12.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox12.TabIndex = 10
        Me.PictureBox12.TabStop = False
        '
        'PictureBox13
        '
        Me.PictureBox13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox13.Image = CType(resources.GetObject("PictureBox13.Image"), System.Drawing.Image)
        Me.PictureBox13.Location = New System.Drawing.Point(136, 8)
        Me.PictureBox13.Name = "PictureBox13"
        Me.PictureBox13.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox13.TabIndex = 10
        Me.PictureBox13.TabStop = False
        '
        'PictureBox14
        '
        Me.PictureBox14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox14.Image = CType(resources.GetObject("PictureBox14.Image"), System.Drawing.Image)
        Me.PictureBox14.Location = New System.Drawing.Point(8, 72)
        Me.PictureBox14.Name = "PictureBox14"
        Me.PictureBox14.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox14.TabIndex = 10
        Me.PictureBox14.TabStop = False
        '
        'PictureBox15
        '
        Me.PictureBox15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox15.Image = CType(resources.GetObject("PictureBox15.Image"), System.Drawing.Image)
        Me.PictureBox15.Location = New System.Drawing.Point(72, 72)
        Me.PictureBox15.Name = "PictureBox15"
        Me.PictureBox15.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox15.TabIndex = 10
        Me.PictureBox15.TabStop = False
        '
        'PictureBox16
        '
        Me.PictureBox16.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox16.Image = CType(resources.GetObject("PictureBox16.Image"), System.Drawing.Image)
        Me.PictureBox16.Location = New System.Drawing.Point(136, 72)
        Me.PictureBox16.Name = "PictureBox16"
        Me.PictureBox16.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox16.TabIndex = 10
        Me.PictureBox16.TabStop = False
        '
        'PictureBox17
        '
        Me.PictureBox17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox17.Image = CType(resources.GetObject("PictureBox17.Image"), System.Drawing.Image)
        Me.PictureBox17.Location = New System.Drawing.Point(8, 136)
        Me.PictureBox17.Name = "PictureBox17"
        Me.PictureBox17.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox17.TabIndex = 11
        Me.PictureBox17.TabStop = False
        '
        'PictureBox19
        '
        Me.PictureBox19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox19.Enabled = False
        Me.PictureBox19.Image = CType(resources.GetObject("PictureBox19.Image"), System.Drawing.Image)
        Me.PictureBox19.Location = New System.Drawing.Point(136, 136)
        Me.PictureBox19.Name = "PictureBox19"
        Me.PictureBox19.Size = New System.Drawing.Size(60, 60)
        Me.PictureBox19.TabIndex = 12
        Me.PictureBox19.TabStop = False
        '
        'PictureBox10
        '
        Me.PictureBox10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(64, 16)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(144, 144)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox10.TabIndex = 11
        Me.PictureBox10.TabStop = False
        Me.PictureBox10.Tag = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox_Size)
        Me.GroupBox1.Controls.Add(Me.GroupBox3)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown1)
        Me.GroupBox1.Location = New System.Drawing.Point(488, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(584, 320)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Settings"
        '
        'GroupBox_Size
        '
        Me.GroupBox_Size.Controls.Add(Me.BrightnessValue)
        Me.GroupBox_Size.Controls.Add(Me.Label15)
        Me.GroupBox_Size.Controls.Add(Me.ValueThickness)
        Me.GroupBox_Size.Controls.Add(Me.Label6)
        Me.GroupBox_Size.Controls.Add(Me.ValueInDia)
        Me.GroupBox_Size.Controls.Add(Me.Label5)
        Me.GroupBox_Size.Controls.Add(Me.valuedetecteddiameter)
        Me.GroupBox_Size.Controls.Add(Me.ValueY)
        Me.GroupBox_Size.Controls.Add(Me.ValueX)
        Me.GroupBox_Size.Controls.Add(Me.Button_SettingOK)
        Me.GroupBox_Size.Controls.Add(Me.Label4)
        Me.GroupBox_Size.Controls.Add(Me.Label3)
        Me.GroupBox_Size.Controls.Add(Me.Label2)
        Me.GroupBox_Size.Location = New System.Drawing.Point(8, 16)
        Me.GroupBox_Size.Name = "GroupBox_Size"
        Me.GroupBox_Size.Size = New System.Drawing.Size(320, 296)
        Me.GroupBox_Size.TabIndex = 18
        Me.GroupBox_Size.TabStop = False
        Me.GroupBox_Size.Text = "Size"
        '
        'BrightnessValue
        '
        Me.BrightnessValue.Location = New System.Drawing.Point(96, 152)
        Me.BrightnessValue.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.BrightnessValue.Name = "BrightnessValue"
        Me.BrightnessValue.Size = New System.Drawing.Size(64, 20)
        Me.BrightnessValue.TabIndex = 16
        Me.BrightnessValue.Value = New Decimal(New Integer() {128, 0, 0, 0})
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(8, 152)
        Me.Label15.Name = "Label15"
        Me.Label15.TabIndex = 15
        Me.Label15.Text = "Brightness"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueThickness
        '
        Me.ValueThickness.DecimalPlaces = 2
        Me.ValueThickness.Enabled = False
        Me.ValueThickness.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueThickness.Location = New System.Drawing.Point(96, 128)
        Me.ValueThickness.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.ValueThickness.Name = "ValueThickness"
        Me.ValueThickness.Size = New System.Drawing.Size(64, 20)
        Me.ValueThickness.TabIndex = 14
        Me.ValueThickness.Value = New Decimal(New Integer() {2, 0, 0, 65536})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 128)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Thickness:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ValueInDia
        '
        Me.ValueInDia.DecimalPlaces = 2
        Me.ValueInDia.Enabled = False
        Me.ValueInDia.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueInDia.Location = New System.Drawing.Point(96, 40)
        Me.ValueInDia.Maximum = New Decimal(New Integer() {6, 0, 0, 0})
        Me.ValueInDia.Name = "ValueInDia"
        Me.ValueInDia.Size = New System.Drawing.Size(64, 20)
        Me.ValueInDia.TabIndex = 12
        Me.ValueInDia.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Inner Diameter:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'valuedetecteddiameter
        '
        Me.valuedetecteddiameter.DecimalPlaces = 2
        Me.valuedetecteddiameter.Enabled = False
        Me.valuedetecteddiameter.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.valuedetecteddiameter.Location = New System.Drawing.Point(96, 16)
        Me.valuedetecteddiameter.Maximum = New Decimal(New Integer() {6, 0, 0, 0})
        Me.valuedetecteddiameter.Name = "valuedetecteddiameter"
        Me.valuedetecteddiameter.Size = New System.Drawing.Size(64, 20)
        Me.valuedetecteddiameter.TabIndex = 9
        Me.valuedetecteddiameter.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'ValueY
        '
        Me.ValueY.DecimalPlaces = 2
        Me.ValueY.Enabled = False
        Me.ValueY.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueY.Location = New System.Drawing.Point(96, 96)
        Me.ValueY.Maximum = New Decimal(New Integer() {15, 0, 0, 0})
        Me.ValueY.Name = "ValueY"
        Me.ValueY.Size = New System.Drawing.Size(64, 20)
        Me.ValueY.TabIndex = 8
        Me.ValueY.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'ValueX
        '
        Me.ValueX.DecimalPlaces = 2
        Me.ValueX.Enabled = False
        Me.ValueX.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.ValueX.Location = New System.Drawing.Point(96, 72)
        Me.ValueX.Maximum = New Decimal(New Integer() {15, 0, 0, 0})
        Me.ValueX.Name = "ValueX"
        Me.ValueX.Size = New System.Drawing.Size(64, 20)
        Me.ValueX.TabIndex = 7
        Me.ValueX.Value = New Decimal(New Integer() {2, 0, 0, 0})
        '
        'Button_SettingOK
        '
        Me.Button_SettingOK.Location = New System.Drawing.Point(192, 16)
        Me.Button_SettingOK.Name = "Button_SettingOK"
        Me.Button_SettingOK.Size = New System.Drawing.Size(120, 272)
        Me.Button_SettingOK.TabIndex = 6
        Me.Button_SettingOK.Text = "Ok"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(72, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 23)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Y:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(72, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "X:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Outer Diameter:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button_Test)
        Me.GroupBox3.Controls.Add(Me.GroupBox5)
        Me.GroupBox3.Location = New System.Drawing.Point(336, 16)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(240, 296)
        Me.GroupBox3.TabIndex = 17
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Find Fiducial"
        '
        'Button_Test
        '
        Me.Button_Test.Enabled = False
        Me.Button_Test.Location = New System.Drawing.Point(56, 88)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(144, 56)
        Me.Button_Test.TabIndex = 14
        Me.Button_Test.Text = "Find"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label46)
        Me.GroupBox5.Controls.Add(Me.Label44)
        Me.GroupBox5.Controls.Add(Me.Label43)
        Me.GroupBox5.Controls.Add(Me.Label42)
        Me.GroupBox5.Controls.Add(Me.TextBox_RPosX)
        Me.GroupBox5.Controls.Add(Me.TextBox_RPosY)
        Me.GroupBox5.Controls.Add(Me.TextBox_Score)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.Label20)
        Me.GroupBox5.Controls.Add(Me.Label49)
        Me.GroupBox5.Location = New System.Drawing.Point(8, 168)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(224, 120)
        Me.GroupBox5.TabIndex = 73
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Results:"
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(64, 24)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(24, 24)
        Me.Label46.TabIndex = 66
        Me.Label46.Text = "X="
        Me.Label46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(64, 48)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(24, 24)
        Me.Label44.TabIndex = 67
        Me.Label44.Text = "Y="
        Me.Label44.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(184, 24)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(32, 16)
        Me.Label43.TabIndex = 68
        Me.Label43.Text = "mm"
        Me.Label43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(184, 48)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(32, 16)
        Me.Label42.TabIndex = 69
        Me.Label42.Text = "mm"
        Me.Label42.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_RPosX
        '
        Me.TextBox_RPosX.Location = New System.Drawing.Point(88, 24)
        Me.TextBox_RPosX.Name = "TextBox_RPosX"
        Me.TextBox_RPosX.ReadOnly = True
        Me.TextBox_RPosX.Size = New System.Drawing.Size(96, 20)
        Me.TextBox_RPosX.TabIndex = 70
        Me.TextBox_RPosX.Text = "0"
        Me.TextBox_RPosX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RPosY
        '
        Me.TextBox_RPosY.Location = New System.Drawing.Point(88, 48)
        Me.TextBox_RPosY.Name = "TextBox_RPosY"
        Me.TextBox_RPosY.ReadOnly = True
        Me.TextBox_RPosY.Size = New System.Drawing.Size(96, 20)
        Me.TextBox_RPosY.TabIndex = 71
        Me.TextBox_RPosY.Text = "0"
        Me.TextBox_RPosY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_Score
        '
        Me.TextBox_Score.Location = New System.Drawing.Point(88, 88)
        Me.TextBox_Score.Name = "TextBox_Score"
        Me.TextBox_Score.ReadOnly = True
        Me.TextBox_Score.Size = New System.Drawing.Size(96, 20)
        Me.TextBox_Score.TabIndex = 15
        Me.TextBox_Score.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 20)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Score:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(192, 88)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(24, 23)
        Me.Label20.TabIndex = 72
        Me.Label20.Text = "%"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(8, 24)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(56, 48)
        Me.Label49.TabIndex = 65
        Me.Label49.Text = "Offset from center:"
        Me.Label49.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(72, 256)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(24, 23)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Button1"
        Me.Button1.Visible = False
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(104, 256)
        Me.NumericUpDown1.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(64, 20)
        Me.NumericUpDown1.TabIndex = 0
        Me.NumericUpDown1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown1.Value = New Decimal(New Integer() {127, 0, 0, 0})
        Me.NumericUpDown1.Visible = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label11)
        Me.GroupBox6.Controls.Add(Me.Label12)
        Me.GroupBox6.Controls.Add(Me.TextBox_Dia)
        Me.GroupBox6.Controls.Add(Me.Label7)
        Me.GroupBox6.Controls.Add(Me.Label8)
        Me.GroupBox6.Controls.Add(Me.Label9)
        Me.GroupBox6.Controls.Add(Me.Label10)
        Me.GroupBox6.Controls.Add(Me.TextBox_X)
        Me.GroupBox6.Controls.Add(Me.TextBox_Y)
        Me.GroupBox6.Controls.Add(Me.Button_Diameter)
        Me.GroupBox6.Controls.Add(Me.Button_Distance)
        Me.GroupBox6.Controls.Add(Me.Label13)
        Me.GroupBox6.Controls.Add(Me.Label14)
        Me.GroupBox6.Controls.Add(Me.TextBox_Dist)
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(136, 184)
        Me.GroupBox6.TabIndex = 21
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Measurement Tools:"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 144)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(32, 24)
        Me.Label11.TabIndex = 78
        Me.Label11.Text = "Dia="
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(96, 144)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(32, 16)
        Me.Label12.TabIndex = 79
        Me.Label12.Text = "mm"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Dia
        '
        Me.TextBox_Dia.Location = New System.Drawing.Point(48, 144)
        Me.TextBox_Dia.Name = "TextBox_Dia"
        Me.TextBox_Dia.ReadOnly = True
        Me.TextBox_Dia.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_Dia.TabIndex = 80
        Me.TextBox_Dia.Text = "0"
        Me.TextBox_Dia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 72)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(24, 24)
        Me.Label7.TabIndex = 72
        Me.Label7.Text = "X="
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 96)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(24, 24)
        Me.Label8.TabIndex = 73
        Me.Label8.Text = "Y="
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(96, 72)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 16)
        Me.Label9.TabIndex = 74
        Me.Label9.Text = "mm"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(96, 96)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(32, 16)
        Me.Label10.TabIndex = 75
        Me.Label10.Text = "mm"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_X
        '
        Me.TextBox_X.Location = New System.Drawing.Point(48, 72)
        Me.TextBox_X.Name = "TextBox_X"
        Me.TextBox_X.ReadOnly = True
        Me.TextBox_X.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_X.TabIndex = 76
        Me.TextBox_X.Text = "0"
        Me.TextBox_X.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_Y
        '
        Me.TextBox_Y.Location = New System.Drawing.Point(48, 96)
        Me.TextBox_Y.Name = "TextBox_Y"
        Me.TextBox_Y.ReadOnly = True
        Me.TextBox_Y.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_Y.TabIndex = 77
        Me.TextBox_Y.Text = "0"
        Me.TextBox_Y.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_Diameter
        '
        Me.Button_Diameter.Location = New System.Drawing.Point(32, 40)
        Me.Button_Diameter.Name = "Button_Diameter"
        Me.Button_Diameter.TabIndex = 1
        Me.Button_Diameter.Text = "Diameter"
        '
        'Button_Distance
        '
        Me.Button_Distance.Location = New System.Drawing.Point(32, 16)
        Me.Button_Distance.Name = "Button_Distance"
        Me.Button_Distance.TabIndex = 0
        Me.Button_Distance.Text = "Distance"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 120)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(32, 24)
        Me.Label13.TabIndex = 78
        Me.Label13.Text = "Dis="
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(96, 120)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(32, 16)
        Me.Label14.TabIndex = 79
        Me.Label14.Text = "mm"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Dist
        '
        Me.TextBox_Dist.Location = New System.Drawing.Point(48, 120)
        Me.TextBox_Dist.Name = "TextBox_Dist"
        Me.TextBox_Dist.ReadOnly = True
        Me.TextBox_Dist.Size = New System.Drawing.Size(48, 20)
        Me.TextBox_Dist.TabIndex = 80
        Me.TextBox_Dist.Text = "0"
        Me.TextBox_Dist.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button_Teach
        '
        Me.Button_Teach.Enabled = False
        Me.Button_Teach.Location = New System.Drawing.Point(248, 24)
        Me.Button_Teach.Name = "Button_Teach"
        Me.Button_Teach.Size = New System.Drawing.Size(136, 72)
        Me.Button_Teach.TabIndex = 20
        Me.Button_Teach.Text = "Teach"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.PictureBox20)
        Me.GroupBox2.Location = New System.Drawing.Point(1080, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(184, 248)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Preview"
        '
        'PictureBox20
        '
        Me.PictureBox20.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.PictureBox20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox20.Location = New System.Drawing.Point(8, 16)
        Me.PictureBox20.Name = "PictureBox20"
        Me.PictureBox20.Size = New System.Drawing.Size(168, 224)
        Me.PictureBox20.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox20.TabIndex = 0
        Me.PictureBox20.TabStop = False
        '
        'Button_EndFi
        '
        Me.Button_EndFi.Location = New System.Drawing.Point(1080, 264)
        Me.Button_EndFi.Name = "Button_EndFi"
        Me.Button_EndFi.Size = New System.Drawing.Size(96, 32)
        Me.Button_EndFi.TabIndex = 1
        Me.Button_EndFi.Text = "End Fiducial"
        Me.Button_EndFi.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.tbModelInfo)
        Me.GroupBox4.Controls.Add(Me.Button_Load)
        Me.GroupBox4.Controls.Add(Me.TabControl1)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(472, 320)
        Me.GroupBox4.TabIndex = 17
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Teach Fiducial"
        '
        'tbModelInfo
        '
        Me.tbModelInfo.Location = New System.Drawing.Point(16, 288)
        Me.tbModelInfo.Name = "tbModelInfo"
        Me.tbModelInfo.ReadOnly = True
        Me.tbModelInfo.Size = New System.Drawing.Size(440, 20)
        Me.tbModelInfo.TabIndex = 17
        Me.tbModelInfo.Text = ""
        '
        'Button_Load
        '
        Me.Button_Load.Location = New System.Drawing.Point(16, 232)
        Me.Button_Load.Name = "Button_Load"
        Me.Button_Load.Size = New System.Drawing.Size(96, 48)
        Me.Button_Load.TabIndex = 13
        Me.Button_Load.Text = "Load Fiducial"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Location = New System.Drawing.Point(8, 16)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(440, 208)
        Me.TabControl1.TabIndex = 16
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.PictureBox10)
        Me.TabPage3.Controls.Add(Me.Button_Teach)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(432, 182)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Fiducial"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Panel1)
        Me.TabPage1.Controls.Add(Me.Panel_BlackOnWhite)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(344, 198)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Black on White"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.GroupBox6)
        Me.Panel1.Location = New System.Drawing.Point(200, 7)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(136, 193)
        Me.Panel1.TabIndex = 10
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Panel_WhiteOnBlack)
        Me.TabPage2.Controls.Add(Me.Panel2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(344, 198)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "White on Black"
        '
        'Panel2
        '
        Me.Panel2.Location = New System.Drawing.Point(200, 7)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(136, 193)
        Me.Panel2.TabIndex = 12
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Button_Fid_Ok)
        Me.GroupBox7.Controls.Add(Me.GroupBox4)
        Me.GroupBox7.Controls.Add(Me.GroupBox2)
        Me.GroupBox7.Controls.Add(Me.Button_Fid_Cancel)
        Me.GroupBox7.Controls.Add(Me.GroupBox1)
        Me.GroupBox7.Controls.Add(Me.Button_EndFi)
        Me.GroupBox7.Location = New System.Drawing.Point(0, -8)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(1280, 336)
        Me.GroupBox7.TabIndex = 20
        Me.GroupBox7.TabStop = False
        '
        'FiducialForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1272, 328)
        Me.Controls.Add(Me.GroupBox7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FiducialForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Fiducial Mark: First Fiducial"
        Me.TopMost = True
        Me.Panel_BlackOnWhite.ResumeLayout(False)
        Me.Panel_WhiteOnBlack.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox_Size.ResumeLayout(False)
        CType(Me.BrightnessValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueThickness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueInDia, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.valuedetecteddiameter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ValueX, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Shared m_instance As FiducialForm

    Public Shared ReadOnly Property Instance() As FiducialForm
        Get
            If m_instance Is Nothing Then
                m_instance = New FiducialForm
            Else
                If m_instance.IsDisposed Then
                    m_instance = New FiducialForm
                End If
            End If
            Return m_instance
        End Get

    End Property

#Region "Global Variables"
    Dim Ok_clicked As Boolean = False
    Dim Fiducial_no As Integer = 1
    Dim FID As Integer
    Dim SQImgNo As Integer
    Dim ImgNo As Integer
    Dim FileLoad As Integer = 0
    Dim clicked As Integer = 1
#End Region

    Function Fi_No(ByVal Num As Integer)
        Fiducial_no = Num
        FidFilename = Fiducial_no.ToString + ".mno"
    End Function
    Dim Status As Integer = 0 '0=yet done, 1=done, 2=cancel
    Dim FidFilename As String
    Function GetFiducialStatus() As Integer
        Return Status
    End Function
    Function GetFiducialFilename() As String
        Return FidFilename
    End Function
    Function SetFiducialFilename(ByVal filename As String)
        FidFilename = filename
    End Function
    Function ResetFiducialStatus()
        Status = 0
    End Function
    Private Sub Button_Fid_Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Fid_Ok.Click
        tbModelInfo.Text = ""
        Button_Load.Enabled = True
        If FID > 0 Or isEdit Then
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                'To save the seach region if editing the fiducial
                If isEdit Then
                    FrmVision.ClickPM = 1
                    FrmVision.CustomizeFiducial_PM(False)
                    FrmVision.SaveSearchRegion(Fiducial_no, FID)
                    'End save search region
                End If
                If TeachClick Then
                    FrmVision.SaveFiducial(Fiducial_no, FID)
                End If
                FidFilename = Fiducial_no.ToString + ".mno"
                'Fiducial_no = Fiducial_no + 1
                Me.Visible = False
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
            'To save the seach region if editing the fiducial
                If isEdit Then
                    FrmVision.ClickPM = 1
                    FrmVision.CustomizeFiducial_PM(False)
                    FrmVision.SaveSearchRegion(Fiducial_no, FID)
                    'End save search region
                End If
                If TeachClick Then
                    FrmVision.SaveFiducial(Fiducial_no, FID)
                End If
                FidFilename = Fiducial_no.ToString + ".mno"
                PictureBoxUp()
                'FrmVision.AxPatternMatching2.Models.Item(1).Save("C:\IDS\DLL Export Device Vision\Fiducial\2.mmo")
                Me.Visible = False
                'Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
            Status = 1
        Else
        If Not isEdit Then
            MsgBox("Please select a fiducial mark~!")
        Else
            RemoveImage20()
            Me.Visible = False
        End If
        End If
        TeachClick = False
        isEdit = False
        FrmVision.ClickPM = 0
        FrmVision.isEdit = False
        FiducialMark_form.Button_Teach.Enabled = False
        If Not (FormCloseEvent Is Nothing) Then
            Status = 1
            FormCloseEvent()
        End If
    End Sub
    Private Sub Button_Fid_Ok1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Fiducial black on white
        If PictureBox1.BorderStyle = BorderStyle.Fixed3D Then
            'circle
            'Ok_clicked = True
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                FrmVision.SaveFiducial(Fiducial_no, FID)
                Fiducial_no = Fiducial_no + 1
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
                FrmVision.SaveFiducial(Fiducial_no, FID)
                PictureBoxUp()
                'FrmVision.AxPatternMatching2.Models.Item(1).Save("C:\IDS\DLL Export Device Vision\Fiducial\2.mmo")
                Me.Visible = False
                Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
        ElseIf PictureBox2.BorderStyle = BorderStyle.Fixed3D Then
            'circle
            'Ok_clicked = True
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                FrmVision.SaveFiducial(Fiducial_no, FID)
                Fiducial_no = Fiducial_no + 1
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
                FrmVision.SaveFiducial(Fiducial_no, FID)
                PictureBoxUp()
                'FrmVision.AxPatternMatching2.Models.Item(1).Save("C:\IDS\DLL Export Device Vision\Fiducial\2.mmo")
                Me.Visible = False
                Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
        ElseIf PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox8.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox9.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
            'Fiducial white on black
        ElseIf PictureBox11.BorderStyle = BorderStyle.Fixed3D Then
            'circle
            Me.Visible = False
        ElseIf PictureBox12.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox18.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False
        ElseIf PictureBox19.BorderStyle = BorderStyle.Fixed3D Then
            Me.Visible = False

        ElseIf PictureBox10.BorderStyle = BorderStyle.Fixed3D Then 'custom fiducial
            FrmVision.ClearDisplay()
            FrmVision.DisplayIndicator()
            'Ok_clicked = True
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                FrmVision.SaveFiducial(Fiducial_no, FID)
                Fiducial_no = Fiducial_no + 1
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
                FrmVision.SaveFiducial(Fiducial_no, FID)
                PictureBoxUp()
                Me.Visible = False
                Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
        ElseIf FileLoad = 1 Then
            FrmVision.ClearDisplay()
            FrmVision.DisplayIndicator()
            'Ok_clicked = True
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                FrmVision.SaveFiducial(Fiducial_no, FID)
                Fiducial_no = Fiducial_no + 1
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
                FrmVision.SaveFiducial(Fiducial_no, FID)
                PictureBoxUp()
                Me.Visible = False
                Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
            FileLoad = 0
        Else

            MsgBox("Please select a fiducial mark~!")
        End If

    End Sub ' unwanted if the above click is ok.
    Private Sub Button_Fid_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Fid_Cancel.Click
        tbModelInfo.Text = ""
        Button_Load.Enabled = True
        PictureBoxUp()
        FrmVision.ClearDisplay()
        FrmVision.DisplayIndicator()
        Me.Visible = False
        'Fiducial_no = 1
        'Me.Text = "Fiducial Mark: First Fiducial"
        TabControl1.SelectedIndex = 0
        Button_Fid_Ok.Enabled = False
        Button_EndFi.Enabled = False
        FrmVision.CustomizeFiducial_PM(True)
        Status = 2
        TeachClick = False
        isEdit = False
        FrmVision.isEdit = False
        FiducialMark_form.Button_Teach.Enabled = False
        If Not (FormCloseEvent Is Nothing) Then
            FormCloseEvent()
        End If
    End Sub
    Private Sub FiducialMark_form_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        FrmVision.ClearDisplay()
        FrmVision.DisplayIndicator()
        'Panel2.Controls.Remove(FrmVision.PanelVision)
        If Ok_clicked = False Then
            PictureBoxUp()
        End If
    End Sub
    Private Sub FiducialMark_form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ValueX.Value = 2
        ValueY.Value = 2
        valuedetecteddiameter.Value = 2
        PictureBoxUp()
        Me.Text = "Fiducial Mark: First Fiducial"
        Button_Fid_Ok.Enabled = False
        If Fiducial_no = 2 Then
            Button_EndFi.Enabled = False
        ElseIf Fiducial_no = 1 Then
            Button_EndFi.Enabled = True
        End If
        TabPage1.Visible = False
        TabPage2.Visible = False
    End Sub
#Region "PictureBox"
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        If PictureBox1.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox1.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 1
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            valuedetecteddiameter.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        If PictureBox2.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox2.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 2
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            valuedetecteddiameter.Enabled = True
            ValueInDia.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        If PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox3.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 3
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueThickness.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        If PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox4.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 4
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        If PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox5.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 5
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            'ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        If PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox6.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 6
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            'ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        If PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox7.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White
            FID = 7
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        If PictureBox8.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox8.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.White

        End If
    End Sub
    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        If PictureBox9.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
            'FrmVision.Timer_ModelRegion.Stop()
            FrmVision.DisplayIndicator()
            Button_Teach.Enabled = False
        Else
            PictureBoxUp()
            PictureBox9.BorderStyle = BorderStyle.Fixed3D 'choosen
            PictureBox20.BackColor = Color.White
            FrmVision.modelregionDrawing()
            'FrmVision.Timer_ModelRegion.Start()
            'FrmVision.Button_Teach.Enabled = True
            'FrmVision.Button_Teach.Text = "Select the model"
            FID = 9

            Button_Teach.Enabled = True
            Button_Teach.Text = "Next"
        End If
    End Sub
    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        If PictureBox11.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox11.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 11
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            valuedetecteddiameter.Enabled = True
            Button_SettingOK.Enabled = True
        End If
    End Sub
    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        If PictureBox12.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox12.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 12
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            valuedetecteddiameter.Enabled = True
            ValueInDia.Enabled = True
            Button_SettingOK.Enabled = True
        End If
    End Sub
    Private Sub PictureBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click
        If PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox13.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 13
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueThickness.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        If PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox14.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 14
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        If PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox15.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 15
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            'ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        If PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox16.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 16
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            'ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox17.Click
        If PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox17.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 17
            FrmVision.SearchRegion()
            FrmVision.SearchRegionDrawing()
            ValueX.Enabled = True
            ValueY.Enabled = True
            Button_SettingOK.Enabled = True

        End If
    End Sub
    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        If PictureBox18.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox18.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black

        End If
    End Sub
    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        If PictureBox19.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
        Else
            PictureBoxUp()
            PictureBox19.BorderStyle = BorderStyle.Fixed3D
            PictureBox20.BackColor = Color.Black
            FID = 19

        End If
    End Sub
    Public Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        If PictureBox10.BorderStyle = BorderStyle.Fixed3D Then
            PictureBoxUp()
            'FrmVision.Timer_ModelRegion.Stop()
            FrmVision.DisplayIndicator()
            Button_Teach.Enabled = False
            Button_Load.Enabled = True
        Else
            PictureBoxUp()
            PictureBox10.BorderStyle = BorderStyle.Fixed3D 'choosen
            FrmVision.DisplayIndicator()
            FrmVision.modelregionDrawing()
            FrmVision.SearchRegionDrawing()
            Button_Load.Enabled = False
            'FrmVision.Timer_ModelRegion.Start()
            'FrmVision.Button_Teach.Enabled = True
            'FrmVision.Button_Teach.Text = "Select the model"

            FID = 10

            Button_Teach.Enabled = True
            'Button_Teach.Text = "Next"
            Button_Teach.Text = "Finish"

        End If
    End Sub
#End Region

    Private Sub PictureBoxUp()
        PictureBox1.BorderStyle = BorderStyle.FixedSingle
        PictureBox2.BorderStyle = BorderStyle.FixedSingle
        PictureBox3.BorderStyle = BorderStyle.FixedSingle
        PictureBox4.BorderStyle = BorderStyle.FixedSingle
        PictureBox5.BorderStyle = BorderStyle.FixedSingle
        PictureBox6.BorderStyle = BorderStyle.FixedSingle
        PictureBox7.BorderStyle = BorderStyle.FixedSingle
        PictureBox8.BorderStyle = BorderStyle.FixedSingle
        PictureBox9.BorderStyle = BorderStyle.FixedSingle

        PictureBox11.BorderStyle = BorderStyle.FixedSingle
        PictureBox12.BorderStyle = BorderStyle.FixedSingle
        PictureBox13.BorderStyle = BorderStyle.FixedSingle
        PictureBox14.BorderStyle = BorderStyle.FixedSingle
        PictureBox15.BorderStyle = BorderStyle.FixedSingle
        PictureBox16.BorderStyle = BorderStyle.FixedSingle
        PictureBox17.BorderStyle = BorderStyle.FixedSingle
        PictureBox18.BorderStyle = BorderStyle.FixedSingle
        PictureBox19.BorderStyle = BorderStyle.FixedSingle

        PictureBox10.BorderStyle = BorderStyle.FixedSingle
        FileLoad = 0
        FID = 0
        valuedetecteddiameter.Enabled = False
        ValueX.Enabled = False
        ValueY.Enabled = False
        ValueThickness.Enabled = False
        ValueInDia.Enabled = False
        ValueX.Value = 2
        ValueY.Value = 2
        valuedetecteddiameter.Value = 2

        RemoveImage20()

        clicked = 1
        FrmVision.DisplayIndicator()
        FrmVision.Reset()
        Button_Teach.Enabled = False
        Button_Test.Enabled = False
        Button_SettingOK.Enabled = False
        Button_Fid_Ok.Enabled = False
        Button_EndFi.Enabled = False
        FrmVision.CustomizeFiducial_PM(True)
    End Sub
    Sub RemoveImage20()
        If PictureBox20.Image Is Nothing Then
        Else
            'Dim image As Image = PictureBox20.Image
            PictureBox20.Image.Dispose()
            PictureBox20.Image = Nothing
            'image.Dispose()
        End If
    End Sub
    Private Sub Button_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Test.Click
        Try
            FrmVision.DisplayIndicator()
            'If FID = 10 Then

            'ElseIf FID = 20 Then

            'Else
            '    FrmVision.PatternMatching_settings()
            'End If
            FrmVision.Test_Fiducial(FID, 0, 0)
            FrmVision.SearchRegionDrawing(False)
            TextBox_Score.Text = FrmVision.Score(1)
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
    Private Sub Button_SettingOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_SettingOK.Click
        Button_Test.Enabled = True
        Button_Fid_Ok.Enabled = True
        If Fiducial_no = 2 Then
            Button_EndFi.Enabled = False
        Else
            Button_EndFi.Enabled = True
        End If
    End Sub
    Dim TeachClick As Boolean = False
    Private Sub Button_Teach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Teach.Click
        If PictureBox10.BorderStyle = BorderStyle.Fixed3D Then
            FrmVision.CustomizeFiducial_PM(False) 'for model region
        ElseIf PictureBox9.BorderStyle = BorderStyle.Fixed3D Then

        End If
        'FrmVision.SearchRegion()
        'FrmVision.SearchRegionDrawing()
        If PictureBox10.BorderStyle = BorderStyle.Fixed3D Then
            FrmVision.CustomizeFiducial_PM(False) 'for search region
        ElseIf PictureBox9.BorderStyle = BorderStyle.Fixed3D Then

        End If
        Button_Teach.Enabled = False
        Button_Test.Enabled = True
        Button_Fid_Ok.Enabled = True
        If Fiducial_no = 2 Then
            Button_EndFi.Enabled = False
        Else
            Button_EndFi.Enabled = True
        End If
        tbModelInfo.Text = ""
        Button_Load.Enabled = True
        TeachClick = True
    End Sub
    Private Sub Button_Load_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Load.Click
        'If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
        'PictureBox20.Image = New Bitmap(OpenFileDialog1.FileName)
        'Me.Width = PictureBox20.Image.Width + Me.Width - Me.ClientSize.Width
        'Me.Height = PictureBox20.Image.Height + Me.Height - Me.ClientSize.Height
        'Dim file_name As String = OpenFileDialog1.FileName
        'file_name = file_name.Substring(file_name.LastIndexOf("\") + 1)
        'Me.Text = "[" & file_name & "]"
        'SaveFileDialog_Image.FileName = OpenFileDialog_Image.FileName
        'End If
        PictureBoxUp()
        FileLoad = 1
        Dim Bool As Boolean = FrmVision.LoadFiducial()
        If Bool = True Then
            FrmVision.SearchRegionDrawing(False)
            FID = 20
            Button_Test.Enabled = True
            Button_Fid_Ok.Enabled = True
            If Fiducial_no = 2 Then
                Button_EndFi.Enabled = False
            Else
                Button_EndFi.Enabled = True
            End If
        End If
    End Sub
    Private Sub ValueX_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueX.ValueChanged
        If SQImgNo <> 0 Then
            RemoveImage20()
        End If
        If PictureBox2.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRing(ValueX.Value(), ValueInDia.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        End If
    End Sub
    Private Sub ValueY_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueY.ValueChanged
        If SQImgNo <> 0 Then
            RemoveImage20()
        End If
        If PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        End If
    End Sub
    Private Sub ValueThickness_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueThickness.ValueChanged
        If SQImgNo <> 0 Then
            RemoveImage20()
        End If
        If PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        End If
    End Sub
    Private Sub valuedetecteddiameter_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles valuedetecteddiameter.ValueChanged
        Dim Fdiameter As Double
        Try
            Fdiameter = valuedetecteddiameter.Value
            If (Fdiameter > 6) Or (0 < Fdiameter And Fdiameter < 0.2) Then
                If Fdiameter < 0.2 Then
                    valuedetecteddiameter.Value = 0.2
                ElseIf Fdiameter > 6 Then
                    valuedetecteddiameter.Value = 6
                End If
                'MsgBox("Please re-enter diameter. Ranging from 0.2-6 mm")
            ElseIf Fdiameter = 0 Then
            Else
                RemoveImage20()
                If PictureBox1.BorderStyle = BorderStyle.Fixed3D Or PictureBox11.BorderStyle = BorderStyle.Fixed3D Then
                    Dim Success As Boolean = FrmVision.FCircle(Fdiameter)
                    If Success = True Then
                        PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
                        PictureBox20.Image = PictureBox20.Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
                    End If
                ElseIf PictureBox2.BorderStyle = BorderStyle.Fixed3D Or PictureBox12.BorderStyle = BorderStyle.Fixed3D Then
                    If valuedetecteddiameter.Value < ValueInDia.Value Then
                        ValueInDia.Value = valuedetecteddiameter.Value '- 0.01
                    End If
                    Dim Success As Boolean = FrmVision.FRing(valuedetecteddiameter.Value, ValueInDia.Value)
                    If Success = True Then
                        PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
                        PictureBox20.Image = PictureBox20.Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
                    End If
                End If
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
    Private Sub ValueInDia_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ValueInDia.ValueChanged
        Dim Fdiameter As Double
        Try
            Fdiameter = valuedetecteddiameter.Value
            If (Fdiameter > 6) Or (0 < Fdiameter And Fdiameter < 0.2) Then
                If Fdiameter < 0.2 Then
                    valuedetecteddiameter.Value = 0.2
                ElseIf Fdiameter > 6 Then
                    valuedetecteddiameter.Value = 6
                End If
                'MsgBox("Please re-enter diameter. Ranging from 0.2-6 mm")
            ElseIf Fdiameter = 0 Then

            Else
                RemoveImage20()

                If PictureBox1.BorderStyle = BorderStyle.Fixed3D Or PictureBox11.BorderStyle = BorderStyle.Fixed3D Then
                    ImgNo = FrmVision.FCircle(Fdiameter)
                    PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
                    PictureBox20.Image = PictureBox20.Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\Circle" + ImgNo.ToString + ".bmp")
                ElseIf PictureBox2.BorderStyle = BorderStyle.Fixed3D Or PictureBox12.BorderStyle = BorderStyle.Fixed3D Then
                    If valuedetecteddiameter.Value < ValueInDia.Value Then
                        ValueInDia.Value = valuedetecteddiameter.Value '- 0.01
                    End If
                    Dim Success As Boolean = FrmVision.FRing(valuedetecteddiameter.Value, ValueInDia.Value)
                    If Success = True Then
                        PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
                        PictureBox20.Image = PictureBox20.Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\" + "temp.bmp")
                    End If
                End If
            End If
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

#Region "Production"
    Dim FI_OffX, FI_OffY As Double

    Function FiducialOutput(ByVal _del_x, ByVal _del_y)
        'MsgBox(_del_x & "     " & _del_y)
    End Function
#End Region
#Region "Unwanted Buttons"
    Dim start As Boolean = True
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If start = True Then
            Timer1.Start()
            start = False
        Else
            Timer1.Stop()
            start = True
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            FrmVision.Test_Fiducial(FID, 0, 0)
            TextBox_Score.Text = FrmVision.Score(1)
        Catch ex As Exception

            ExceptionDisplay(ex)
        End Try
    End Sub
    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        FrmVision.BinarizeThreshold(NumericUpDown1.Value)
    End Sub
#End Region

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 1 Then
            Panel_BlackOnWhite.BringToFront()
            Panel1.Controls.Add(GroupBox6)
            GroupBox6.Location = New Point(0, 0)
            PictureBoxUp()
        ElseIf TabControl1.SelectedIndex = 2 Then
            Panel_WhiteOnBlack.BringToFront()
            Panel2.Controls.Add(GroupBox6)
            GroupBox6.Location = New Point(0, 0)
            PictureBoxUp()
        ElseIf TabControl1.SelectedIndex = 0 Then
            PictureBoxUp()
        End If
    End Sub

    Private Sub ValueY_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueY.Enter
        If SQImgNo <> 0 Then
            RemoveImage20()
        End If
        If PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        End If
    End Sub
    Private Sub ValueX_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ValueX.Enter
        If SQImgNo <> 0 Then
            RemoveImage20()
        End If
        If PictureBox2.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRing(ValueX.Value(), ValueInDia.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox3.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox13.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FCross(ValueX.Value(), ValueThickness.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox4.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox14.BorderStyle = BorderStyle.Fixed3D Then
            SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox5.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox15.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FLFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox6.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox16.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FRFlag(ValueX.Value()) ', ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox7.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        ElseIf PictureBox17.BorderStyle = BorderStyle.Fixed3D Then
            'SQImgNo = FrmVision.FSquare(ValueX.Value(), ValueY.Value)
            SQImgNo = FrmVision.FDiamond(ValueX.Value(), ValueY.Value)
            PictureBox20.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox20.Image = Image.FromFile("C:\IDS\SRC\DLL Export Device Vision\Fiducial\sq" + SQImgNo.ToString + ".bmp")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_EndFi.Click
        Status = 3
        If FID > 0 Then
            If Fiducial_no = 1 Then
                Me.Text = "Fiducial Mark: Second Fiducial"
                '==Save Fiducial===
                FrmVision.SaveFiducial(Fiducial_no, FID)
                FidFilename = Fiducial_no.ToString + ".mno"
                'Fiducial_no = Fiducial_no + 1
                Me.Visible = False
                PictureBoxUp()
            ElseIf Fiducial_no = 2 Then
                FrmVision.SaveFiducial(Fiducial_no, FID)
                FidFilename = Fiducial_no.ToString + ".mno"
                PictureBoxUp()
                'FrmVision.AxPatternMatching2.Models.Item(1).Save("C:\IDS\DLL Export Device Vision\Fiducial\2.mmo")
                Me.Visible = False
                'Fiducial_no = 1
                Me.Text = "Fiducial Mark: First Fiducial"
            End If
        Else
            MsgBox("Please select a fiducial mark~!")
        End If
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrightnessValue.TextChanged
        If BrightnessValue.Focused = True Then   'Xu Long 20/02/2012
            FrmVision.SetBrightness(BrightnessValue.Value)
        End If
    End Sub
End Class
