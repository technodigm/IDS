Imports System.Math
Public Class GD
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        GDInitialization()
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
    Friend WithEvents AxDisplay1 As AxMIL.AxDisplay
    Friend WithEvents AxSystem1 As AxMIL.AxSystem
    Friend WithEvents AxImage1 As AxMIL.AxImage
    Friend WithEvents AxGraphicContext1 As AxMIL.AxGraphicContext
    Friend WithEvents AxGraphicContext2 As AxMIL.AxGraphicContext
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents AxGraphicContext3 As AxMIL.AxGraphicContext
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Button13 As System.Windows.Forms.Button
    Friend WithEvents Button14 As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown3 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents Button16 As System.Windows.Forms.Button
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As iViewCore.PictureBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Button18 As System.Windows.Forms.Button
    Friend WithEvents Button19 As System.Windows.Forms.Button
    Friend WithEvents Button20 As System.Windows.Forms.Button
    Friend WithEvents Button21 As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents NumericUpDown4 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown5 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button22 As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown6 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button24 As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As iViewCore.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(GD))
        Me.AxDisplay1 = New AxMIL.AxDisplay
        Me.AxSystem1 = New AxMIL.AxSystem
        Me.AxImage1 = New AxMIL.AxImage
        Me.AxGraphicContext1 = New AxMIL.AxGraphicContext
        Me.AxGraphicContext2 = New AxMIL.AxGraphicContext
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.AxGraphicContext3 = New AxMIL.AxGraphicContext
        Me.Button8 = New System.Windows.Forms.Button
        Me.Button9 = New System.Windows.Forms.Button
        Me.Button10 = New System.Windows.Forms.Button
        Me.Button11 = New System.Windows.Forms.Button
        Me.Button12 = New System.Windows.Forms.Button
        Me.Button13 = New System.Windows.Forms.Button
        Me.Button14 = New System.Windows.Forms.Button
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown3 = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button15 = New System.Windows.Forms.Button
        Me.Button16 = New System.Windows.Forms.Button
        Me.Button17 = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.PictureBox2 = New iViewCore.PictureBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.PictureBox1 = New iViewCore.PictureBox
        Me.Button18 = New System.Windows.Forms.Button
        Me.Button19 = New System.Windows.Forms.Button
        Me.Button20 = New System.Windows.Forms.Button
        Me.Button21 = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.NumericUpDown4 = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.NumericUpDown5 = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Button22 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.NumericUpDown6 = New System.Windows.Forms.NumericUpDown
        Me.Button24 = New System.Windows.Forms.Button
        CType(Me.AxDisplay1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxSystem1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxImage1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxGraphicContext3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxDisplay1
        '
        Me.AxDisplay1.ContainingControl = Me
        Me.AxDisplay1.Enabled = True
        Me.AxDisplay1.Location = New System.Drawing.Point(0, 0)
        Me.AxDisplay1.Name = "AxDisplay1"
        Me.AxDisplay1.OcxState = CType(resources.GetObject("AxDisplay1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDisplay1.Size = New System.Drawing.Size(600, 600)
        Me.AxDisplay1.TabIndex = 0
        '
        'AxSystem1
        '
        Me.AxSystem1.Enabled = True
        Me.AxSystem1.Location = New System.Drawing.Point(8, 8)
        Me.AxSystem1.Name = "AxSystem1"
        Me.AxSystem1.OcxState = CType(resources.GetObject("AxSystem1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxSystem1.Size = New System.Drawing.Size(32, 32)
        Me.AxSystem1.TabIndex = 1
        '
        'AxImage1
        '
        Me.AxImage1.Enabled = True
        Me.AxImage1.Location = New System.Drawing.Point(40, 8)
        Me.AxImage1.Name = "AxImage1"
        Me.AxImage1.OcxState = CType(resources.GetObject("AxImage1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxImage1.Size = New System.Drawing.Size(32, 32)
        Me.AxImage1.TabIndex = 2
        '
        'AxGraphicContext1
        '
        Me.AxGraphicContext1.Enabled = True
        Me.AxGraphicContext1.Location = New System.Drawing.Point(144, 8)
        Me.AxGraphicContext1.Name = "AxGraphicContext1"
        Me.AxGraphicContext1.OcxState = CType(resources.GetObject("AxGraphicContext1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext1.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext1.TabIndex = 3
        '
        'AxGraphicContext2
        '
        Me.AxGraphicContext2.Enabled = True
        Me.AxGraphicContext2.Location = New System.Drawing.Point(104, 8)
        Me.AxGraphicContext2.Name = "AxGraphicContext2"
        Me.AxGraphicContext2.OcxState = CType(resources.GetObject("AxGraphicContext2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext2.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext2.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(0, 440)
        Me.Button1.Name = "Button1"
        Me.Button1.TabIndex = 5
        Me.Button1.Text = "Clear"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(0, 48)
        Me.Button2.Name = "Button2"
        Me.Button2.TabIndex = 6
        Me.Button2.Text = "Line"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(0, 96)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 7
        Me.Button3.Text = "Circle"
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(0, 120)
        Me.Button4.Name = "Button4"
        Me.Button4.TabIndex = 8
        Me.Button4.Text = "CircleClear"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(0, 504)
        Me.Button5.Name = "Button5"
        Me.Button5.TabIndex = 9
        Me.Button5.Text = "Save"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(0, 472)
        Me.Button6.Name = "Button6"
        Me.Button6.TabIndex = 10
        Me.Button6.Text = "Load"
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(0, 72)
        Me.Button7.Name = "Button7"
        Me.Button7.TabIndex = 11
        Me.Button7.Text = "LineClear"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 344)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Label1"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(120, 344)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Label2"
        '
        'AxGraphicContext3
        '
        Me.AxGraphicContext3.Enabled = True
        Me.AxGraphicContext3.Location = New System.Drawing.Point(72, 8)
        Me.AxGraphicContext3.Name = "AxGraphicContext3"
        Me.AxGraphicContext3.OcxState = CType(resources.GetObject("AxGraphicContext3.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxGraphicContext3.Size = New System.Drawing.Size(32, 32)
        Me.AxGraphicContext3.TabIndex = 15
        '
        'Button8
        '
        Me.Button8.Location = New System.Drawing.Point(0, 144)
        Me.Button8.Name = "Button8"
        Me.Button8.TabIndex = 16
        Me.Button8.Text = "Dot"
        '
        'Button9
        '
        Me.Button9.Location = New System.Drawing.Point(0, 168)
        Me.Button9.Name = "Button9"
        Me.Button9.TabIndex = 17
        Me.Button9.Text = "DotClear"
        '
        'Button10
        '
        Me.Button10.Location = New System.Drawing.Point(0, 192)
        Me.Button10.Name = "Button10"
        Me.Button10.TabIndex = 18
        Me.Button10.Text = "Rectangle"
        '
        'Button11
        '
        Me.Button11.Location = New System.Drawing.Point(0, 216)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(75, 32)
        Me.Button11.TabIndex = 19
        Me.Button11.Text = "RectangleClear"
        '
        'Button12
        '
        Me.Button12.Location = New System.Drawing.Point(0, 248)
        Me.Button12.Name = "Button12"
        Me.Button12.TabIndex = 20
        Me.Button12.Text = "Arc"
        '
        'Button13
        '
        Me.Button13.Location = New System.Drawing.Point(0, 272)
        Me.Button13.Name = "Button13"
        Me.Button13.TabIndex = 21
        Me.Button13.Text = "ArcClear"
        '
        'Button14
        '
        Me.Button14.Location = New System.Drawing.Point(0, 296)
        Me.Button14.Name = "Button14"
        Me.Button14.TabIndex = 22
        Me.Button14.Text = "SpiralCircle"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.DecimalPlaces = 2
        Me.NumericUpDown1.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NumericUpDown1.Location = New System.Drawing.Point(16, 808)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(56, 20)
        Me.NumericUpDown1.TabIndex = 23
        Me.NumericUpDown1.Value = New Decimal(New Integer() {5, 0, 0, 65536})
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.DecimalPlaces = 2
        Me.NumericUpDown2.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NumericUpDown2.Location = New System.Drawing.Point(72, 24)
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.TabIndex = 24
        Me.NumericUpDown2.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'NumericUpDown3
        '
        Me.NumericUpDown3.DecimalPlaces = 2
        Me.NumericUpDown3.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NumericUpDown3.Location = New System.Drawing.Point(72, 48)
        Me.NumericUpDown3.Name = "NumericUpDown3"
        Me.NumericUpDown3.TabIndex = 25
        Me.NumericUpDown3.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 26
        Me.Label3.Text = "NoOfRound"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Pitch"
        '
        'Button15
        '
        Me.Button15.Location = New System.Drawing.Point(8, 776)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(64, 30)
        Me.Button15.TabIndex = 28
        Me.Button15.Text = "Button15"
        '
        'Button16
        '
        Me.Button16.Image = CType(resources.GetObject("Button16.Image"), System.Drawing.Image)
        Me.Button16.ImageAlign = System.Drawing.ContentAlignment.BottomRight
        Me.Button16.Location = New System.Drawing.Point(264, 320)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(35, 35)
        Me.Button16.TabIndex = 29
        '
        'Button17
        '
        Me.Button17.Image = CType(resources.GetObject("Button17.Image"), System.Drawing.Image)
        Me.Button17.ImageAlign = System.Drawing.ContentAlignment.BottomRight
        Me.Button17.Location = New System.Drawing.Point(296, 320)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(35, 35)
        Me.Button17.TabIndex = 30
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.Button16)
        Me.Panel1.Controls.Add(Me.Button17)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.AxDisplay1)
        Me.Panel1.Location = New System.Drawing.Point(80, 48)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(356, 378)
        Me.Panel1.TabIndex = 31
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Gray
        Me.PictureBox2.Image = Nothing
        Me.PictureBox2.Location = New System.Drawing.Point(32, 16)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(300, 300)
        Me.PictureBox2.TabIndex = 4
        Me.PictureBox2.ViewMode = iViewCore.PictureBox.EViewMode.FitImage
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.PictureBox1)
        Me.Panel2.Location = New System.Drawing.Point(440, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1200, 1200)
        Me.Panel2.TabIndex = 32
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Gray
        Me.PictureBox1.Image = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(96, 56)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(600, 600)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.ViewMode = iViewCore.PictureBox.EViewMode.FitImage
        '
        'Button18
        '
        Me.Button18.Location = New System.Drawing.Point(0, 728)
        Me.Button18.Name = "Button18"
        Me.Button18.TabIndex = 33
        Me.Button18.Text = "Button18"
        '
        'Button19
        '
        Me.Button19.Location = New System.Drawing.Point(0, 752)
        Me.Button19.Name = "Button19"
        Me.Button19.TabIndex = 34
        Me.Button19.Text = "Button19"
        '
        'Button20
        '
        Me.Button20.Location = New System.Drawing.Point(0, 320)
        Me.Button20.Name = "Button20"
        Me.Button20.Size = New System.Drawing.Size(75, 32)
        Me.Button20.TabIndex = 35
        Me.Button20.Text = "SpiralCircleClear"
        '
        'Button21
        '
        Me.Button21.Location = New System.Drawing.Point(0, 352)
        Me.Button21.Name = "Button21"
        Me.Button21.TabIndex = 36
        Me.Button21.Text = "ChipEdge"
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"Dot", "One edge", "Two edges"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 24)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 37
        Me.ComboBox1.Text = "Dot"
        '
        'NumericUpDown4
        '
        Me.NumericUpDown4.Location = New System.Drawing.Point(80, 48)
        Me.NumericUpDown4.Maximum = New Decimal(New Integer() {4, 0, 0, 0})
        Me.NumericUpDown4.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown4.Name = "NumericUpDown4"
        Me.NumericUpDown4.Size = New System.Drawing.Size(56, 20)
        Me.NumericUpDown4.TabIndex = 38
        Me.NumericUpDown4.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "MainEdge"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.NumericUpDown5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(80, 520)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(200, 144)
        Me.GroupBox1.TabIndex = 40
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "ChipEdge"
        '
        'NumericUpDown5
        '
        Me.NumericUpDown5.Location = New System.Drawing.Point(104, 112)
        Me.NumericUpDown5.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown5.Name = "NumericUpDown5"
        Me.NumericUpDown5.Size = New System.Drawing.Size(56, 20)
        Me.NumericUpDown5.TabIndex = 43
        Me.NumericUpDown5.Value = New Decimal(New Integer() {20, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 42
        Me.Label6.Text = "EdgeClearance"
        '
        'CheckBox1
        '
        Me.CheckBox1.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox1.Location = New System.Drawing.Point(16, 80)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.TabIndex = 41
        Me.CheckBox1.Text = "Cw"
        '
        'Button22
        '
        Me.Button22.BackColor = System.Drawing.SystemColors.Control
        Me.Button22.Location = New System.Drawing.Point(0, 376)
        Me.Button22.Name = "Button22"
        Me.Button22.Size = New System.Drawing.Size(75, 32)
        Me.Button22.TabIndex = 41
        Me.Button22.Text = "ChipEdgeClear"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox2.Controls.Add(Me.NumericUpDown3)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Location = New System.Drawing.Point(80, 440)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 80)
        Me.GroupBox2.TabIndex = 42
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Spiral"
        '
        'NumericUpDown6
        '
        Me.NumericUpDown6.Location = New System.Drawing.Point(16, 544)
        Me.NumericUpDown6.Maximum = New Decimal(New Integer() {20, 0, 0, 0})
        Me.NumericUpDown6.Name = "NumericUpDown6"
        Me.NumericUpDown6.Size = New System.Drawing.Size(40, 20)
        Me.NumericUpDown6.TabIndex = 43
        '
        'Button24
        '
        Me.Button24.Location = New System.Drawing.Point(0, 608)
        Me.Button24.Name = "Button24"
        Me.Button24.TabIndex = 45
        Me.Button24.Text = "Load"
        '
        'GD
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1272, 1009)
        Me.Controls.Add(Me.Button24)
        Me.Controls.Add(Me.NumericUpDown6)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Button22)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button21)
        Me.Controls.Add(Me.Button20)
        Me.Controls.Add(Me.Button19)
        Me.Controls.Add(Me.Button18)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.NumericUpDown1)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.AxGraphicContext3)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.AxGraphicContext2)
        Me.Controls.Add(Me.AxGraphicContext1)
        Me.Controls.Add(Me.AxImage1)
        Me.Controls.Add(Me.AxSystem1)
        Me.Controls.Add(Me.Button15)
        Me.Controls.Add(Me.Panel2)
        Me.Name = "GD"
        Me.Text = "GD"
        CType(Me.AxDisplay1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxSystem1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxImage1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxGraphicContext3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.NumericUpDown4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.NumericUpDown6, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Global Variables"
    Dim ImagePointer_Picturebox2 As New IntPtr
    Dim clickno = 1
    Dim a1, a2, a3, a4, b1, b2, b3, b4 As Double
    Dim myPen As System.Drawing.Pen
    Dim Ratio As Double = 1
    Private MyImage As Bitmap
    Private MyImage1 As Bitmap
    Dim PanelPositionX As Double
    Dim PanelPositionY As Double
    Dim RobotRefX, RobotRefY As Double
    Dim Level As Integer = 100
    Dim PFileName As String = "C:\Documents and Settings\student\Desktop\nan.bmp"
#End Region
    Sub GDInitialization()
        If AxSystem1.IsAllocated = True Then
            AxSystem1.Free()
        End If
        AxSystem1.AutomaticAllocation() = True
        AxImage1.AutomaticAllocation = True
        AxDisplay1.AutomaticAllocation = True
        AxSystem1.Allocate()
        AxDisplay1.OverlayWriting = True
        AxDisplay1.Visible = True
        AxDisplay1.OverlayFlicker = False
        AxDisplay1.OverlayKeyColor = System.Drawing.Color.Black
        AxDisplay1.OverlayKeyCondition = MIL.DispOverlayKeyConditionConstants.dispEqualToColor
        AxDisplay1.LUTEnabled = False
        AxDisplay1.OverlayLUT.GenerateRamp(0, 0, 255, 255) 'for display after using overlay
        AxDisplay1.OverlayImage.Clear()
    End Sub
    Sub ClearMILDisplay()
        AxDisplay1.OverlayImage.Clear()
    End Sub
#Region "Model"
    Function Circle(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal Draw As Boolean)
        Dim a, b, c, x, y, radius, x12, y12, x13, y13, m12, m13, m21, m31, c12, c13 As Double
        Dim ynew As Integer
        Dim xnew As Integer
        'Dim pen1 As New Pen(Color.Blue, 5)
        'g = PanelVision.CreateGraphics

        'If (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1) < 2 ^ 30 & (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1) < 2 ^ 30 & ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1)) < 2 ^ 30 Then
        If (y1 <> y2) Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1)
            c = ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x1 - x3) / (y3 - y1)) * x + 0.5 * a
        ElseIf y1 = y2 Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y3 - y2)) * (x3 * x3 + y3 * y3 - x2 * x2 - y2 * y2)
            c = ((x2 - x3) / (y3 - y2)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x2 - x3) / (y3 - y2)) * x + 0.5 * b
        End If

        If (Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))) < (2 ^ 30) Then
            radius = Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))
            'FormlocationX = Form1.ActiveForm.Location.X
            'FormlocationY = Form1.ActiveForm.Location.Y
            xnew = Convert.ToInt64(x - radius) '- FormlocationX
            ynew = Convert.ToInt64(y - radius) '- FormlocationY
            'g.DrawArc(pen1, xnew, ynew, 2 * (Convert.ToInt16(radius)), 2 * (Convert.ToInt16(radius)), 0, 360)
            If Draw = True Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = (xnew) '- (2000 / 22)
                    .StartY() = (ynew) '- (2000 / 22)
                    .EndX() = (xnew + 2 * radius) '+ (2000 / 22)
                    .EndY() = (ynew + 2 * radius) '+ (2000 / 22)
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                AxGraphicContext3.Arc(0, 360, False)
            Else
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = (xnew) '- (2000 / 22)
                    .StartY() = (ynew) '- (2000 / 22)
                    .EndX() = (xnew + 2 * radius) '+ (2000 / 22)
                    .EndY() = (ynew + 2 * radius) '+ (2000 / 22)
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                AxGraphicContext2.Arc(0, 360, False)
            End If
        End If
    End Function
    Function Line(ByVal x1 As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal Draw As Boolean)
        If Draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext3.LineSegment()
        Else
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext2.LineSegment()
        End If
    End Function
    Function Dot(ByVal x1 As Double, ByVal y1 As Double, ByVal Draw As Boolean)
        If Draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
            End With
            AxGraphicContext3.Dot()
        Else
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
            End With
            AxGraphicContext2.Dot()
        End If
    End Function
    Function Rectangle(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal Draw As Boolean)
        Dim x4 As Double
        If x2 < x1 Then
            x4 = x3 + Abs(x1 - x2)
        Else
            x4 = x3 - Abs(x1 - x2)
        End If

        Dim y4 As Double
        If y2 < y3 Then
            y4 = (y1 + Abs(y2 - y3))
        Else
            y4 = (y1 - Abs(y2 - y3))
        End If

        If Draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext3.LineSegment()
        Else
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext2.LineSegment()
        End If
    End Function
    Function Arc(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal Draw As Boolean)
        Dim a, b, c, x, y, radius, x12, y12, x13, y13, m12, m13, m21, m31, c12, c13 As Double
        Dim ynew As Integer
        Dim xnew As Integer
        'Dim pen1 As New Pen(Color.Blue, 5)
        'g = PanelVision.CreateGraphics

        'If (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1) < 2 ^ 30 & (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1) < 2 ^ 30 & ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1)) < 2 ^ 30 Then
        If (y1 <> y2) Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1)
            c = ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x1 - x3) / (y3 - y1)) * x + 0.5 * a
        ElseIf y1 = y2 Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y3 - y2)) * (x3 * x3 + y3 * y3 - x2 * x2 - y2 * y2)
            c = ((x2 - x3) / (y3 - y2)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x2 - x3) / (y3 - y2)) * x + 0.5 * b
        End If

        Dim MLine As Double = (y1 - y3) / (x1 - x3)
        Dim CLine As Double = y1 - MLine * x1
        Dim Forward As Boolean

        If x1 = x3 Then
            If y1 < y3 Then
                If x2 > x1 Then
                    '3-2-1
                    Forward = False
                Else
                    '1-2-3
                    Forward = True
                End If
            Else
                If x2 > x1 Then
                    '1-2-3
                    Forward = True
                Else
                    '3-2-1
                    Forward = False
                End If
            End If
        ElseIf x1 < x3 Then
            If y2 - (MLine * x2 + CLine) < 0 Then
                '3-2-1
                Forward = False
            Else
                '1-2-3
                Forward = True
            End If
        ElseIf x1 > x3 Then
            If y2 - (MLine * x2 + CLine) < 0 Then
                '1-2-3
                Forward = True
            Else
                '3-2-1
                Forward = False
            End If
        End If

        Dim angle1 As Double
        Dim angle3 As Double
        angle1 = -Atan((y1 - y) / (x1 - x)) / PI * 180 'deg
        angle3 = -Atan((y3 - y) / (x3 - x)) / PI * 180 'deg

        If x1 > x Then
            If angle1 > 0 Then
                'first quater
                angle1 = angle1
            Else
                'forth quater
                angle1 = 360 + angle1
            End If
        Else
            If angle1 > 0 Then
                'Third quater
                angle1 = angle1 + 180
            Else
                'Second Quater
                angle1 = 180 + angle1
            End If
        End If
        If x3 > x Then
            If angle3 > 0 Then
                'first quater
                angle3 = angle3
            Else
                'forth quater
                angle3 = 360 + angle3
            End If
        Else
            If angle3 > 0 Then
                'Third quater
                angle3 = angle3 + 180
            Else
                'Second Quater
                angle3 = 180 + angle3
            End If
        End If

        If (Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))) < (2 ^ 30) Then
            radius = Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))
            'FormlocationX = Form1.ActiveForm.Location.X
            'FormlocationY = Form1.ActiveForm.Location.Y
            xnew = Convert.ToInt64(x - radius) '- FormlocationX
            ynew = Convert.ToInt64(y - radius) '- FormlocationY
            'g.DrawArc(pen1, xnew, ynew, 2 * (Convert.ToInt16(radius)), 2 * (Convert.ToInt16(radius)), 0, 360)

            If Draw = True Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = (xnew) '- (2000 / 22)
                    .StartY() = (ynew) '- (2000 / 22)
                    .EndX() = (xnew + 2 * radius) '+ (2000 / 22)
                    .EndY() = (ynew + 2 * radius) '+ (2000 / 22)
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                If Forward = True Then
                    AxGraphicContext3.Arc(angle1, angle3, False)
                Else
                    AxGraphicContext3.Arc(angle3, angle1, False)
                End If

            Else
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = (xnew) '- (2000 / 22)
                    .StartY() = (ynew) '- (2000 / 22)
                    .EndX() = (xnew + 2 * radius) '+ (2000 / 22)
                    .EndY() = (ynew + 2 * radius) '+ (2000 / 22)
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                If Forward = True Then
                    AxGraphicContext2.Arc(angle1, angle3, False)
                Else
                    AxGraphicContext2.Arc(angle3, angle1, False)
                End If
            End If
        End If
    End Function
    Function CircleSpiral(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal Pitch As Double, ByVal NoOfRound As Integer, ByVal CW_CCW As Boolean, ByVal Draw As Boolean)
        Dim a, b, c, x, y, radius, x12, y12, x13, y13, m12, m13, m21, m31, c12, c13 As Double
        Dim ynew As Integer
        Dim xnew As Integer
        'Dim pen1 As New Pen(Color.Blue, 5)
        'g = PanelVision.CreateGraphics

        'If (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1) < 2 ^ 30 & (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1) < 2 ^ 30 & ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1)) < 2 ^ 30 Then
        If (y1 <> y2) Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y2 - y1)) * (x2 * x2 + y2 * y2 - x1 * x1 - y1 * y1)
            c = ((x1 - x2) / (y2 - y1)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x1 - x3) / (y3 - y1)) * x + 0.5 * a
        ElseIf y1 = y2 Then
            a = (1.0 / (y3 - y1)) * (x3 * x3 + y3 * y3 - x1 * x1 - y1 * y1)
            b = (1.0 / (y3 - y2)) * (x3 * x3 + y3 * y3 - x2 * x2 - y2 * y2)
            c = ((x2 - x3) / (y3 - y2)) - ((x1 - x3) / (y3 - y1))

            x = (0.5 * (a - b)) / c
            y = ((x2 - x3) / (y3 - y2)) * x + 0.5 * b
        End If

        Dim MLine As Double = (y1 - y3) / (x1 - x3)
        Dim CLine As Double = y1 - MLine * x1
        Dim Forward As Boolean

        If x1 = x3 Then
            If y1 < y3 Then
                If x2 > x1 Then
                    '3-2-1
                    Forward = False
                Else
                    '1-2-3
                    Forward = True
                End If
            Else
                If x2 > x1 Then
                    '1-2-3
                    Forward = True
                Else
                    '3-2-1
                    Forward = False
                End If
            End If
        ElseIf x1 < x3 Then
            If y2 - (MLine * x2 + CLine) < 0 Then
                '3-2-1
                Forward = False
            Else
                '1-2-3
                Forward = True
            End If
        ElseIf x1 > x3 Then
            If y2 - (MLine * x2 + CLine) < 0 Then
                '1-2-3
                Forward = True
            Else
                '3-2-1
                Forward = False
            End If
        End If

        Dim angle1 As Double
        Dim angle3 As Double

        angle1 = -Atan((y1 - y) / (x1 - x)) / PI * 180 'deg
        angle3 = -Atan((y3 - y) / (x3 - x)) / PI * 180 'deg

        If x1 > x Then
            If angle1 > 0 Then
                'first quater
                angle1 = angle1
            Else
                'forth quater
                angle1 = 360 + angle1
            End If
        Else
            If angle1 > 0 Then
                'Third quater
                angle1 = angle1 + 180
            Else
                'Second Quater
                angle1 = 180 + angle1
            End If
        End If
        If x3 > x Then
            If angle3 > 0 Then
                'first quater
                angle3 = angle3
            Else
                'forth quater
                angle3 = 360 + angle3
            End If
        Else
            If angle3 > 0 Then
                'Third quater
                angle3 = angle3 + 180
            Else
                'Second Quater
                angle3 = 180 + angle3
            End If
        End If

        If (Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))) < (2 ^ 30) Then
            radius = Sqrt((x - x1) * (x - x1) + (y - y1) * (y - y1))
            xnew = Convert.ToInt64(x - radius)
            ynew = Convert.ToInt64(y - radius)

            Dim AvailableRound As Integer = Fix(radius / Pitch)
            If NoOfRound < AvailableRound Then
                NoOfRound = NoOfRound
            Else
                NoOfRound = AvailableRound
            End If
            Dim angle_NoOfRound = (angle1 * PI / 180 + AvailableRound * 2 * PI) - (NoOfRound * 2 * PI) 'radian
            angle1 = angle1 * PI / 180 + AvailableRound * 2 * PI 'radian

            Dim Constant As Double
            Constant = radius / angle1

            If Draw = True Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x '- (Constant * angle1 * Cos(angle1))
                    .StartY() = y '- (Constant * angle1 * Sin(angle1))
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                AxGraphicContext3.Dot()
                If CW_CCW = False Then
                    For angle1 = angle1 To 0 Step -0.05
                        If angle1 < angle_NoOfRound Then
                        Else
                            If Pitch < (Constant * angle1) Then
                                'draw
                                With AxGraphicContext3.DrawingRegion()
                                    .StartX() = x - (Constant * angle1 * Cos(angle1))
                                    .StartY() = y - (Constant * angle1 * Sin(angle1))
                                End With
                                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                                AxGraphicContext3.Dot()

                            Else
                                'do nothing
                            End If

                        End If
                    Next angle1
                ElseIf CW_CCW = True Then
                    For angle1 = -angle1 To 0 Step 0.05
                        If angle1 > -angle_NoOfRound Then
                        Else
                            If Pitch < Abs(Constant * angle1) Then
                                'draw
                                With AxGraphicContext3.DrawingRegion()
                                    .StartX() = x - (Constant * angle1 * Cos(angle1))
                                    .StartY() = y - (Constant * angle1 * Sin(angle1))
                                End With
                                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                                AxGraphicContext3.Dot()
                            Else
                                'do nothing
                            End If
                        End If
                    Next angle1
                End If
            Else '==============================================================
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x '- (Constant * angle1 * Cos(angle1))
                    .StartY() = y '- (Constant * angle1 * Sin(angle1))
                End With
                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                AxGraphicContext2.Dot()
                If CW_CCW = False Then
                    For angle1 = angle1 To 0 Step -0.05
                        If angle1 < angle_NoOfRound Then
                        Else
                            If Pitch < (Constant * angle1) Then
                                'draw
                                With AxGraphicContext2.DrawingRegion()
                                    .StartX() = x - (Constant * angle1 * Cos(angle1))
                                    .StartY() = y - (Constant * angle1 * Sin(angle1))
                                End With
                                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                                AxGraphicContext2.Dot()
                            Else
                                'do nothing
                            End If

                        End If
                    Next angle1
                ElseIf CW_CCW = True Then
                    For angle1 = -angle1 To 0 Step 0.05
                        If angle1 > -angle_NoOfRound Then
                        Else
                            If Pitch < Abs(Constant * angle1) Then
                                'draw
                                With AxGraphicContext2.DrawingRegion()
                                    .StartX() = x - (Constant * angle1 * Cos(angle1))
                                    .StartY() = y - (Constant * angle1 * Sin(angle1))
                                End With
                                'AxGraphicContext1.BackgroundColor = AxGraphicContext1.BackgroundColor.White
                                'AxGraphicContext1.ForegroundColor = AxGraphicContext1.BackgroundColor.Black
                                AxGraphicContext2.Dot()
                            Else
                                'do nothing
                            End If
                        End If
                    Next angle1
                End If
            End If
        End If
    End Function
    Function ChipEdge(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal x4 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal y4 As Double, ByVal DispenseModel As Integer, ByVal Cw_CCW As Boolean, ByVal MainEdge As Integer, ByVal EdgeClearance As Double, ByVal draw As Boolean)
        If draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext3.LineSegment()
            If DispenseModel = 0 Then
                If MainEdge = 1 Then
                    Dim xCenter As Double = (x1 + x2) / 2
                    Dim yCenter As Double = (y1 + y2) / 2
                    Dim M1 As Double = (y2 - y1) / (x2 - x1)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        With AxGraphicContext3.DrawingRegion()
                            .StartX() = xCenter
                            .StartY() = yCenter - EdgeClearance
                        End With
                        AxGraphicContext3.Dot()
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext3.Dot()
                        Else
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext3.Dot()
                        End If
                    End If
                ElseIf MainEdge = 2 Then
                    Dim xCenter As Double = (x2 + x3) / 2
                    Dim yCenter As Double = (y2 + y3) / 2
                    Dim M1 As Double = (y2 - y3) / (x2 - x3)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        With AxGraphicContext3.DrawingRegion()
                            .StartX() = xCenter + EdgeClearance
                            .StartY() = yCenter
                        End With
                        AxGraphicContext3.Dot()
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 < 0 Then
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext3.Dot()
                        Else
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext3.Dot()
                        End If
                    End If
                ElseIf MainEdge = 3 Then
                    Dim xCenter As Double = (x3 + x4) / 2
                    Dim yCenter As Double = (y3 + y4) / 2
                    Dim M1 As Double = (y3 - y4) / (x3 - x4)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        With AxGraphicContext3.DrawingRegion()
                            .StartX() = xCenter
                            .StartY() = yCenter + EdgeClearance
                        End With
                        AxGraphicContext3.Dot()
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext3.Dot()
                        Else
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext3.Dot()
                        End If
                    End If
                ElseIf MainEdge = 4 Then
                    Dim xCenter As Double = (x1 + x4) / 2
                    Dim yCenter As Double = (y1 + y4) / 2
                    Dim M1 As Double = (y1 - y4) / (x1 - x4)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        With AxGraphicContext3.DrawingRegion()
                            .StartX() = xCenter - EdgeClearance
                            .StartY() = yCenter
                        End With
                        AxGraphicContext3.Dot()
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext3.Dot()
                        Else
                            With AxGraphicContext3.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext3.Dot()
                        End If
                    End If
                End If
            ElseIf DispenseModel = 1 Then
                If MainEdge = 1 Then
                    Edge1(x1, x2, y1, y2, EdgeClearance, True)
                ElseIf MainEdge = 2 Then
                    Edge2(x2, x3, y2, y3, EdgeClearance, True)
                ElseIf MainEdge = 3 Then
                    Edge3(x3, x4, y3, y4, EdgeClearance, True)
                ElseIf MainEdge = 4 Then
                    Edge4(x1, x4, y1, y4, EdgeClearance, True)
                End If
            ElseIf DispenseModel = 2 Then
                If MainEdge = 1 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge1(x1, x2, y1, y2, EdgeClearance, True)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge2(x2, x3, y2, y3, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                        'Rectangle(newx1, x2, newx2, newy1, y2, newy2, True)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge4(x1, x4, y1, y4, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                        'Rectangle(newx1, x1, newx2, newy1, y1, newy2, True)
                    End If
                ElseIf MainEdge = 2 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge2(x2, x3, y2, y3, EdgeClearance, True)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge3(x3, x4, y3, y4, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge1(x1, x2, y1, y2, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    End If
                ElseIf MainEdge = 3 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge3(x3, x4, y3, y4, EdgeClearance, True)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge4(x1, x4, y1, y4, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge2(x2, x3, y2, y3, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    End If
                ElseIf MainEdge = 4 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge4(x1, x4, y1, y4, EdgeClearance, True)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge1(x1, x2, y1, y2, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge3(x3, x4, y3, y4, EdgeClearance, True)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, True)
                    End If
                End If
            End If
        Else 'clear
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            AxGraphicContext2.LineSegment()
            If DispenseModel = 0 Then
                If MainEdge = 1 Then
                    Dim xCenter As Double = (x1 + x2) / 2
                    Dim yCenter As Double = (y1 + y2) / 2
                    Dim M1 As Double = (y2 - y1) / (x2 - x1)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        With AxGraphicContext2.DrawingRegion()
                            .StartX() = xCenter
                            .StartY() = yCenter - EdgeClearance
                        End With
                        AxGraphicContext2.Dot()
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext2.Dot()
                        Else
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext2.Dot()
                        End If
                    End If
                ElseIf MainEdge = 2 Then
                    Dim xCenter As Double = (x2 + x3) / 2
                    Dim yCenter As Double = (y2 + y3) / 2
                    Dim M1 As Double = (y2 - y3) / (x2 - x3)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        With AxGraphicContext2.DrawingRegion()
                            .StartX() = xCenter + EdgeClearance
                            .StartY() = yCenter
                        End With
                        AxGraphicContext2.Dot()
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 < 0 Then
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext2.Dot()
                        Else
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext2.Dot()
                        End If
                    End If
                ElseIf MainEdge = 3 Then
                    Dim xCenter As Double = (x3 + x4) / 2
                    Dim yCenter As Double = (y3 + y4) / 2
                    Dim M1 As Double = (y3 - y4) / (x3 - x4)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        With AxGraphicContext2.DrawingRegion()
                            .StartX() = xCenter
                            .StartY() = yCenter + EdgeClearance
                        End With
                        AxGraphicContext2.Dot()
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext2.Dot()
                        Else
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_1
                                .StartY() = y2_1
                            End With
                            AxGraphicContext2.Dot()
                        End If
                    End If
                ElseIf MainEdge = 4 Then
                    Dim xCenter As Double = (x1 + x4) / 2
                    Dim yCenter As Double = (y1 + y4) / 2
                    Dim M1 As Double = (y1 - y4) / (x1 - x4)
                    Dim M2 As Double = -1 / M1
                    Dim C2 As Double = yCenter - M2 * xCenter

                    If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                        With AxGraphicContext2.DrawingRegion()
                            .StartX() = xCenter - EdgeClearance
                            .StartY() = yCenter
                        End With
                        AxGraphicContext2.Dot()
                    ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                        MsgBox("Impossible")
                    Else
                        Dim A As Double = (1 + M2 ^ 2)
                        Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                        Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                        Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                        Dim y2_1 As Double = M2 * x2_1 + C2
                        Dim y2_2 As Double = M2 * x2_2 + C2
                        If m1 > 0 Then
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext2.Dot()
                        Else
                            With AxGraphicContext2.DrawingRegion()
                                .StartX() = x2_2
                                .StartY() = y2_2
                            End With
                            AxGraphicContext2.Dot()
                        End If
                    End If
                End If
            ElseIf DispenseModel = 1 Then
                If MainEdge = 1 Then
                    Edge1(x1, x2, y1, y2, EdgeClearance, False)
                ElseIf MainEdge = 2 Then
                    Edge2(x2, x3, y2, y3, EdgeClearance, False)
                ElseIf MainEdge = 3 Then
                    Edge3(x3, x4, y3, y4, EdgeClearance, False)
                ElseIf MainEdge = 4 Then
                    Edge4(x1, x4, y1, y4, EdgeClearance, False)
                End If
            ElseIf DispenseModel = 2 Then
                If MainEdge = 1 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge1(x1, x2, y1, y2, EdgeClearance, False)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge2(x2, x3, y2, y3, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                        'Rectangle(newx1, x2, newx2, newy1, y2, newy2, True)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge4(x1, x4, y1, y4, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                        'Rectangle(newx1, x1, newx2, newy1, y1, newy2, True)
                    End If
                ElseIf MainEdge = 2 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge2(x2, x3, y2, y3, EdgeClearance, False)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge3(x3, x4, y3, y4, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge1(x1, x2, y1, y2, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    End If
                ElseIf MainEdge = 3 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge3(x3, x4, y3, y4, EdgeClearance, False)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge4(x1, x4, y1, y4, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge2(x2, x3, y2, y3, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m1 = 0 And (m2 > 2 ^ 30 Or m2 < (-(2) ^ 30)) Then
                            cornerx = Px2
                            cornery = PY1
                            'elseIf m1 = 0 Then
                        ElseIf m2 > 2 ^ 30 Or m2 < (-(2) ^ 30) Then
                            cornerx = px2
                            cornery = m1 * cornerx + c1
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    End If
                ElseIf MainEdge = 4 Then
                    Dim Array(4) As Double
                    Dim Array1(4) As Double
                    Array = Edge4(x1, x4, y1, y4, EdgeClearance, False)
                    If Cw_CCW = True Then
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge1(x1, x2, y1, y2, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    Else
                        Dim M1 = Array(0)
                        Dim C1 = Array(1)
                        Dim Px1 = array(2)
                        Dim Py1 = array(3)
                        Array1 = Edge3(x3, x4, y3, y4, EdgeClearance, False)
                        Dim M2 = Array1(0)
                        Dim C2 = Array1(1)
                        Dim Px2 = array1(2)
                        Dim Py2 = array1(3)
                        Dim CornerX = (C1 - C2) / (M2 - M1)
                        Dim CornerY = M1 * CornerX + C1
                        If m2 = 0 And (m1 > 2 ^ 30 Or m1 < (-(2) ^ 30)) Then
                            cornerx = Px1
                            cornery = PY2
                            'elseIf m1 = 0 Then
                        ElseIf m1 > 2 ^ 30 Or m1 < (-(2) ^ 30) Then
                            cornerx = px1
                            cornery = m2 * cornerx + c2
                        End If
                        EdgesConnect(Px1, CornerX, Px2, Py1, CornerY, Py2, False)
                    End If
                End If
            End If
        End If
    End Function
    Function Edge1(ByVal x1, ByVal x2, ByVal y1, ByVal y2, ByVal EdgeClearance, ByVal Draw) As Array
        Dim Array(4) As Double
        Dim xCenter As Double = (x1 + x2) / 2
        Dim yCenter As Double = (y1 + y2) / 2
        Dim M1 As Double = (y2 - y1) / (x2 - x1)
        Dim M2 As Double = -1 / M1
        Dim C2 As Double = yCenter - M2 * xCenter
        Dim C1 As Double = yCenter - M1 * xCenter
        If Draw = True Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                MsgBox("Impossible")
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX = x1
                    .StartY = y1 - EdgeClearance
                    .EndX = x2
                    .EndY = y2 - EdgeClearance
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = 0
                Array(1) = y1 - EdgeClearance
                Array(2) = x2
                Array(3) = y2 - EdgeClearance
            Else 'done
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C3 = y2_1 - M1 * x2_1
                Dim C4 = C3 - C1
                Dim Theta = Acos((EdgeClearance / (-C4)))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)
                If M1 > 0 Then
                    DelX = DelX
                    DelY = -DelY
                ElseIf M1 < 0 Then
                    C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = DelY
                End If
                With AxGraphicContext3.DrawingRegion()
                    .StartX = x1 + DelX
                    .StartY = y1 + DelY
                    .EndX = x2 + DelX
                    .EndY = y2 + DelY
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x2 + DelX
                Array(3) = y2 + DelY
            End If
            Return Array
        ElseIf Draw = False Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                MsgBox("Impossible")
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                With AxGraphicContext2.DrawingRegion()
                    .StartX = x1
                    .StartY = y1 - EdgeClearance
                    .EndX = x2
                    .EndY = y2 - EdgeClearance
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = 0
                Array(1) = y1 - EdgeClearance
                Array(2) = x2
                Array(3) = y2 - EdgeClearance
            Else 'done
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C3 = y2_1 - M1 * x2_1
                Dim C4 = C3 - C1
                Dim Theta = Acos((EdgeClearance / (-C4)))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)
                If M1 > 0 Then
                    DelX = DelX
                    DelY = -DelY
                ElseIf M1 < 0 Then
                    C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = DelY
                End If
                With AxGraphicContext2.DrawingRegion()
                    .StartX = x1 + DelX
                    .StartY = y1 + DelY
                    .EndX = x2 + DelX
                    .EndY = y2 + DelY
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x2 + DelX
                Array(3) = y2 + DelY
            End If
            Return Array
        End If
    End Function
    Function Edge2(ByVal x2, ByVal x3, ByVal y2, ByVal y3, ByVal EdgeClearance, ByVal Draw) As Array
        Dim Array(4) As Double
        Dim xCenter As Double = (x2 + x3) / 2
        Dim yCenter As Double = (y2 + y3) / 2
        Dim M1 As Double = (y3 - y2) / (x3 - x2)
        Dim M2 As Double = -1 / M1
        Dim C2 As Double = yCenter - M2 * xCenter
        Dim C1 As Double = yCenter - M1 * xCenter
        If Draw = True Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x2 + EdgeClearance
                    .StartY() = y2
                    .EndX = x3 + EdgeClearance
                    .EndY = y3
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C1
                Array(2) = x3 + EdgeClearance
                Array(3) = y3
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                MsgBox("Impossible")
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C4 As Double
                Dim C3 = y2_1 - M1 * x2_1
                If M1 < 0 Then
                    C4 = C3 - C1
                ElseIf M1 > 0 Then
                    C4 = C1 - C3
                End If
                Dim Theta = Acos((EdgeClearance / C4))

                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)

                If M1 > 0 Then
                    DelX = DelX
                    DelY = -DelY
                ElseIf M1 < 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = DelX
                    DelY = DelY
                End If

                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x2 + DelX
                    .StartY() = y2 + DelY
                    .EndX = x3 + DelX
                    .EndY = y3 + DelY
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x2 + DelX
                Array(3) = y2 + DelY
            End If
            Return Array
        ElseIf Draw = False Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x2 + EdgeClearance
                    .StartY() = y2
                    .EndX = x3 + EdgeClearance
                    .EndY = y3
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C1
                Array(2) = x3 + EdgeClearance
                Array(3) = y3
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                MsgBox("Impossible")
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C4 As Double
                Dim C3 = y2_1 - M1 * x2_1
                If M1 < 0 Then
                    C4 = C3 - C1
                ElseIf M1 > 0 Then
                    C4 = C1 - C3
                End If
                Dim Theta = Acos((EdgeClearance / C4))

                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)

                If M1 > 0 Then
                    DelX = DelX
                    DelY = -DelY
                ElseIf M1 < 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = DelX
                    DelY = DelY
                End If

                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x2 + DelX
                    .StartY() = y2 + DelY
                    .EndX = x3 + DelX
                    .EndY = y3 + DelY
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x2 + DelX
                Array(3) = y2 + DelY
            End If
            Return Array
        End If
    End Function
    Function Edge3(ByVal x3, ByVal x4, ByVal y3, ByVal y4, ByVal EdgeClearance, ByVal Draw) As Array
        Dim Array(4) As Double
        Dim xCenter As Double = (x3 + x4) / 2
        Dim yCenter As Double = (y3 + y4) / 2
        Dim M1 As Double = (y4 - y3) / (x4 - x3)
        Dim M2 As Double = -1 / M1
        Dim C2 As Double = yCenter - M2 * xCenter
        Dim C1 As Double = yCenter - M1 * xCenter
        If Draw = True Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                MsgBox("Impossible")
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x3
                    .StartY() = y3 + EdgeClearance
                    .EndX = x4
                    .EndY = y4 + EdgeClearance
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = 0
                Array(1) = y3 + EdgeClearance
                Array(2) = x4
                Array(3) = y4 + EdgeClearance
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C3 = y2_2 - M1 * x2_2
                Dim C4 = C3 - C1
                Dim Theta = Acos((EdgeClearance / C4))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)

                If M1 > 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = DelY
                ElseIf M1 < 0 Then
                    DelX = DelX
                    DelY = DelY
                End If

                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x3 + DelX
                    .StartY() = y3 + DelY
                    .EndX = x4 + DelX
                    .EndY = y4 + DelY
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x4 + DelX
                Array(3) = y4 + DelY
            End If
            Return Array
        ElseIf Draw = False Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                MsgBox("Impossible")
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x3
                    .StartY() = y3 + EdgeClearance
                    .EndX = x4
                    .EndY = y4 + EdgeClearance
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = 0
                Array(1) = y3 + EdgeClearance
                Array(2) = x4
                Array(3) = y4 + EdgeClearance
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C3 = y2_2 - M1 * x2_2
                Dim C4 = C3 - C1
                Dim Theta = Acos((EdgeClearance / C4))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)

                If M1 > 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = DelY
                ElseIf M1 < 0 Then
                    DelX = DelX
                    DelY = DelY
                End If

                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x3 + DelX
                    .StartY() = y3 + DelY
                    .EndX = x4 + DelX
                    .EndY = y4 + DelY
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x4 + DelX
                Array(3) = y4 + DelY
            End If
            Return Array
        End If
    End Function
    Function Edge4(ByVal x1, ByVal x4, ByVal y1, ByVal y4, ByVal EdgeClearance, ByVal Draw) As Array
        Dim Array(4) As Double
        Dim xCenter As Double = (x1 + x4) / 2
        Dim yCenter As Double = (y1 + y4) / 2
        Dim M1 As Double = (y4 - y1) / (x4 - x1)
        Dim M2 As Double = -1 / M1
        Dim C2 As Double = yCenter - M2 * xCenter
        Dim C1 As Double = yCenter - M1 * xCenter
        If Draw = True Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x1 - EdgeClearance
                    .StartY() = y1
                    .EndX = x4 - EdgeClearance
                    .EndY = y4
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C1
                Array(2) = x4 - EdgeClearance
                Array(3) = y4
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                MsgBox("Impossible")
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C4 As Double
                Dim C3 = y2_2 - M1 * x2_2

                If M1 < 0 Then
                    C4 = C1 - C3
                ElseIf M1 > 0 Then
                    C4 = C3 - C1
                End If
                Dim Theta = Acos((EdgeClearance / C4))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)
                If M1 > 0 Then
                    DelX = -DelX
                    DelY = DelY
                ElseIf M1 < 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = -DelY
                End If

                With AxGraphicContext3.DrawingRegion()
                    .StartX() = x1 + DelX
                    .StartY() = y1 + DelY
                    .EndX = x4 + DelX
                    .EndY = y4 + DelY
                End With
                AxGraphicContext3.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x4 + DelX
                Array(3) = y4 + DelY
            End If
            Return Array
        ElseIf Draw = False Then
            If M1 > 10 ^ 50 Or M1 < -10 ^ 50 Then
                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x1 - EdgeClearance
                    .StartY() = y1
                    .EndX = x4 - EdgeClearance
                    .EndY = y4
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C1
                Array(2) = x4 - EdgeClearance
                Array(3) = y4
            ElseIf M2 > 10 ^ 50 Or M2 < -10 ^ 50 Then
                MsgBox("Impossible")
            Else
                Dim A As Double = (1 + M2 ^ 2)
                Dim B As Double = (2 * M2 * C2 - 2 * xCenter - 2 * yCenter * M2)
                Dim C As Double = (xCenter ^ 2 + yCenter ^ 2 - 2 * yCenter * C2 + C2 ^ 2 - EdgeClearance ^ 2)
                Dim x2_1 As Double = (-B + Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim x2_2 As Double = (-B - Sqrt(B ^ 2 - (4 * A * C))) / (2 * A)
                Dim y2_1 As Double = M2 * x2_1 + C2
                Dim y2_2 As Double = M2 * x2_2 + C2
                Dim C4 As Double
                Dim C3 = y2_2 - M1 * x2_2

                If M1 < 0 Then
                    C4 = C1 - C3
                ElseIf M1 > 0 Then
                    C4 = C3 - C1
                End If
                Dim Theta = Acos((EdgeClearance / C4))
                Dim DelX As Double = EdgeClearance * Sin(Theta)
                Dim DelY As Double = EdgeClearance * Cos(Theta)
                If M1 > 0 Then
                    DelX = -DelX
                    DelY = DelY
                ElseIf M1 < 0 Then
                    'C3 = y2_2 - x2_2 * M1 'for set 2 data for quadratic ' suppose to be this set of data ' else dely=-dely
                    DelX = -DelX
                    DelY = -DelY
                End If

                With AxGraphicContext2.DrawingRegion()
                    .StartX() = x1 + DelX
                    .StartY() = y1 + DelY
                    .EndX = x4 + DelX
                    .EndY = y4 + DelY
                End With
                AxGraphicContext2.LineSegment()
                Array(0) = M1
                Array(1) = C3
                Array(2) = x4 + DelX
                Array(3) = y4 + DelY
            End If
            Return Array
        End If
    End Function
    Function EdgesConnect(ByVal x1, ByVal x2, ByVal x3, ByVal y1, ByVal y2, ByVal y3, ByVal Draw)
        Dim x4 As Double
        If x2 < x1 Then
            x4 = x3 + Abs(x1 - x2)
        Else
            x4 = x3 - Abs(x1 - x2)
        End If

        Dim y4 As Double
        If y2 < y3 Then
            y4 = (y1 + Abs(y2 - y3))
        Else
            y4 = (y1 - Abs(y2 - y3))
        End If

        If Draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            'AxGraphicContext3.LineSegment()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            'AxGraphicContext3.LineSegment()
        ElseIf Draw = False Then
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x2
                .EndY() = y2
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x2
                .StartY() = y2
                .EndX() = x3
                .EndY() = y3
            End With
            AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x3
                .StartY() = y3
                .EndX() = x4
                .EndY() = y4
            End With
            'AxGraphicContext2.LineSegment()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1
                .StartY() = y1
                .EndX() = x4
                .EndY() = y4
            End With
            'AxGraphicContext2.LineSegment()
        End If
    End Function
    Function Fiducial(ByVal x1 As Double, ByVal x2 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal Draw As Boolean)
        If Draw = True Then
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x1 - 10
                .StartY() = y1 - 10
                .EndX() = x1 + 10
                .EndY() = y1 + 10
            End With
            AxGraphicContext3.Cross()
            With AxGraphicContext3.DrawingRegion()
                .StartX() = x2 - 10
                .StartY() = y2 - 10
                .EndX() = x2 + 10
                .EndY() = y2 + 10
            End With
            AxGraphicContext3.Cross()
        Else
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x1 - 10
                .StartY() = y1 - 10
                .EndX() = x1 + 10
                .EndY() = y1 + 10
            End With
            AxGraphicContext2.Cross()
            With AxGraphicContext2.DrawingRegion()
                .StartX() = x2 - 10
                .StartY() = y2 - 10
                .EndX() = x2 + 10
                .EndY() = y2 + 10
            End With
            AxGraphicContext2.Cross()
        End If
    End Function
#End Region
#Region "Input"
    Function LoadGDFile(ByVal FilePath As String)
        FilePath = FilePath '+ "GD.bmp"
        AxDisplay1.OverlayImage.Load(FilePath)
        AxDisplay1.Width = AxDisplay1.OverlayImage.SizeX
        AxDisplay1.Height = AxDisplay1.OverlayImage.SizeY
        If AxImage1.IsAllocated = True Then
            AxImage1.Free()
        End If
        AxImage1.SizeX = AxDisplay1.OverlayImage.SizeX
        AxImage1.SizeY = AxDisplay1.OverlayImage.SizeY
        AxImage1.Allocate()
        ShowMyImage(FilePath)
        If MyImage.Size.Height > MyImage.Size.Width Then
            Level = 300 / MyImage.Size.Height * 100
        Else
            Level = 300 / MyImage.Size.Width * 100
        End If
        'load RobotRefX and RobotRefY
    End Function
    Function SetGD(ByVal SizeX As Double, ByVal SizeY As Double, ByVal Posx As Double, ByVal PosY As Double)
        AxDisplay1.Width = SizeX * 4
        AxDisplay1.Height = SizeY * 4
        If AxImage1.IsAllocated = True Then
            AxImage1.Free()
        End If
        AxImage1.SizeX = SizeX * 4
        AxImage1.SizeY = SizeY * 4
        AxImage1.Allocate()
        RobotRefX = Posx 'Robot's
        RobotRefY = PosY 'Robot's
    End Function
    Function ShowMyImage(ByVal fileToDisplay As String) 'picturebox2
        ' Sets up an image object to be displayed.
        If Not (MyImage Is Nothing) Then
            MyImage.Dispose()
        End If
        ' Stretches the image to fit the pictureBox. 
        PictureBox2.ViewMode = iViewCore.PictureBox.EViewMode.FitImage
        MyImage = New Bitmap(fileToDisplay)
        If MyImage.Size.Height < 300 And MyImage.Size.Width < 300 Then
            PictureBox2.Image = CType(MyImage, Image)
            If MyImage.Size.Height > MyImage.Size.Width Then
                Level = 300 / MyImage.Size.Height * 100
            Else
                Level = 300 / MyImage.Size.Width * 100
            End If
            PictureBox2.Zoom(Level)
        Else
            PictureBox2.Image = CType(MyImage, Image)
            If MyImage.Size.Height > MyImage.Size.Width Then
                Level = 300 / MyImage.Size.Height * 100
            Else
                Level = 300 / MyImage.Size.Width * 100
            End If
        End If
    End Function
#End Region
#Region "Output"
    Function SaveGDFile(ByVal FilePath As String)
        AxDisplay1.OverlayImage.Save(FilePath + "GD.bmp")
    End Function
#End Region
    '===============================================
#Region "Picturebox2 Zoom"
    Dim MouseRDownFlag As Boolean = False
    Dim MouseFlag As Boolean = False
    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Try
            Level = Level - 25
            If Level < 0 Or Level = 0 Then
                Level = Level + 25
            End If
            PictureBox2.Zoom(Level)
        Catch ex As Exception

        End Try
    End Sub 'zoom out
    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Try
            Level = Level + 25
            If Level > 400 Then
                Level = Level - 25
            End If
            PictureBox2.Zoom(Level)
        Catch ex As Exception
        End Try
    End Sub 'zoom in
    Private Sub PictureBox2_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseWheel
        If MouseFlag = True Then
            If e.Delta < 0 Then
                Level = Level - 25
            Else
                Level = Level + 25
                If Level > 400 Then
                    Level = Level - 25
                End If
            End If
            If Level < 0 Or Level = 0 Then
                Level = Level + 25
            End If
            PictureBox2.Zoom(Level)
        End If
    End Sub
    Private Sub PictureBox2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseDown
        If MouseButtons = MouseButtons.Right Then
            PictureBox2.Zoom(100)
            'PictureBox1.Zoom(100)
            MouseRDownFlag = True
        End If
    End Sub
    Private Sub PictureBox2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseUp
        If MouseRDownFlag = True Then
            ShowMyImage(PFileName)
            MouseRDownFlag = False
        End If
    End Sub
    Private Sub PictureBox2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox2.MouseEnter
        MouseFlag = True
    End Sub
    Private Sub PictureBox2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox2.MouseLeave
        MouseFlag = False
    End Sub
#End Region
#Region "Production"
    Dim RobotPosX As Double
    Dim RobotPosY As Double
    Private Sub ShowMyImage_Production(ByVal fileToDisplay As String) 'picturebox1
        PictureBox1.Image = Image.FromFile(fileToDisplay)
        ' Sets up an image object to be displayed.
        If Not (MyImage1 Is Nothing) Then
            MyImage1.Dispose()
        End If
        'fileToDisplay = PFileName
        ' Stretches the image to fit the pictureBox. 
        PictureBox1.ViewMode = iViewCore.PictureBox.EViewMode.FitImage
        MyImage1 = New Bitmap(fileToDisplay)
        If MyImage1.Size.Height < 600 And MyImage1.Size.Width < 600 Then
            PictureBox1.Image = CType(MyImage1, Image)
            If MyImage1.Size.Height > MyImage1.Size.Width Then
                Ratio = 600 / MyImage1.Size.Height * 100
            Else
                Ratio = 600 / MyImage1.Size.Width * 100
            End If
            PictureBox1.Zoom(Ratio)
        Else
            PictureBox1.Image = CType(MyImage1, Image)
            If MyImage1.Size.Height > MyImage1.Size.Width Then
                Ratio = 600 / MyImage1.Size.Height * 100
            Else
                Ratio = 600 / MyImage1.Size.Width * 100
            End If
        End If
    End Sub
    Function RobotPos()
        Cross_Iview(RobotPosX * Ratio / 100 - RobotRefX, RobotPosY * Ratio / 100 - RobotRefY)
    End Function
    Function Cross_Iview(ByVal x As Integer, ByVal y As Integer)
        'x = (x - RobotRefX) * 4 'for robot
        'y = (y - RobotRefY) * 4
        Try
            myPen = New System.Drawing.Pen(System.Drawing.Color.Red)
            PictureBox1.Refresh()
            PictureBox1.CreateGraphics.DrawRectangle(myPen, x - 5, y - 5, 10, 10)
            PictureBox1.CreateGraphics.DrawLine(myPen, x - 10, y, x + 10, y)
            PictureBox1.CreateGraphics.DrawLine(myPen, x, y - 10, x, y + 10)
        Catch ex As Exception

        End Try

    End Function
    Function LoadGDFile_Production(ByVal FilePath As String, ByVal X As Double, ByVal Y As Double)
        FilePath = FilePath '+ "GD.bmp"
        ShowMyImage_Production(FilePath)
        'load RobotRefX and RobotRefY
        RobotRefX = X
        RobotRefY = Y
    End Function 'useless
#End Region
#Region "MIL"
    Dim ClickRegion As Integer = 5
    Shared mousedownID As Integer = 0
    Dim MouseClickX, MouseClickY As Integer
    Shared Drag_Flag As Integer = 0
    Private Sub AxDisplay1_MouseMoveEvent(ByVal sender As Object, ByVal e As AxMIL.IDisplayEvents_MouseMoveEvent) Handles AxDisplay1.MouseMoveEvent
        PanelPositionX = e.x
        PanelPositionY = e.y
        Label1.Text = Convert.ToString(PanelPositionX)
        Label2.Text = Convert.ToString(PanelPositionY)

        If Drag_Flag = 1 Then
            'AxImage1.Clear()
            'modelregionDrawing(e.x, e.y)
        Else
            Cross(PanelPositionX, PanelPositionY)
        End If
    End Sub
    Private Sub AxDisplay1_MouseDownEvent1(ByVal sender As Object, ByVal e As AxMIL.IDisplayEvents_MouseDownEvent) Handles AxDisplay1.MouseDownEvent
        MouseClickX = e.x
        MouseClickY = e.y
        Drag_Flag = 1

        If MouseButtons = MouseButtons.Right Then
            'AxDisplay1.Zoom(1, 1)
            'PictureBox2.ViewMode = iViewCore.PictureBox.EViewMode.FitImage
            'PictureBox2.Refresh()
        End If
    End Sub
    Private Sub AxDisplay1_MouseUpEvent1(ByVal sender As Object, ByVal e As AxMIL.IDisplayEvents_MouseUpEvent) Handles AxDisplay1.MouseUpEvent
        Drag_Flag = 0
        Try
            Dim ZoomX As Integer = 1 / (Abs(PanelPositionX - MouseClickX) / 300)
            Dim ZoomY As Integer = 1 / (Abs(PanelPositionY - MouseClickY) / 300)
            'AxDisplay1.Zoom(ZoomX, ZoomY)
        Catch ex As Exception
            'MsgBox("error")
        End Try
    End Sub
    Sub modelregionDrawing(ByVal x, ByVal y)
        With AxGraphicContext1.DrawingRegion
            .EndX = x
            .EndY = y
            .StartX = MouseClickX
            .StartY = MouseClickY
        End With
        'AxGraphicContext1.Rectangle()
        AxDisplay1.PanX = MouseClickX - x
        AxDisplay1.PanY = MouseClickY - y
    End Sub
    Function Cross(ByVal x, ByVal y)
        'x = (x - RobotRefX) * 4 'for robot
        'y = (y - RobotRefY) * 4
        AxImage1.Clear()
        With AxGraphicContext1.DrawingRegion()
            .StartX() = x - 10
            .StartY() = y - 10
            .EndX() = x + 10
            .EndY() = y + 10
        End With
        AxGraphicContext1.Cross()
        With AxGraphicContext1.DrawingRegion()
            .StartX() = x - 5
            .StartY() = y - 5
            .EndX() = x + 5
            .EndY() = y + 5
        End With
        AxGraphicContext1.Rectangle(False, 0)
    End Function
#End Region
    '===========================================
#Region "Event Fired button"
    Dim GGID As Integer = 0
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        GGID = 1
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        GGID = 2
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        GGID = 3
    End Sub
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        GGID = 4
    End Sub
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        GGID = 5
    End Sub
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        GGID = 6
    End Sub
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        GGID = 7
    End Sub
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        GGID = 8
    End Sub
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        GGID = 9
    End Sub
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        GGID = 10
    End Sub
    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        GGID = 11
    End Sub
    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        GGID = 12
    End Sub
    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        GGID = 13
    End Sub
    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        GGID = 14
    End Sub
#End Region
#Region "Test n unwanted"
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        ShowMyImage_Production(PFileName)
        'PictureBox1.Image = Image.FromFile(PFileName)
    End Sub
    Private Sub PictureBox1_MouseMove1(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        Cross_Iview(e.X * Ratio / 100, e.Y * Ratio / 100) '(e.X * Ratio / 100-Robotrefx, e.Y * Ratio / 100-Robotrefy) 
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        AxDisplay1.OverlayImage.Clear()
    End Sub 'clear
    Private Sub AxDisplay1_ClickEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles AxDisplay1.ClickEvent
        If GGID <> 0 Then
            If clickno = 1 Then
                a1 = PanelPositionX
                b1 = PanelPositionY
                clickno = 2
                If GGID = 5 Or GGID = 6 Then
                    If GGID = 5 Then
                        Dot(a1, b1, True)
                    Else
                        Dot(a1, b1, False)
                    End If
                    GGID = 0
                    clickno = 1
                End If
            ElseIf clickno = 2 Then
                a2 = PanelPositionX
                b2 = PanelPositionY
                clickno = 3
                If GGID = 1 Or GGID = 4 Then
                    If GGID = 1 Then
                        Line(a1, a2, b1, b2, True)
                    Else
                        Line(a1, a2, b1, b2, False)
                    End If
                    GGID = 0
                    clickno = 1
                End If
            ElseIf clickno = 3 Then
                a3 = PanelPositionX
                b3 = PanelPositionY
                clickno = 4
                If GGID = 2 Or GGID = 3 Or GGID = 7 Or GGID = 9 Or GGID = 10 Or GGID = 11 Or GGID = 12 Then
                    If GGID = 2 Then
                        Circle(a1, a2, a3, b1, b2, b3, True)
                    ElseIf GGID = 3 Then
                        Circle(a1, a2, a3, b1, b2, b3, False)
                    ElseIf GGID = 7 Then
                        Rectangle(a1, a2, a3, b1, b2, b3, True)
                    ElseIf GGID = 8 Then
                        Rectangle(a1, a2, a3, b1, b2, b3, False)
                    ElseIf GGID = 9 Then
                        Arc(a1, a2, a3, b1, b2, b3, True)
                    ElseIf GGID = 10 Then
                        Arc(a1, a2, a3, b1, b2, b3, False)
                    ElseIf GGID = 11 Then
                        CircleSpiral(a1, a2, a3, b1, b2, b3, NumericUpDown3.Value, NumericUpDown2.Value, True, True)
                    ElseIf GGID = 12 Then
                        CircleSpiral(a1, a2, a3, b1, b2, b3, NumericUpDown3.Value, NumericUpDown2.Value, False, False)
                    End If
                    GGID = 0
                    clickno = 1
                Else
                End If
            ElseIf clickno = 4 Then
                a4 = PanelPositionX
                b4 = PanelPositionY
                If GGID = 13 Then
                    ChipEdge(a1, a2, a3, a4, b1, b2, b3, b4, _DispenseModel, _Cw, _MainEdge, _EdgeClearance, True)
                ElseIf GGID = 14 Then
                    ChipEdge(a1, a2, a3, a4, b1, b2, b3, b4, _DispenseModel, _Cw, _MainEdge, _EdgeClearance, False)
                End If
                GGID = 0
                clickno = 1
            End If
        Else
            GGID = 0
            clickno = 1
        End If
        AxDisplay1.OverlayImage.Save(PFileName)
        ShowMyImage(PFileName)
    End Sub 'robot keyin
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        AxDisplay1.OverlayImage.Save("")
    End Sub 'save
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        AxDisplay1.OverlayImage.Load("")
        If AxImage1.IsAllocated = True Then
            AxImage1.Free()
        End If
        AxImage1.SizeX = AxDisplay1.OverlayImage.SizeX
        AxImage1.SizeY = AxDisplay1.OverlayImage.SizeY
        AxImage1.Allocate()
        AxDisplay1.Width = AxDisplay1.OverlayImage.SizeX
        AxDisplay1.Height = AxDisplay1.OverlayImage.SizeY
        Try
            'AxDisplay1.OverlayImage.Save(PFileName)
        Catch ex As Exception

        End Try
        ShowMyImage(PFileName)
        LoadGDFile(PFileName)
    End Sub 'load
    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        SetGD(250, 150, 0, 0)
    End Sub
    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        SetGD(20, 20, 0, 0)
    End Sub
    Private Sub NumericUpDown6_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown6.ValueChanged
        GGID = NumericUpDown6.Value
    End Sub
#End Region
#Region "ChipEdge"
    Dim _Cw = True
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If _Cw = True Then
            _Cw = False
        Else
            _Cw = True
        End If
    End Sub
    Dim _MainEdge As Integer = 1
    Private Sub NumericUpDown4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.ValueChanged
        _MainEdge = NumericUpDown4.Value
    End Sub
    Dim _EdgeClearance As Double = 0
    Private Sub NumericUpDown5_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown5.ValueChanged
        _EdgeClearance = NumericUpDown5.Value
    End Sub
    Dim _DispenseModel As Integer = 0

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = "Dot" Then
            _DispenseModel = 0
        ElseIf ComboBox1.Text = "One edge" Then
            _DispenseModel = 1
        ElseIf ComboBox1.Text = "Two edges" Then
            _DispenseModel = 2
        End If

    End Sub
#End Region
End Class

'Comments~
'==========
'1) Spiral Rect haven done
'