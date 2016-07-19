
Imports DLL_Export_Service

Public Class VisionSetup

    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "
    ''''Dim frm As FormVision
    Public Sub New()
        ''''Public Sub New(ByVal pt As FormVision)
        MyBase.New()
        '''frm = pt
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
    Friend WithEvents GroupBox12 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_RDia As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Score As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RRot As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RPosY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RPosX As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_RSizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents Label45 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label50 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Label52 As System.Windows.Forms.Label
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents TextBox_RSizeX As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents NSensorYPos As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents NSensorXPos As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox_RefBlkRect As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_Rot As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_Pos As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_Size As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents TextBox_PosY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PosX As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_SizeX As System.Windows.Forms.TextBox
    Friend WithEvents Label56 As System.Windows.Forms.Label
    Friend WithEvents TextBox_SizeY As System.Windows.Forms.TextBox
    Friend WithEvents Label57 As System.Windows.Forms.Label
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Imaging As System.Windows.Forms.GroupBox
    Friend WithEvents Label58 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_ROI As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label59 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_Contrast As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label60 As System.Windows.Forms.Label
    Friend WithEvents Label61 As System.Windows.Forms.Label
    Friend WithEvents Label62 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_Brightness As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_Threshold As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox10 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton_Outside_In As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Inside_out As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Horizontal As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_Verical As System.Windows.Forms.RadioButton
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonBack As System.Windows.Forms.Button
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Vertical_Horizontal As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_NoOfRows As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown_Spacing As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Calibrate As System.Windows.Forms.Button
    Friend WithEvents TextBox_PixelY As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_PixelX As System.Windows.Forms.TextBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Ratio As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label77 As System.Windows.Forms.Label
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown_NoOfColumns As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents PanelToBeAdded2 As System.Windows.Forms.Panel
    Friend WithEvents btGetReferenceCenter As System.Windows.Forms.Button
    Friend WithEvents rtbTips As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(VisionSetup))
        Me.GroupBox12 = New System.Windows.Forms.GroupBox
        Me.TextBox_RDia = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextBox_Score = New System.Windows.Forms.TextBox
        Me.TextBox_RRot = New System.Windows.Forms.TextBox
        Me.TextBox_RPosY = New System.Windows.Forms.TextBox
        Me.TextBox_RPosX = New System.Windows.Forms.TextBox
        Me.TextBox_RSizeY = New System.Windows.Forms.TextBox
        Me.Label41 = New System.Windows.Forms.Label
        Me.Label42 = New System.Windows.Forms.Label
        Me.Label43 = New System.Windows.Forms.Label
        Me.Label44 = New System.Windows.Forms.Label
        Me.Label45 = New System.Windows.Forms.Label
        Me.Label46 = New System.Windows.Forms.Label
        Me.Label47 = New System.Windows.Forms.Label
        Me.Label48 = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.Label50 = New System.Windows.Forms.Label
        Me.Label51 = New System.Windows.Forms.Label
        Me.Label52 = New System.Windows.Forms.Label
        Me.Label53 = New System.Windows.Forms.Label
        Me.TextBox_RSizeX = New System.Windows.Forms.TextBox
        Me.rtbTips = New System.Windows.Forms.RichTextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label11 = New System.Windows.Forms.Label
        Me.NSensorYPos = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.NSensorXPos = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label77 = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Button_Calibrate = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.TextBox_Ratio = New System.Windows.Forms.TextBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.TextBox_PixelY = New System.Windows.Forms.TextBox
        Me.TextBox_PixelX = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.NumericUpDown_Spacing = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.NumericUpDown_NoOfRows = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown_NoOfColumns = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ButtonBack = New System.Windows.Forms.Button
        Me.GroupBox_RefBlkRect = New System.Windows.Forms.GroupBox
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.btGetReferenceCenter = New System.Windows.Forms.Button
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.NumericUpDown_Rot = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown_Pos = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown_Size = New System.Windows.Forms.NumericUpDown
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.Label35 = New System.Windows.Forms.Label
        Me.Label36 = New System.Windows.Forms.Label
        Me.Label37 = New System.Windows.Forms.Label
        Me.Label38 = New System.Windows.Forms.Label
        Me.Label39 = New System.Windows.Forms.Label
        Me.Label54 = New System.Windows.Forms.Label
        Me.Label55 = New System.Windows.Forms.Label
        Me.TextBox_PosY = New System.Windows.Forms.TextBox
        Me.TextBox_PosX = New System.Windows.Forms.TextBox
        Me.TextBox_SizeX = New System.Windows.Forms.TextBox
        Me.Label56 = New System.Windows.Forms.Label
        Me.TextBox_SizeY = New System.Windows.Forms.TextBox
        Me.Label57 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button_Test = New System.Windows.Forms.Button
        Me.Label40 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Button6 = New System.Windows.Forms.Button
        Me.GroupBox_Imaging = New System.Windows.Forms.GroupBox
        Me.Label58 = New System.Windows.Forms.Label
        Me.NumericUpDown_ROI = New System.Windows.Forms.NumericUpDown
        Me.Label59 = New System.Windows.Forms.Label
        Me.NumericUpDown_Contrast = New System.Windows.Forms.NumericUpDown
        Me.Label60 = New System.Windows.Forms.Label
        Me.Label61 = New System.Windows.Forms.Label
        Me.Label62 = New System.Windows.Forms.Label
        Me.NumericUpDown_Brightness = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown_Threshold = New System.Windows.Forms.NumericUpDown
        Me.GroupBox10 = New System.Windows.Forms.GroupBox
        Me.RadioButton_Outside_In = New System.Windows.Forms.RadioButton
        Me.RadioButton_Inside_out = New System.Windows.Forms.RadioButton
        Me.GroupBox_Vertical_Horizontal = New System.Windows.Forms.GroupBox
        Me.RadioButton_Horizontal = New System.Windows.Forms.RadioButton
        Me.RadioButton_Verical = New System.Windows.Forms.RadioButton
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PanelToBeAdded2 = New System.Windows.Forms.Panel
        Me.Label20 = New System.Windows.Forms.Label
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.GroupBox12.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.NumericUpDown_Spacing, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_NoOfRows, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_NoOfColumns, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox_RefBlkRect.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        CType(Me.NumericUpDown_Rot, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Pos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Size, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_Imaging.SuspendLayout()
        CType(Me.NumericUpDown_ROI, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Contrast, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Brightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_Threshold, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox_Vertical_Horizontal.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        Me.PanelToBeAdded2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.TextBox_RDia)
        Me.GroupBox12.Controls.Add(Me.Label12)
        Me.GroupBox12.Controls.Add(Me.Label15)
        Me.GroupBox12.Controls.Add(Me.TextBox_Score)
        Me.GroupBox12.Controls.Add(Me.TextBox_RRot)
        Me.GroupBox12.Controls.Add(Me.TextBox_RPosY)
        Me.GroupBox12.Controls.Add(Me.TextBox_RPosX)
        Me.GroupBox12.Controls.Add(Me.TextBox_RSizeY)
        Me.GroupBox12.Controls.Add(Me.Label41)
        Me.GroupBox12.Controls.Add(Me.Label42)
        Me.GroupBox12.Controls.Add(Me.Label43)
        Me.GroupBox12.Controls.Add(Me.Label44)
        Me.GroupBox12.Controls.Add(Me.Label45)
        Me.GroupBox12.Controls.Add(Me.Label46)
        Me.GroupBox12.Controls.Add(Me.Label47)
        Me.GroupBox12.Controls.Add(Me.Label48)
        Me.GroupBox12.Controls.Add(Me.Label49)
        Me.GroupBox12.Controls.Add(Me.Label50)
        Me.GroupBox12.Controls.Add(Me.Label51)
        Me.GroupBox12.Controls.Add(Me.Label52)
        Me.GroupBox12.Controls.Add(Me.Label53)
        Me.GroupBox12.Controls.Add(Me.TextBox_RSizeX)
        Me.GroupBox12.Controls.Add(Me.rtbTips)
        Me.GroupBox12.Location = New System.Drawing.Point(8, 496)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(480, 232)
        Me.GroupBox12.TabIndex = 61
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Results"
        '
        'TextBox_RDia
        '
        Me.TextBox_RDia.Location = New System.Drawing.Point(136, 67)
        Me.TextBox_RDia.Name = "TextBox_RDia"
        Me.TextBox_RDia.ReadOnly = True
        Me.TextBox_RDia.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RDia.TabIndex = 68
        Me.TextBox_RDia.Text = "0"
        Me.TextBox_RDia.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(208, 67)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(48, 23)
        Me.Label12.TabIndex = 66
        Me.Label12.Text = "mm"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(88, 67)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(48, 24)
        Me.Label15.TabIndex = 67
        Me.Label15.Text = "Dia="
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_Score
        '
        Me.TextBox_Score.Location = New System.Drawing.Point(136, 192)
        Me.TextBox_Score.Name = "TextBox_Score"
        Me.TextBox_Score.ReadOnly = True
        Me.TextBox_Score.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_Score.TabIndex = 43
        Me.TextBox_Score.Text = "0"
        Me.TextBox_Score.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RRot
        '
        Me.TextBox_RRot.Location = New System.Drawing.Point(136, 168)
        Me.TextBox_RRot.Name = "TextBox_RRot"
        Me.TextBox_RRot.ReadOnly = True
        Me.TextBox_RRot.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RRot.TabIndex = 65
        Me.TextBox_RRot.Text = "0"
        Me.TextBox_RRot.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RPosY
        '
        Me.TextBox_RPosY.Location = New System.Drawing.Point(136, 128)
        Me.TextBox_RPosY.Name = "TextBox_RPosY"
        Me.TextBox_RPosY.ReadOnly = True
        Me.TextBox_RPosY.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RPosY.TabIndex = 64
        Me.TextBox_RPosY.Text = "0"
        Me.TextBox_RPosY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RPosX
        '
        Me.TextBox_RPosX.Location = New System.Drawing.Point(136, 104)
        Me.TextBox_RPosX.Name = "TextBox_RPosX"
        Me.TextBox_RPosX.ReadOnly = True
        Me.TextBox_RPosX.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RPosX.TabIndex = 63
        Me.TextBox_RPosX.Text = "0"
        Me.TextBox_RPosX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_RSizeY
        '
        Me.TextBox_RSizeY.Location = New System.Drawing.Point(136, 42)
        Me.TextBox_RSizeY.Name = "TextBox_RSizeY"
        Me.TextBox_RSizeY.ReadOnly = True
        Me.TextBox_RSizeY.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RSizeY.TabIndex = 62
        Me.TextBox_RSizeY.Text = "0"
        Me.TextBox_RSizeY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label41
        '
        Me.Label41.Location = New System.Drawing.Point(208, 42)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(48, 23)
        Me.Label41.TabIndex = 44
        Me.Label41.Text = "mm"
        Me.Label41.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label42
        '
        Me.Label42.Location = New System.Drawing.Point(208, 128)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(48, 16)
        Me.Label42.TabIndex = 55
        Me.Label42.Text = "mm"
        Me.Label42.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label43
        '
        Me.Label43.Location = New System.Drawing.Point(208, 104)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(48, 16)
        Me.Label43.TabIndex = 54
        Me.Label43.Text = "mm"
        Me.Label43.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label44
        '
        Me.Label44.Location = New System.Drawing.Point(96, 128)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(40, 24)
        Me.Label44.TabIndex = 53
        Me.Label44.Text = "Y="
        Me.Label44.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label45
        '
        Me.Label45.Location = New System.Drawing.Point(96, 42)
        Me.Label45.Name = "Label45"
        Me.Label45.Size = New System.Drawing.Size(40, 24)
        Me.Label45.TabIndex = 52
        Me.Label45.Text = "Y="
        Me.Label45.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label46
        '
        Me.Label46.Location = New System.Drawing.Point(96, 104)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(40, 24)
        Me.Label46.TabIndex = 51
        Me.Label46.Text = "X="
        Me.Label46.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label47
        '
        Me.Label47.Location = New System.Drawing.Point(96, 16)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(40, 24)
        Me.Label47.TabIndex = 50
        Me.Label47.Text = "X="
        Me.Label47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label48
        '
        Me.Label48.Location = New System.Drawing.Point(16, 168)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(80, 16)
        Me.Label48.TabIndex = 49
        Me.Label48.Text = "Rotation:"
        Me.Label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label49
        '
        Me.Label49.Location = New System.Drawing.Point(16, 104)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(80, 24)
        Me.Label49.TabIndex = 48
        Me.Label49.Text = "Position:"
        Me.Label49.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label50
        '
        Me.Label50.Location = New System.Drawing.Point(16, 16)
        Me.Label50.Name = "Label50"
        Me.Label50.Size = New System.Drawing.Size(80, 24)
        Me.Label50.TabIndex = 47
        Me.Label50.Text = "Size:"
        Me.Label50.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label51
        '
        Me.Label51.Location = New System.Drawing.Point(208, 168)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(48, 16)
        Me.Label51.TabIndex = 46
        Me.Label51.Text = "deg"
        Me.Label51.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label52
        '
        Me.Label52.Location = New System.Drawing.Point(208, 16)
        Me.Label52.Name = "Label52"
        Me.Label52.Size = New System.Drawing.Size(48, 23)
        Me.Label52.TabIndex = 45
        Me.Label52.Text = "mm"
        Me.Label52.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label53
        '
        Me.Label53.Location = New System.Drawing.Point(16, 192)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(80, 24)
        Me.Label53.TabIndex = 42
        Me.Label53.Text = "Score:"
        Me.Label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_RSizeX
        '
        Me.TextBox_RSizeX.Location = New System.Drawing.Point(136, 16)
        Me.TextBox_RSizeX.Name = "TextBox_RSizeX"
        Me.TextBox_RSizeX.ReadOnly = True
        Me.TextBox_RSizeX.Size = New System.Drawing.Size(72, 27)
        Me.TextBox_RSizeX.TabIndex = 61
        Me.TextBox_RSizeX.Text = "0"
        Me.TextBox_RSizeX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'rtbTips
        '
        Me.rtbTips.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.rtbTips.Location = New System.Drawing.Point(264, 16)
        Me.rtbTips.Name = "rtbTips"
        Me.rtbTips.ReadOnly = True
        Me.rtbTips.Size = New System.Drawing.Size(208, 208)
        Me.rtbTips.TabIndex = 66
        Me.rtbTips.Text = "Tips:"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(296, 32)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Camera Calibration Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 45
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.NSensorYPos)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.NSensorXPos)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(528, 864)
        Me.Panel2.TabIndex = 44
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(376, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(16, 23)
        Me.Label11.TabIndex = 66
        Me.Label11.Text = ")"
        '
        'NSensorYPos
        '
        Me.NSensorYPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NSensorYPos.Location = New System.Drawing.Point(320, 24)
        Me.NSensorYPos.Name = "NSensorYPos"
        Me.NSensorYPos.Size = New System.Drawing.Size(56, 23)
        Me.NSensorYPos.TabIndex = 62
        Me.NSensorYPos.Text = "100"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(296, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 23)
        Me.Label4.TabIndex = 61
        Me.Label4.Text = ","
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(216, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 23)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "("
        '
        'NSensorXPos
        '
        Me.NSensorXPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NSensorXPos.Location = New System.Drawing.Point(232, 24)
        Me.NSensorXPos.Name = "NSensorXPos"
        Me.NSensorXPos.Size = New System.Drawing.Size(56, 23)
        Me.NSensorXPos.TabIndex = 59
        Me.NSensorXPos.Text = "100"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 24)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "Camera Position:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Controls.Add(Me.Label77)
        Me.GroupBox3.Controls.Add(Me.PictureBox2)
        Me.GroupBox3.Controls.Add(Me.Button_Calibrate)
        Me.GroupBox3.Controls.Add(Me.GroupBox4)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.NumericUpDown_Spacing)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.NumericUpDown_NoOfRows)
        Me.GroupBox3.Controls.Add(Me.NumericUpDown_NoOfColumns)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.Button5)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(8, 56)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(496, 800)
        Me.GroupBox3.TabIndex = 41
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Locate Reference Block Position"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(248, 376)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 50)
        Me.Button1.TabIndex = 31
        Me.Button1.Text = "Save"
        '
        'Label77
        '
        Me.Label77.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label77.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label77.Location = New System.Drawing.Point(24, 456)
        Me.Label77.Name = "Label77"
        Me.Label77.Size = New System.Drawing.Size(456, 48)
        Me.Label77.TabIndex = 30
        Me.Label77.Text = "2. Jog camera to the approximate center of the reference block, and click Next."
        '
        'PictureBox2
        '
        Me.PictureBox2.AccessibleName = ""
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(160, 512)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(191, 159)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox2.TabIndex = 29
        Me.PictureBox2.TabStop = False
        '
        'Button_Calibrate
        '
        Me.Button_Calibrate.Location = New System.Drawing.Point(368, 152)
        Me.Button_Calibrate.Name = "Button_Calibrate"
        Me.Button_Calibrate.Size = New System.Drawing.Size(100, 50)
        Me.Button_Calibrate.TabIndex = 27
        Me.Button_Calibrate.Text = "Calibrate"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TextBox_Ratio)
        Me.GroupBox4.Controls.Add(Me.Label16)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.TextBox_PixelY)
        Me.GroupBox4.Controls.Add(Me.TextBox_PixelX)
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.Label18)
        Me.GroupBox4.Controls.Add(Me.Label19)
        Me.GroupBox4.Controls.Add(Me.Label17)
        Me.GroupBox4.Location = New System.Drawing.Point(80, 224)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(280, 128)
        Me.GroupBox4.TabIndex = 26
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Results:"
        '
        'TextBox_Ratio
        '
        Me.TextBox_Ratio.Location = New System.Drawing.Point(120, 88)
        Me.TextBox_Ratio.Name = "TextBox_Ratio"
        Me.TextBox_Ratio.ReadOnly = True
        Me.TextBox_Ratio.TabIndex = 28
        Me.TextBox_Ratio.Text = ""
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(10, 88)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(104, 23)
        Me.Label16.TabIndex = 27
        Me.Label16.Text = "Pixel Ratio:"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(8, 24)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(104, 23)
        Me.Label13.TabIndex = 24
        Me.Label13.Text = "Pixel Size Z:"
        '
        'TextBox_PixelY
        '
        Me.TextBox_PixelY.Location = New System.Drawing.Point(120, 56)
        Me.TextBox_PixelY.Name = "TextBox_PixelY"
        Me.TextBox_PixelY.ReadOnly = True
        Me.TextBox_PixelY.TabIndex = 23
        Me.TextBox_PixelY.Text = ""
        '
        'TextBox_PixelX
        '
        Me.TextBox_PixelX.Location = New System.Drawing.Point(120, 24)
        Me.TextBox_PixelX.Name = "TextBox_PixelX"
        Me.TextBox_PixelX.ReadOnly = True
        Me.TextBox_PixelX.TabIndex = 22
        Me.TextBox_PixelX.Text = ""
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(8, 56)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(104, 23)
        Me.Label14.TabIndex = 25
        Me.Label14.Text = "Pixel Size Y:"
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(224, 24)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(48, 23)
        Me.Label18.TabIndex = 29
        Me.Label18.Text = "mm"
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(224, 56)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(48, 23)
        Me.Label19.TabIndex = 30
        Me.Label19.Text = "mm"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(224, 88)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(48, 23)
        Me.Label17.TabIndex = 28
        Me.Label17.Text = "mm"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(304, 176)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(48, 23)
        Me.Label10.TabIndex = 21
        Me.Label10.Text = "mm"
        '
        'NumericUpDown_Spacing
        '
        Me.NumericUpDown_Spacing.DecimalPlaces = 3
        Me.NumericUpDown_Spacing.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.NumericUpDown_Spacing.Location = New System.Drawing.Point(200, 176)
        Me.NumericUpDown_Spacing.Maximum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.NumericUpDown_Spacing.Name = "NumericUpDown_Spacing"
        Me.NumericUpDown_Spacing.Size = New System.Drawing.Size(104, 27)
        Me.NumericUpDown_Spacing.TabIndex = 20
        Me.NumericUpDown_Spacing.Value = New Decimal(New Integer() {15, 0, 0, 65536})
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(32, 176)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(160, 23)
        Me.Label8.TabIndex = 19
        Me.Label8.Text = "Dot to Dot Spacing:"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 136)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 23)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Number of Rows:"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(32, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 23)
        Me.Label6.TabIndex = 17
        Me.Label6.Text = "No of Columns:"
        '
        'NumericUpDown_NoOfRows
        '
        Me.NumericUpDown_NoOfRows.Location = New System.Drawing.Point(200, 136)
        Me.NumericUpDown_NoOfRows.Name = "NumericUpDown_NoOfRows"
        Me.NumericUpDown_NoOfRows.Size = New System.Drawing.Size(104, 27)
        Me.NumericUpDown_NoOfRows.TabIndex = 15
        Me.NumericUpDown_NoOfRows.Value = New Decimal(New Integer() {8, 0, 0, 0})
        '
        'NumericUpDown_NoOfColumns
        '
        Me.NumericUpDown_NoOfColumns.Location = New System.Drawing.Point(200, 96)
        Me.NumericUpDown_NoOfColumns.Name = "NumericUpDown_NoOfColumns"
        Me.NumericUpDown_NoOfColumns.Size = New System.Drawing.Size(104, 27)
        Me.NumericUpDown_NoOfColumns.TabIndex = 14
        Me.NumericUpDown_NoOfColumns.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label5.Location = New System.Drawing.Point(20, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(388, 32)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "1. Jog camera to the vision calibration position. "
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(368, 736)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 50)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Next"
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(368, 376)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(100, 50)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Cancel"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ButtonSave)
        Me.GroupBox1.Controls.Add(Me.ButtonBack)
        Me.GroupBox1.Controls.Add(Me.GroupBox_RefBlkRect)
        Me.GroupBox1.Controls.Add(Me.GroupBox12)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 56)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 800)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Locate Reference Block Position"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(368, 736)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 50)
        Me.Button2.TabIndex = 62
        Me.Button2.Text = "Cancel"
        '
        'Label3
        '
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(24, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(456, 40)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "3. Click 5 points on the reference block's edges according to vertical/ horizonta" & _
        "l direction."
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(200, 736)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(100, 50)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        '
        'ButtonBack
        '
        Me.ButtonBack.Location = New System.Drawing.Point(32, 736)
        Me.ButtonBack.Name = "ButtonBack"
        Me.ButtonBack.Size = New System.Drawing.Size(100, 50)
        Me.ButtonBack.TabIndex = 7
        Me.ButtonBack.Text = "Back"
        '
        'GroupBox_RefBlkRect
        '
        Me.GroupBox_RefBlkRect.Controls.Add(Me.GroupBox9)
        Me.GroupBox_RefBlkRect.Controls.Add(Me.GroupBox_Imaging)
        Me.GroupBox_RefBlkRect.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox_RefBlkRect.Name = "GroupBox_RefBlkRect"
        Me.GroupBox_RefBlkRect.Size = New System.Drawing.Size(480, 408)
        Me.GroupBox_RefBlkRect.TabIndex = 0
        Me.GroupBox_RefBlkRect.TabStop = False
        Me.GroupBox_RefBlkRect.Text = "Reference Block Rectangle"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.btGetReferenceCenter)
        Me.GroupBox9.Controls.Add(Me.Label29)
        Me.GroupBox9.Controls.Add(Me.Label30)
        Me.GroupBox9.Controls.Add(Me.NumericUpDown_Rot)
        Me.GroupBox9.Controls.Add(Me.NumericUpDown_Pos)
        Me.GroupBox9.Controls.Add(Me.NumericUpDown_Size)
        Me.GroupBox9.Controls.Add(Me.Label31)
        Me.GroupBox9.Controls.Add(Me.Label32)
        Me.GroupBox9.Controls.Add(Me.Label33)
        Me.GroupBox9.Controls.Add(Me.Label34)
        Me.GroupBox9.Controls.Add(Me.Label35)
        Me.GroupBox9.Controls.Add(Me.Label36)
        Me.GroupBox9.Controls.Add(Me.Label37)
        Me.GroupBox9.Controls.Add(Me.Label38)
        Me.GroupBox9.Controls.Add(Me.Label39)
        Me.GroupBox9.Controls.Add(Me.Label54)
        Me.GroupBox9.Controls.Add(Me.Label55)
        Me.GroupBox9.Controls.Add(Me.TextBox_PosY)
        Me.GroupBox9.Controls.Add(Me.TextBox_PosX)
        Me.GroupBox9.Controls.Add(Me.TextBox_SizeX)
        Me.GroupBox9.Controls.Add(Me.Label56)
        Me.GroupBox9.Controls.Add(Me.TextBox_SizeY)
        Me.GroupBox9.Controls.Add(Me.Label57)
        Me.GroupBox9.Controls.Add(Me.Button4)
        Me.GroupBox9.Controls.Add(Me.Button_Test)
        Me.GroupBox9.Controls.Add(Me.Label40)
        Me.GroupBox9.Controls.Add(Me.Label28)
        Me.GroupBox9.Controls.Add(Me.Button6)
        Me.GroupBox9.Location = New System.Drawing.Point(8, 24)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(466, 184)
        Me.GroupBox9.TabIndex = 1
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Rectangle's Parameters"
        '
        'btGetReferenceCenter
        '
        Me.btGetReferenceCenter.Location = New System.Drawing.Point(272, 136)
        Me.btGetReferenceCenter.Name = "btGetReferenceCenter"
        Me.btGetReferenceCenter.Size = New System.Drawing.Size(64, 40)
        Me.btGetReferenceCenter.TabIndex = 44
        Me.btGetReferenceCenter.Text = "Get"
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(432, 72)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(24, 23)
        Me.Label29.TabIndex = 38
        Me.Label29.Text = "%"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(432, 32)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(24, 23)
        Me.Label30.TabIndex = 37
        Me.Label30.Text = "%"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumericUpDown_Rot
        '
        Me.NumericUpDown_Rot.DecimalPlaces = 1
        Me.NumericUpDown_Rot.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NumericUpDown_Rot.Location = New System.Drawing.Point(104, 112)
        Me.NumericUpDown_Rot.Maximum = New Decimal(New Integer() {5, 0, 0, 0})
        Me.NumericUpDown_Rot.Name = "NumericUpDown_Rot"
        Me.NumericUpDown_Rot.Size = New System.Drawing.Size(72, 27)
        Me.NumericUpDown_Rot.TabIndex = 36
        Me.NumericUpDown_Rot.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Rot.Value = New Decimal(New Integer() {50, 0, 0, 65536})
        '
        'NumericUpDown_Pos
        '
        Me.NumericUpDown_Pos.DecimalPlaces = 1
        Me.NumericUpDown_Pos.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NumericUpDown_Pos.Location = New System.Drawing.Point(384, 72)
        Me.NumericUpDown_Pos.Name = "NumericUpDown_Pos"
        Me.NumericUpDown_Pos.Size = New System.Drawing.Size(44, 27)
        Me.NumericUpDown_Pos.TabIndex = 35
        Me.NumericUpDown_Pos.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Pos.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'NumericUpDown_Size
        '
        Me.NumericUpDown_Size.DecimalPlaces = 1
        Me.NumericUpDown_Size.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NumericUpDown_Size.Location = New System.Drawing.Point(384, 32)
        Me.NumericUpDown_Size.Name = "NumericUpDown_Size"
        Me.NumericUpDown_Size.Size = New System.Drawing.Size(44, 27)
        Me.NumericUpDown_Size.TabIndex = 34
        Me.NumericUpDown_Size.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Size.Value = New Decimal(New Integer() {50, 0, 0, 65536})
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.Location = New System.Drawing.Point(360, 72)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(24, 23)
        Me.Label31.TabIndex = 32
        Me.Label31.Text = "?"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(360, 32)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(24, 23)
        Me.Label32.TabIndex = 31
        Me.Label32.Text = "?"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label33
        '
        Me.Label33.Location = New System.Drawing.Point(312, 72)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(44, 23)
        Me.Label33.TabIndex = 30
        Me.Label33.Text = "mm,"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.Location = New System.Drawing.Point(176, 72)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(44, 23)
        Me.Label34.TabIndex = 29
        Me.Label34.Text = "mm,"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label35
        '
        Me.Label35.Location = New System.Drawing.Point(224, 72)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(16, 24)
        Me.Label35.TabIndex = 28
        Me.Label35.Text = "Y"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label36
        '
        Me.Label36.Location = New System.Drawing.Point(224, 32)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(16, 24)
        Me.Label36.TabIndex = 27
        Me.Label36.Text = "Y"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label37
        '
        Me.Label37.Location = New System.Drawing.Point(88, 72)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(16, 24)
        Me.Label37.TabIndex = 26
        Me.Label37.Text = "X"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label38
        '
        Me.Label38.Location = New System.Drawing.Point(88, 32)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(16, 24)
        Me.Label38.TabIndex = 25
        Me.Label38.Text = "X"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label39
        '
        Me.Label39.Location = New System.Drawing.Point(8, 112)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(80, 24)
        Me.Label39.TabIndex = 24
        Me.Label39.Text = "Rotation:"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label54
        '
        Me.Label54.Location = New System.Drawing.Point(8, 32)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(56, 24)
        Me.Label54.TabIndex = 22
        Me.Label54.Text = "Size:"
        Me.Label54.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label55
        '
        Me.Label55.Location = New System.Drawing.Point(176, 112)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(44, 23)
        Me.Label55.TabIndex = 21
        Me.Label55.Text = "deg,"
        Me.Label55.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_PosY
        '
        Me.TextBox_PosY.Location = New System.Drawing.Point(240, 72)
        Me.TextBox_PosY.Name = "TextBox_PosY"
        Me.TextBox_PosY.Size = New System.Drawing.Size(68, 27)
        Me.TextBox_PosY.TabIndex = 18
        Me.TextBox_PosY.Text = "0"
        Me.TextBox_PosY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_PosX
        '
        Me.TextBox_PosX.Location = New System.Drawing.Point(104, 72)
        Me.TextBox_PosX.Name = "TextBox_PosX"
        Me.TextBox_PosX.Size = New System.Drawing.Size(68, 27)
        Me.TextBox_PosX.TabIndex = 16
        Me.TextBox_PosX.Text = "555.555"
        Me.TextBox_PosX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox_SizeX
        '
        Me.TextBox_SizeX.Location = New System.Drawing.Point(104, 32)
        Me.TextBox_SizeX.Name = "TextBox_SizeX"
        Me.TextBox_SizeX.Size = New System.Drawing.Size(68, 27)
        Me.TextBox_SizeX.TabIndex = 14
        Me.TextBox_SizeX.Text = "55.555"
        Me.TextBox_SizeX.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label56
        '
        Me.Label56.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label56.Location = New System.Drawing.Point(176, 32)
        Me.Label56.Name = "Label56"
        Me.Label56.Size = New System.Drawing.Size(44, 23)
        Me.Label56.TabIndex = 15
        Me.Label56.Text = "mm,"
        Me.Label56.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_SizeY
        '
        Me.TextBox_SizeY.Location = New System.Drawing.Point(240, 32)
        Me.TextBox_SizeY.Name = "TextBox_SizeY"
        Me.TextBox_SizeY.Size = New System.Drawing.Size(68, 27)
        Me.TextBox_SizeY.TabIndex = 12
        Me.TextBox_SizeY.Text = "0"
        Me.TextBox_SizeY.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label57
        '
        Me.Label57.Location = New System.Drawing.Point(312, 32)
        Me.Label57.Name = "Label57"
        Me.Label57.Size = New System.Drawing.Size(44, 23)
        Me.Label57.TabIndex = 13
        Me.Label57.Text = "mm,"
        Me.Label57.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(232, 112)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(56, 23)
        Me.Button4.TabIndex = 42
        Me.Button4.Text = "HideROI"
        '
        'Button_Test
        '
        Me.Button_Test.Location = New System.Drawing.Point(344, 128)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(112, 48)
        Me.Button_Test.TabIndex = 2
        Me.Button_Test.Text = "Test"
        '
        'Label40
        '
        Me.Label40.Location = New System.Drawing.Point(8, 72)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(80, 24)
        Me.Label40.TabIndex = 23
        Me.Label40.Text = "Position"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(83, 112)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(24, 23)
        Me.Label28.TabIndex = 33
        Me.Label28.Text = "?"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(24, 144)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(232, 32)
        Me.Button6.TabIndex = 43
        Me.Button6.Text = "Reset Chip Edge Points"
        '
        'GroupBox_Imaging
        '
        Me.GroupBox_Imaging.Controls.Add(Me.Label58)
        Me.GroupBox_Imaging.Controls.Add(Me.NumericUpDown_ROI)
        Me.GroupBox_Imaging.Controls.Add(Me.Label59)
        Me.GroupBox_Imaging.Controls.Add(Me.NumericUpDown_Contrast)
        Me.GroupBox_Imaging.Controls.Add(Me.Label60)
        Me.GroupBox_Imaging.Controls.Add(Me.Label61)
        Me.GroupBox_Imaging.Controls.Add(Me.Label62)
        Me.GroupBox_Imaging.Controls.Add(Me.NumericUpDown_Brightness)
        Me.GroupBox_Imaging.Controls.Add(Me.NumericUpDown_Threshold)
        Me.GroupBox_Imaging.Controls.Add(Me.GroupBox10)
        Me.GroupBox_Imaging.Controls.Add(Me.GroupBox_Vertical_Horizontal)
        Me.GroupBox_Imaging.Location = New System.Drawing.Point(8, 208)
        Me.GroupBox_Imaging.Name = "GroupBox_Imaging"
        Me.GroupBox_Imaging.Size = New System.Drawing.Size(466, 192)
        Me.GroupBox_Imaging.TabIndex = 2
        Me.GroupBox_Imaging.TabStop = False
        Me.GroupBox_Imaging.Text = "Imaging"
        '
        'Label58
        '
        Me.Label58.Location = New System.Drawing.Point(416, 152)
        Me.Label58.Name = "Label58"
        Me.Label58.Size = New System.Drawing.Size(40, 23)
        Me.Label58.TabIndex = 45
        Me.Label58.Text = "mm"
        Me.Label58.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumericUpDown_ROI
        '
        Me.NumericUpDown_ROI.DecimalPlaces = 1
        Me.NumericUpDown_ROI.Increment = New Decimal(New Integer() {1, 0, 0, 65536})
        Me.NumericUpDown_ROI.Location = New System.Drawing.Point(360, 152)
        Me.NumericUpDown_ROI.Maximum = New Decimal(New Integer() {3, 0, 0, 0})
        Me.NumericUpDown_ROI.Name = "NumericUpDown_ROI"
        Me.NumericUpDown_ROI.Size = New System.Drawing.Size(56, 27)
        Me.NumericUpDown_ROI.TabIndex = 44
        Me.NumericUpDown_ROI.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_ROI.Value = New Decimal(New Integer() {20, 0, 0, 65536})
        '
        'Label59
        '
        Me.Label59.Location = New System.Drawing.Point(352, 128)
        Me.Label59.Name = "Label59"
        Me.Label59.Size = New System.Drawing.Size(80, 23)
        Me.Label59.TabIndex = 43
        Me.Label59.Text = "ROI size:"
        Me.Label59.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumericUpDown_Contrast
        '
        Me.NumericUpDown_Contrast.Location = New System.Drawing.Point(248, 152)
        Me.NumericUpDown_Contrast.Name = "NumericUpDown_Contrast"
        Me.NumericUpDown_Contrast.ReadOnly = True
        Me.NumericUpDown_Contrast.Size = New System.Drawing.Size(64, 27)
        Me.NumericUpDown_Contrast.TabIndex = 41
        Me.NumericUpDown_Contrast.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Contrast.Value = New Decimal(New Integer() {55, 0, 0, 0})
        '
        'Label60
        '
        Me.Label60.Location = New System.Drawing.Point(240, 128)
        Me.Label60.Name = "Label60"
        Me.Label60.Size = New System.Drawing.Size(88, 23)
        Me.Label60.TabIndex = 40
        Me.Label60.Text = "Contrast:"
        Me.Label60.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label61
        '
        Me.Label61.Location = New System.Drawing.Point(120, 128)
        Me.Label61.Name = "Label61"
        Me.Label61.Size = New System.Drawing.Size(96, 23)
        Me.Label61.TabIndex = 39
        Me.Label61.Text = "Brightness:"
        Me.Label61.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label62
        '
        Me.Label62.Location = New System.Drawing.Point(16, 128)
        Me.Label62.Name = "Label62"
        Me.Label62.Size = New System.Drawing.Size(88, 23)
        Me.Label62.TabIndex = 38
        Me.Label62.Text = "Threshold:"
        Me.Label62.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'NumericUpDown_Brightness
        '
        Me.NumericUpDown_Brightness.Location = New System.Drawing.Point(136, 152)
        Me.NumericUpDown_Brightness.Name = "NumericUpDown_Brightness"
        Me.NumericUpDown_Brightness.ReadOnly = True
        Me.NumericUpDown_Brightness.Size = New System.Drawing.Size(64, 27)
        Me.NumericUpDown_Brightness.TabIndex = 37
        Me.NumericUpDown_Brightness.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Brightness.Value = New Decimal(New Integer() {6, 0, 0, 0})
        '
        'NumericUpDown_Threshold
        '
        Me.NumericUpDown_Threshold.Location = New System.Drawing.Point(32, 152)
        Me.NumericUpDown_Threshold.Name = "NumericUpDown_Threshold"
        Me.NumericUpDown_Threshold.Size = New System.Drawing.Size(56, 27)
        Me.NumericUpDown_Threshold.TabIndex = 36
        Me.NumericUpDown_Threshold.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.NumericUpDown_Threshold.Value = New Decimal(New Integer() {15, 0, 0, 0})
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.RadioButton_Outside_In)
        Me.GroupBox10.Controls.Add(Me.RadioButton_Inside_out)
        Me.GroupBox10.Location = New System.Drawing.Point(8, 24)
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.Size = New System.Drawing.Size(448, 48)
        Me.GroupBox10.TabIndex = 46
        Me.GroupBox10.TabStop = False
        Me.GroupBox10.Text = "Finding Direction"
        '
        'RadioButton_Outside_In
        '
        Me.RadioButton_Outside_In.Location = New System.Drawing.Point(248, 16)
        Me.RadioButton_Outside_In.Name = "RadioButton_Outside_In"
        Me.RadioButton_Outside_In.Size = New System.Drawing.Size(144, 24)
        Me.RadioButton_Outside_In.TabIndex = 14
        Me.RadioButton_Outside_In.Text = "Outside to In"
        '
        'RadioButton_Inside_out
        '
        Me.RadioButton_Inside_out.Checked = True
        Me.RadioButton_Inside_out.Location = New System.Drawing.Point(88, 16)
        Me.RadioButton_Inside_out.Name = "RadioButton_Inside_out"
        Me.RadioButton_Inside_out.Size = New System.Drawing.Size(136, 24)
        Me.RadioButton_Inside_out.TabIndex = 13
        Me.RadioButton_Inside_out.TabStop = True
        Me.RadioButton_Inside_out.Text = "Inside to Out"
        '
        'GroupBox_Vertical_Horizontal
        '
        Me.GroupBox_Vertical_Horizontal.Controls.Add(Me.RadioButton_Horizontal)
        Me.GroupBox_Vertical_Horizontal.Controls.Add(Me.RadioButton_Verical)
        Me.GroupBox_Vertical_Horizontal.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox_Vertical_Horizontal.Name = "GroupBox_Vertical_Horizontal"
        Me.GroupBox_Vertical_Horizontal.Size = New System.Drawing.Size(448, 48)
        Me.GroupBox_Vertical_Horizontal.TabIndex = 47
        Me.GroupBox_Vertical_Horizontal.TabStop = False
        Me.GroupBox_Vertical_Horizontal.Text = "Vertical/Horizontal:"
        '
        'RadioButton_Horizontal
        '
        Me.RadioButton_Horizontal.Location = New System.Drawing.Point(248, 16)
        Me.RadioButton_Horizontal.Name = "RadioButton_Horizontal"
        Me.RadioButton_Horizontal.TabIndex = 67
        Me.RadioButton_Horizontal.Text = "Horizontal"
        '
        'RadioButton_Verical
        '
        Me.RadioButton_Verical.Checked = True
        Me.RadioButton_Verical.Location = New System.Drawing.Point(88, 16)
        Me.RadioButton_Verical.Name = "RadioButton_Verical"
        Me.RadioButton_Verical.TabIndex = 66
        Me.RadioButton_Verical.TabStop = True
        Me.RadioButton_Verical.Text = "Vertical"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Panel2)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 40
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'PanelToBeAdded2
        '
        Me.PanelToBeAdded2.Controls.Add(Me.Label20)
        Me.PanelToBeAdded2.Controls.Add(Me.Panel3)
        Me.PanelToBeAdded2.Location = New System.Drawing.Point(560, 8)
        Me.PanelToBeAdded2.Name = "PanelToBeAdded2"
        Me.PanelToBeAdded2.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded2.TabIndex = 40
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label20.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label20.Location = New System.Drawing.Point(0, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(296, 32)
        Me.Label20.TabIndex = 46
        Me.Label20.Text = "Camera Calibration Setup"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.GroupBox1)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel3.Location = New System.Drawing.Point(0, 80)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(528, 864)
        Me.Panel3.TabIndex = 44
        '
        'VisionSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1088, 968)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Controls.Add(Me.PanelToBeAdded2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "VisionSetup"
        Me.Text = "SystemSetup"
        Me.GroupBox12.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        CType(Me.NumericUpDown_Spacing, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_NoOfRows, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_NoOfColumns, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox_RefBlkRect.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        CType(Me.NumericUpDown_Rot, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Pos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Size, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_Imaging.ResumeLayout(False)
        CType(Me.NumericUpDown_ROI, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Contrast, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Brightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_Threshold, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox_Vertical_Horizontal.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.PanelToBeAdded2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
     

#Region "Global Variables"
    Dim NoOfColumns As Integer = 10
    Dim NoOfRows As Integer = 8
    Dim Spacing As Double = 1.5 * 1000
    Dim Clickno As Integer = 1
    Dim Test_Click As Boolean = False
    Dim Inside_out As Boolean = True
    Dim PointX1a, PointY1a, PointX2a, PointY2a, PointX3a, PointY3a, PointX4a, PointY4a, PointX5a, PointY5a As Double
    Dim MainEdge As Integer = 1
    Dim Vertical As Boolean
    Dim P1x, P1y, P2x, P2y, P3x, P3y, P4x, P4y, P5x, P5y As Double
#End Region
#Region "Reset"
    Sub ResetVariables()
        'Reset all variables
        TextBox_SizeX.Text = 0
        TextBox_SizeY.Text = 0
        TextBox_PosX.Text = 0
        TextBox_PosY.Text = 0
        NumericUpDown_Size.Value = 5.0
        NumericUpDown_Pos.Value = 5.0
        NumericUpDown_Rot.Text = 5
        RadioButton_Inside_out.Checked = True
        RadioButton_Outside_In.Checked = False
        RadioButton_Verical.Checked = True
        RadioButton_Horizontal.Checked = False
        NumericUpDown_Threshold.Value = 15
        NumericUpDown_Brightness.Value = 6
        NumericUpDown_ROI.Value = 2
        TextBox_RSizeX.Text = 0
        TextBox_RSizeY.Text = 0
        TextBox_RPosX.Text = 0
        TextBox_RPosY.Text = 0
        TextBox_RRot.Text = 0
        TextBox_Score.Text = 0
        'xu frm.ResetChipEdgePoint()
        '''IDS.Devices.Vision.FrmVision.ResetChipEdgePoint() 'sh

        Timer1.Stop()
        Button_Test.Text = "Test"
        'xu frm.AxDisplay1.OverlayImage.Clear()
        '''IDS.Devices.Vision.FrmVision.ClearDisplay() 'sh

        Clickno = 1
        Test_Click = False
    End Sub
    Sub ResetVariables_Back()
        'Reset all variables
        TextBox_SizeX.Text = 0
        TextBox_SizeY.Text = 0
        TextBox_PosX.Text = 0
        TextBox_PosY.Text = 0
        NumericUpDown_Size.Value = 5.0
        NumericUpDown_Pos.Value = 5.0
        NumericUpDown_Rot.Text = 5.0
        RadioButton_Inside_out.Checked = True
        RadioButton_Outside_In.Checked = False
        RadioButton_Verical.Checked = True
        RadioButton_Horizontal.Checked = False
        NumericUpDown_Threshold.Value = 15
        NumericUpDown_Brightness.Value = 6
        NumericUpDown_ROI.Value = 2
        TextBox_RSizeX.Text = 0
        TextBox_RSizeY.Text = 0
        TextBox_RPosX.Text = 0
        TextBox_RPosY.Text = 0
        TextBox_RRot.Text = 0
        TextBox_Score.Text = 0
        'xu frm.ResetChipEdgePoint()
        '''IDS.Devices.Vision.FrmVision.ResetChipEdgePoint() 'sh

        Timer1.Stop()
        Button_Test.Text = "Test"
        'frm.AxDisplay1.OverlayImage.Clear()
        '''IDS.Devices.Vision.FrmVision.ClearDisplay()

        Clickno = 1
        Test_Click = False
    End Sub
#End Region
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
#Region "Production"
    'Function ChipEdgePointExe()
    'frm.AxDisplay1.OverlayImage.Clear()

    'frm.MeasurementPoint(NumericUpDown_Contrast.Value, NumericUpDown_Threshold.Value, NumericUpDown_Rot.Value, Inside_out, Vertical, True)
    'frm.SearchRegionResultsPoints(3)
    'frm.SearchRegionPoints(3)
    'frm.ChipPointOutPut(3)
    'MsgBox(PointX1a.ToString & " " & PointY1a.ToString & " " & PointX2a.ToString & " " & PointY2a.ToString & " " & PointX3a.ToString & " " & PointY3a.ToString & " " & PointX4a.ToString & " " & PointY4a.ToString)
    'ResetVariables()
    'Me.Visible = False
    'End Function
    'Function GetVariables(ByVal FrmInside_out, ByVal FrmVertical)
    'Inside_out = FrmInside_out
    'Vertical = FrmVertical
    'End Function
    Function CheckFileName(ByVal FileName) As Boolean
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
        NewFullPath = "c:\iDS\SRC\DLL Export Device Vision\ss" & "\" & FileName & "." & FileExt
        Flag = False

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
#End Region 'unwanted
    Dim BlackDot As Boolean = False
    'Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'Me.DialogResult = DialogResult.OK
    'frm.modelregionDrawing()
    'frm.TestDot(BlackDot, NumericUpDown_Binarized.Value, NumericUpDown_Close.Value, NumericUpDown_Open.Value, NumericUpDown_MinArea.Value, NumericUpDown_MaxArea.Value, NumericUpDown_Roughness.Value, NumericUpDown_Compactness.Value)
    'End Sub
    'Private Sub RadioButton_BlackDot_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    'If RadioButton_WhiteDot.Checked = True Then
    'BlackDot = False
    'ElseIf RadioButton_BlackDot.Checked = True Then
    'BlackDot = True
    'End If
    'End Sub
    Private Sub ChipEdgePoints_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Visible = True
        ''TextBox_PixelX.Text = IDS.Devices.Vision.FrmVision.PixelXSize
        ''TextBox_PixelY.Text = frm.PixelYSize
        ''TextBox_Ratio.Text = frm.Pixel_Ratio
        TextBox_PixelX.Text = IDS.Devices.Vision.FrmVision.PixelXSize
        TextBox_PixelY.Text = IDS.Devices.Vision.FrmVision.PixelYSize
        TextBox_Ratio.Text = IDS.Devices.Vision.FrmVision.Pixel_Ratio
    End Sub
    Private Sub NumericUpDown_NoofColumns_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown_NoOfColumns.ValueChanged
        NoOfColumns = NumericUpDown_NoOfColumns.Value
    End Sub
    Private Sub NumericUpDown_NoOfRows_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown_NoOfRows.ValueChanged
        NoOfRows = NumericUpDown_NoOfRows.Value
    End Sub
    Private Sub NumericUpDown_Spacing_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown_Spacing.ValueChanged
        Spacing = NumericUpDown_Spacing.Value
    End Sub
    Private Sub Button_Calibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Calibrate.Click
        ''Dim Errors = frm.CameraCalibration(NoOfColumns, NoOfRows, Spacing)
        ''If Errors = -1 Then
        ''MsgBox("Failed")
        ''Else
        ''TextBox_PixelX.Text = frm.PixelXSize
        ''TextBox_PixelY.Text = frm.PixelYSize
        ''TextBox_Ratio.Text = frm.Pixel_Ratio
        ''End If
        IDS.Devices.Vision.FrmVision.ClearDisplay()
        NoOfColumns = NumericUpDown_NoOfColumns.Value
        NoOfRows = NumericUpDown_NoOfRows.Value
        Spacing = NumericUpDown_Spacing.Value
        Dim Errors = IDS.Devices.Vision.FrmVision.CameraCalibration(NoOfColumns, NoOfRows, Spacing)
        If Errors = -1 Then
            MsgBox("Failed")
        Else
            TextBox_PixelX.Text = IDS.Devices.Vision.FrmVision.PixelXSize
            TextBox_PixelY.Text = IDS.Devices.Vision.FrmVision.PixelYSize
            TextBox_Ratio.Text = IDS.Devices.Vision.FrmVision.Pixel_Ratio
        End If
    End Sub
    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ''frm.SS_Set()
        GroupBox1.BringToFront()
    End Sub
    Private Sub ButtonBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBack.Click
        Panel2.Controls.Remove(GroupBox1)
        'GroupBox1.BringToFront()
        ResetVariables_Back()
        IDS.Devices.Vision.FrmVision.DisableChipEdgeDrawing()
    End Sub
    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        GroupBox3.BringToFront()
    End Sub
    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click


        IDS.Data.Hardware.Camera.BlockSize.X = TextBox_SizeX.Text
        IDS.Data.Hardware.Camera.BlockSize.Y = TextBox_SizeY.Text
        m_Tri.GetIDSState()
        IDS.Data.Hardware.Camera.ReferencePos.X = m_Tri.XPosition
        IDS.Data.Hardware.Camera.ReferencePos.Y = m_Tri.YPosition

        'Manual calculation
        'IDS.Data.Hardware.Camera.ReferencePos.X = 127.271
        'IDS.Data.Hardware.Camera.ReferencePos.Y = -360.832

        IDS.Data.Hardware.Camera.BlockPosTolerance = NumericUpDown_Pos.Value
        IDS.Data.Hardware.Camera.BlockSizeTolerance = NumericUpDown_Size.Value
        IDS.Data.Hardware.Camera.BlockRotationTolerance = NumericUpDown_Rot.Value

        IDS.Data.Hardware.Camera.ImageFindDir = RadioButton_Inside_out.Checked '1 = in->out,  0 = out->in
        IDS.Data.Hardware.Camera.ImageVertical = RadioButton_Verical.Checked
        IDS.Data.Hardware.Camera.Threshold = NumericUpDown_Threshold.Value
        IDS.Data.Hardware.Camera.Brightness = NumericUpDown_Brightness.Value
        IDS.Data.Hardware.Camera.Contrast = NumericUpDown_Contrast.Value
        IDS.Data.Hardware.Camera.ROISize = NumericUpDown_ROI.Value

        IDS.Data.Hardware.Camera.ResultSize.X = TextBox_RSizeX.Text
        IDS.Data.Hardware.Camera.ResultSize.Y = TextBox_RSizeY.Text
        IDS.Data.Hardware.Camera.ResultDia = TextBox_RDia.Text
        'IDS.Data.Hardware.Camera.Resultpos.x = TextBox_RSizeX.Text
        'IDS.Data.Hardware.Camera.Resultpos.y = TextBox_RSizeX.Text
        IDS.Data.Hardware.Camera.ResultRotation = TextBox_RRot.Text

        IDS.Data.SaveData()
        'ResetVariables()
    End Sub
    Private Sub Button_Test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Test.Click
        Dim PointNo = IDS.Devices.Vision.FrmVision.PointNo
        If PointNo <> 5 Then
            MsgBox("Please click 5 points on the display")
        Else
            Dim SS(10) As Double
            SS = IDS.Devices.Vision.FrmVision.SS_ChipParrameter
            TextBox_SizeX.Text = SS(0)
            TextBox_SizeY.Text = SS(1)
            TextBox_PosX.Text = SS(2)
            TextBox_PosY.Text = SS(3)
            If Test_Click = False Then
                Timer1.Start()
                Button_Test.Text = "Stop"
                Test_Click = True
            ElseIf Test_Click = True Then
                Timer1.Stop()
                Test_Click = False
                Button_Test.Text = "Test"
            End If
        End If
    End Sub
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'IDS.Devices.Vision.FrmVision.ClearDisplay()
        IDS.Devices.Vision.FrmVision.SetupVariable(TextBox_SizeX, TextBox_SizeY, TextBox_PosX, TextBox_PosY, NumericUpDown_Size, NumericUpDown_Pos, NumericUpDown_Rot) 'link those textbox to vision's textbox
        IDS.Devices.Vision.FrmVision.SetupRVariable(TextBox_RSizeX, TextBox_RSizeY, TextBox_RPosX, TextBox_RPosY, TextBox_RRot, TextBox_Score)

        Panel2.Controls.Add(Me.GroupBox1)
        GroupBox1.BringToFront()
        GroupBox1.Location = New Point(8, 56) '(8,72)?
        IDS.Devices.Vision.FrmVision.EnableChipEdgeDrawing()
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If Clickno = 1 Then
            Clickno = 0
        ElseIf Clickno = 0 Then
            Clickno = 1
        End If
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ResetVariables()
        Me.Visible = False
        SetTable_a()
    End Sub
    Private Sub RadioButton_Verical_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Verical.CheckedChanged
        Vertical = True
    End Sub
    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Horizontal.CheckedChanged
        Vertical = False
    End Sub
    Private Sub RadioButton_Inside_out_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Inside_out.CheckedChanged
        Inside_out = True
    End Sub
    Private Sub RadioButton_Outside_In_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton_Outside_In.CheckedChanged
        Inside_out = False
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ''frm.AxDisplay1.OverlayImage.Clear()
        ''frm.MeasurementPoint(NumericUpDown_Contrast.Value, NumericUpDown_Threshold.Value, NumericUpDown_Rot.Value, Inside_out, Vertical, 3)
        IDS.Devices.Vision.FrmVision.ClearDisplay()
        IDS.Devices.Vision.FrmVision.MeasurementPoint(NumericUpDown_Contrast.Value, NumericUpDown_Threshold.Value, NumericUpDown_Rot.Value, Inside_out, Vertical, NumericUpDown_ROI.Value, 3)
        Dim ss(10) As Double
        'ss = IDS.Devices.Vision.FrmVision.SS_Results
        'TextBox_RSizeX.Text = ss(0)
        'TextBox_RSizeY.Text = ss(1)
        'TextBox_RPosX.Text = ss(2)
        'TextBox_RPosY.Text = ss(3)
        'TextBox_RRot.Text = ss(4)
        If Clickno = 1 Then
            'IDS.Devices.Vision.FrmVision.SearchRegionResultsPoints(3)
            IDS.Devices.Vision.FrmVision.SearchRegionPoints(3, NumericUpDown_ROI.Value)
        End If
        If Clickno = 1 Then
            ''frm.SearchRegionResultsPoints(3)
            ''frm.SearchRegionPoints(3)
        End If
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click '
        RemovePanel(CurrentControl)
        Timer1.Stop()
        IDS.Devices.Vision.FrmVision.ClearDisplay()
        IDS.Devices.Vision.FrmVision.DisplayIndicator()
        IDS.Devices.Vision.FrmVision.DisableChipEdgeDrawing()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SetTable_b()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Para.RecordID = "FactoryDefault"
        'Para.GroupID = IDS.Data.Admin.User.Group.ID
        'IDS.Data.OpenData()

        IDS.Data.Hardware.Camera.CalibrationNumRows = NumericUpDown_NoOfRows.Value
        IDS.Data.Hardware.Camera.CalibrationNumColumns = NumericUpDown_NoOfColumns.Value
        IDS.Data.Hardware.Camera.CalibrationSpacing = NumericUpDown_Spacing.Value '* 1000 '*1000

        IDS.Data.Hardware.Camera.PixelsToMM.X = TextBox_PixelX.Text
        IDS.Data.Hardware.Camera.PixelsToMM.Y = TextBox_PixelY.Text
        IDS.Data.Hardware.Camera.PixelRatio = TextBox_Ratio.Text

        IDS.Data.SaveData()
    End Sub

    Private Sub RevertData()
        IDS.Data.OpenData()
        Para.RecordID = "FactoryDefault"
        Para.GroupID = IDS.Data.Admin.User.Group.ID
        IDS.Data.OpenData()

        NumericUpDown_NoOfRows.Text = IDS.Data.Hardware.Camera.CalibrationNumRows
        NumericUpDown_NoOfColumns.Text = IDS.Data.Hardware.Camera.CalibrationNumColumns
        NumericUpDown_Spacing.Text = IDS.Data.Hardware.Camera.CalibrationSpacing

        TextBox_SizeX.Text = IDS.Data.Hardware.Camera.BlockSize.X
        TextBox_SizeY.Text = IDS.Data.Hardware.Camera.BlockSize.Y
        TextBox_PosX.Text = IDS.Data.Hardware.Camera.ReferencePos.X
        TextBox_PosY.Text = IDS.Data.Hardware.Camera.ReferencePos.Y
        NumericUpDown_Pos.Text = IDS.Data.Hardware.Camera.BlockPosTolerance
        NumericUpDown_Size.Text = IDS.Data.Hardware.Camera.BlockSizeTolerance
        NumericUpDown_Rot.Text = IDS.Data.Hardware.Camera.BlockRotationTolerance

        RadioButton_Outside_In.Checked = True
        RadioButton_Horizontal.Checked = True
        RadioButton_Inside_out.Checked = IDS.Data.Hardware.Camera.ImageFindDir
        RadioButton_Verical.Checked = IDS.Data.Hardware.Camera.ImageVertical
        NumericUpDown_Threshold.Value = IDS.Data.Hardware.Camera.Threshold
        NumericUpDown_Brightness.Value = IDS.Data.Hardware.Camera.Brightness
        NumericUpDown_Contrast.Value = IDS.Data.Hardware.Camera.Contrast
        NumericUpDown_ROI.Value = IDS.Data.Hardware.Camera.ROISize

    End Sub

    Public Sub SetTable_a() 'Cancel button 1
        'Para.RecordID = "FactoryDefault"
        'Para.GroupID = IDS.Data.Admin.User.Group.ID
        'IDS.Data.OpenData()

        NumericUpDown_NoOfRows.Text = IDS.Data.Hardware.Camera.CalibrationNumRows
        NumericUpDown_NoOfColumns.Text = IDS.Data.Hardware.Camera.CalibrationNumColumns
        NumericUpDown_Spacing.Text = IDS.Data.Hardware.Camera.CalibrationSpacing '/1000

    End Sub

    Public Sub SetTable_b() 'Cancel button 2
        'Para.RecordID = "FactoryDefault"
        'Para.GroupID = IDS.Data.Admin.User.Group.ID
        'IDS.Data.OpenData()

        TextBox_SizeX.Text = IDS.Data.Hardware.Camera.BlockSize.X
        TextBox_SizeY.Text = IDS.Data.Hardware.Camera.BlockSize.Y
        TextBox_PosX.Text = IDS.Data.Hardware.Camera.ReferencePos.X
        TextBox_PosY.Text = IDS.Data.Hardware.Camera.ReferencePos.Y
        NumericUpDown_Pos.Text = IDS.Data.Hardware.Camera.BlockPosTolerance
        NumericUpDown_Size.Text = IDS.Data.Hardware.Camera.BlockSizeTolerance
        NumericUpDown_Rot.Text = IDS.Data.Hardware.Camera.BlockRotationTolerance

        RadioButton_Outside_In.Checked = True
        RadioButton_Horizontal.Checked = True
        RadioButton_Inside_out.Checked = IDS.Data.Hardware.Camera.ImageFindDir
        RadioButton_Verical.Checked = IDS.Data.Hardware.Camera.ImageVertical
        NumericUpDown_Threshold.Value = IDS.Data.Hardware.Camera.Threshold
        NumericUpDown_Brightness.Value = IDS.Data.Hardware.Camera.Brightness
        NumericUpDown_Contrast.Value = IDS.Data.Hardware.Camera.Contrast
        NumericUpDown_ROI.Value = IDS.Data.Hardware.Camera.ROISize

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        IDS.Devices.Vision.FrmVision.ResetChipEdgePoint()
        IDS.Devices.Vision.FrmVision.EnableChipEdgeDrawing()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        IDS.Devices.Vision.FrmVision.DisableChipEdgeDrawing()
    End Sub

    Private Sub btGetReferenceCenter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGetReferenceCenter.Click
        Dim post(1) As Double
        Dim currentPos(1) As Double
        currentPos(0) = m_Tri.XPosition
        currentPos(1) = m_Tri.YPosition
        IDS.Devices.Vision.FrmVision.GetChipCenter_World(currentPos, post(0), post(1))
        rtbTips.Clear()
        rtbTips.AppendText("Reference Point X:" & post(0) & " Y:" & post(1))
        Dim relPost(1) As Double
        relPost(0) = 0
        relPost(1) = 0
        IDS.Data.Hardware.Camera.ReferencePos.X = post(0)
        IDS.Data.Hardware.Camera.ReferencePos.Y = post(1)
        m_Tri.MoveRelative_XY(relPost)
        m_Tri.Move_XY(post)
    End Sub
End Class
