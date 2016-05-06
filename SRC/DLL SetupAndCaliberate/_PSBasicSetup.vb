Imports DLL_DataManager
Imports DLL_Export_Service



Public Class PSBasicSetup
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Lbl_ConveyorCurrentWidth.Text = IDS.Data.HardWare.Conveyor.Width
        'Lbl_DefaultElementTravelSpeed.Text=
        ElementTravelSpeedBar.Maximum = IDS.Data.HardWare.Gantry.MaxSpeedLimit
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Btn_ConveyorStart As System.Windows.Forms.Button
    Friend WithEvents Btn_ConveyorHome As System.Windows.Forms.Button
    Friend WithEvents ConveyorFineAdjust As System.Windows.Forms.NumericUpDown
    Friend WithEvents Btn_SoftHomeTrack As System.Windows.Forms.Button
    Friend WithEvents CB_SoftHomeUseDefault As System.Windows.Forms.CheckBox
    Friend WithEvents CB_DipenserLeftHead As System.Windows.Forms.CheckBox
    Friend WithEvents CB_DipenserRightHead As System.Windows.Forms.CheckBox
    Friend WithEvents Btn_ApplyElementTravelSpeed As System.Windows.Forms.Button
    Friend WithEvents ConveyorNewWidth As System.Windows.Forms.NumericUpDown
    Friend WithEvents NewSoftHomeX As System.Windows.Forms.NumericUpDown
    Friend WithEvents NewSoftHomeY As System.Windows.Forms.NumericUpDown
    Friend WithEvents NewSoftHomeZ As System.Windows.Forms.NumericUpDown
    Friend WithEvents NewElementTravelSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents GraphicBotRightX As System.Windows.Forms.NumericUpDown
    Friend WithEvents GraphicBotRightY As System.Windows.Forms.NumericUpDown
    Friend WithEvents GraphicTopLeftX As System.Windows.Forms.NumericUpDown
    Friend WithEvents GraphicTopLeftY As System.Windows.Forms.NumericUpDown
    Friend WithEvents Btn_WidthForward As System.Windows.Forms.Button
    Friend WithEvents Btn_WidthReverse As System.Windows.Forms.Button
    Friend WithEvents Btn_GraphicsBotRight As System.Windows.Forms.Button
    Friend WithEvents Btn_GraphicsTopLeft As System.Windows.Forms.Button
    Friend WithEvents CB_LifterUp As System.Windows.Forms.CheckBox
    Friend WithEvents CB_ThermalRequired As System.Windows.Forms.CheckBox
    Friend WithEvents CB_ReleaseBoard As System.Windows.Forms.CheckBox
    Friend WithEvents Btn_CancelBasicSetup As System.Windows.Forms.Button
    Friend WithEvents Btn_ApplyBasicSetup As System.Windows.Forms.Button
    Friend WithEvents Cb_SelectMeasurementUnit As System.Windows.Forms.ComboBox
    Friend WithEvents CB_LeftHeadType As System.Windows.Forms.ComboBox
    Friend WithEvents CB_RightHeadType As System.Windows.Forms.ComboBox
    Friend WithEvents Lbl_DefaultElementTravelSpeed As System.Windows.Forms.Label
    Friend WithEvents Lbl_ConveyorCurrentWidth As System.Windows.Forms.Label
    Friend WithEvents ElementTravelSpeedBar As System.Windows.Forms.TrackBar
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents DomainUpDown3 As System.Windows.Forms.DomainUpDown
    Friend WithEvents LabelStep As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton7 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton8 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton9 As System.Windows.Forms.RadioButton
    Public WithEvents PanelLeft As System.Windows.Forms.Panel
    Public WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Public WithEvents PanelRight As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(PSBasicSetup))
        Me.Cb_SelectMeasurementUnit = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ConveyorNewWidth = New System.Windows.Forms.NumericUpDown
        Me.Btn_WidthForward = New System.Windows.Forms.Button
        Me.Btn_WidthReverse = New System.Windows.Forms.Button
        Me.Btn_ConveyorStart = New System.Windows.Forms.Button
        Me.Btn_ConveyorHome = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Lbl_ConveyorCurrentWidth = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.ConveyorFineAdjust = New System.Windows.Forms.NumericUpDown
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.GraphicBotRightX = New System.Windows.Forms.NumericUpDown
        Me.GraphicBotRightY = New System.Windows.Forms.NumericUpDown
        Me.GraphicTopLeftX = New System.Windows.Forms.NumericUpDown
        Me.GraphicTopLeftY = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Btn_GraphicsBotRight = New System.Windows.Forms.Button
        Me.Btn_GraphicsTopLeft = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.CB_LifterUp = New System.Windows.Forms.CheckBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.CB_ThermalRequired = New System.Windows.Forms.CheckBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.CB_ReleaseBoard = New System.Windows.Forms.CheckBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.NewSoftHomeZ = New System.Windows.Forms.NumericUpDown
        Me.NewSoftHomeX = New System.Windows.Forms.NumericUpDown
        Me.NewSoftHomeY = New System.Windows.Forms.NumericUpDown
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Btn_SoftHomeTrack = New System.Windows.Forms.Button
        Me.Label15 = New System.Windows.Forms.Label
        Me.CB_SoftHomeUseDefault = New System.Windows.Forms.CheckBox
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.CB_DipenserLeftHead = New System.Windows.Forms.CheckBox
        Me.CB_LeftHeadType = New System.Windows.Forms.ComboBox
        Me.CB_DipenserRightHead = New System.Windows.Forms.CheckBox
        Me.CB_RightHeadType = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.NewElementTravelSpeed = New System.Windows.Forms.NumericUpDown
        Me.ElementTravelSpeedBar = New System.Windows.Forms.TrackBar
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Lbl_DefaultElementTravelSpeed = New System.Windows.Forms.Label
        Me.Btn_ApplyElementTravelSpeed = New System.Windows.Forms.Button
        Me.Btn_CancelBasicSetup = New System.Windows.Forms.Button
        Me.Btn_ApplyBasicSetup = New System.Windows.Forms.Button
        Me.PanelRight = New System.Windows.Forms.Panel
        Me.Label17 = New System.Windows.Forms.Label
        Me.PanelLeft = New System.Windows.Forms.Panel
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.DomainUpDown3 = New System.Windows.Forms.DomainUpDown
        Me.LabelStep = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.RadioButton8 = New System.Windows.Forms.RadioButton
        Me.RadioButton9 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.RadioButton6 = New System.Windows.Forms.RadioButton
        Me.RadioButton5 = New System.Windows.Forms.RadioButton
        Me.RadioButton7 = New System.Windows.Forms.RadioButton
        Me.RadioButton4 = New System.Windows.Forms.RadioButton
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton3 = New System.Windows.Forms.RadioButton
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.ConveyorNewWidth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ConveyorFineAdjust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.GraphicBotRightX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GraphicBotRightY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GraphicTopLeftX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GraphicTopLeftY, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.NewSoftHomeZ, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NewSoftHomeX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NewSoftHomeY, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        CType(Me.NewElementTravelSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ElementTravelSpeedBar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelRight.SuspendLayout()
        Me.PanelLeft.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.SuspendLayout()
        '
        'Cb_SelectMeasurementUnit
        '
        Me.Cb_SelectMeasurementUnit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Cb_SelectMeasurementUnit.Items.AddRange(New Object() {"mm", "inch"})
        Me.Cb_SelectMeasurementUnit.Location = New System.Drawing.Point(144, 48)
        Me.Cb_SelectMeasurementUnit.MaxDropDownItems = 2
        Me.Cb_SelectMeasurementUnit.Name = "Cb_SelectMeasurementUnit"
        Me.Cb_SelectMeasurementUnit.Size = New System.Drawing.Size(72, 31)
        Me.Cb_SelectMeasurementUnit.TabIndex = 0
        Me.Cb_SelectMeasurementUnit.Text = "mm"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Measurement Unit"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ConveyorNewWidth)
        Me.GroupBox1.Controls.Add(Me.Btn_WidthForward)
        Me.GroupBox1.Controls.Add(Me.Btn_WidthReverse)
        Me.GroupBox1.Controls.Add(Me.Btn_ConveyorStart)
        Me.GroupBox1.Controls.Add(Me.Btn_ConveyorHome)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Lbl_ConveyorCurrentWidth)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ConveyorFineAdjust)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(280, 216)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Conveyor Width"
        '
        'ConveyorNewWidth
        '
        Me.ConveyorNewWidth.Location = New System.Drawing.Point(64, 72)
        Me.ConveyorNewWidth.Name = "ConveyorNewWidth"
        Me.ConveyorNewWidth.Size = New System.Drawing.Size(56, 27)
        Me.ConveyorNewWidth.TabIndex = 12
        '
        'Btn_WidthForward
        '
        Me.Btn_WidthForward.Location = New System.Drawing.Point(40, 168)
        Me.Btn_WidthForward.Name = "Btn_WidthForward"
        Me.Btn_WidthForward.Size = New System.Drawing.Size(75, 32)
        Me.Btn_WidthForward.TabIndex = 11
        Me.Btn_WidthForward.Text = " <<-->>"
        '
        'Btn_WidthReverse
        '
        Me.Btn_WidthReverse.Location = New System.Drawing.Point(160, 168)
        Me.Btn_WidthReverse.Name = "Btn_WidthReverse"
        Me.Btn_WidthReverse.Size = New System.Drawing.Size(75, 32)
        Me.Btn_WidthReverse.TabIndex = 10
        Me.Btn_WidthReverse.Text = ">>-- <<"
        '
        'Btn_ConveyorStart
        '
        Me.Btn_ConveyorStart.Location = New System.Drawing.Point(128, 64)
        Me.Btn_ConveyorStart.Name = "Btn_ConveyorStart"
        Me.Btn_ConveyorStart.Size = New System.Drawing.Size(64, 40)
        Me.Btn_ConveyorStart.TabIndex = 9
        Me.Btn_ConveyorStart.Text = "Start"
        '
        'Btn_ConveyorHome
        '
        Me.Btn_ConveyorHome.Location = New System.Drawing.Point(208, 64)
        Me.Btn_ConveyorHome.Name = "Btn_ConveyorHome"
        Me.Btn_ConveyorHome.Size = New System.Drawing.Size(64, 40)
        Me.Btn_ConveyorHome.TabIndex = 8
        Me.Btn_ConveyorHome.Text = "Home"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(40, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 23)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Fine Adjustment"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(9, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 23)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Goto"
        '
        'Lbl_ConveyorCurrentWidth
        '
        Me.Lbl_ConveyorCurrentWidth.Location = New System.Drawing.Point(144, 32)
        Me.Lbl_ConveyorCurrentWidth.Name = "Lbl_ConveyorCurrentWidth"
        Me.Lbl_ConveyorCurrentWidth.Size = New System.Drawing.Size(56, 23)
        Me.Lbl_ConveyorCurrentWidth.TabIndex = 4
        Me.Lbl_ConveyorCurrentWidth.Text = "200"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Current Width"
        '
        'ConveyorFineAdjust
        '
        Me.ConveyorFineAdjust.Location = New System.Drawing.Point(176, 128)
        Me.ConveyorFineAdjust.Name = "ConveyorFineAdjust"
        Me.ConveyorFineAdjust.Size = New System.Drawing.Size(56, 27)
        Me.ConveyorFineAdjust.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GraphicBotRightX)
        Me.GroupBox2.Controls.Add(Me.GraphicBotRightY)
        Me.GroupBox2.Controls.Add(Me.GraphicTopLeftX)
        Me.GroupBox2.Controls.Add(Me.GraphicTopLeftY)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Btn_GraphicsBotRight)
        Me.GroupBox2.Controls.Add(Me.Btn_GraphicsTopLeft)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.PictureBox1)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 328)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(280, 416)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Graphic Display Dimension"
        '
        'GraphicBotRightX
        '
        Me.GraphicBotRightX.Location = New System.Drawing.Point(192, 288)
        Me.GraphicBotRightX.Name = "GraphicBotRightX"
        Me.GraphicBotRightX.Size = New System.Drawing.Size(56, 27)
        Me.GraphicBotRightX.TabIndex = 20
        '
        'GraphicBotRightY
        '
        Me.GraphicBotRightY.Location = New System.Drawing.Point(192, 320)
        Me.GraphicBotRightY.Name = "GraphicBotRightY"
        Me.GraphicBotRightY.Size = New System.Drawing.Size(56, 27)
        Me.GraphicBotRightY.TabIndex = 19
        Me.GraphicBotRightY.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'GraphicTopLeftX
        '
        Me.GraphicTopLeftX.Location = New System.Drawing.Point(56, 288)
        Me.GraphicTopLeftX.Name = "GraphicTopLeftX"
        Me.GraphicTopLeftX.Size = New System.Drawing.Size(56, 27)
        Me.GraphicTopLeftX.TabIndex = 18
        '
        'GraphicTopLeftY
        '
        Me.GraphicTopLeftY.Location = New System.Drawing.Point(56, 320)
        Me.GraphicTopLeftY.Name = "GraphicTopLeftY"
        Me.GraphicTopLeftY.Size = New System.Drawing.Size(56, 27)
        Me.GraphicTopLeftY.TabIndex = 17
        Me.GraphicTopLeftY.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(168, 320)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(15, 23)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Y"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(32, 320)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(15, 23)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "Y"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(168, 288)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(15, 23)
        Me.Label10.TabIndex = 12
        Me.Label10.Text = "X"
        '
        'Btn_GraphicsBotRight
        '
        Me.Btn_GraphicsBotRight.Location = New System.Drawing.Point(160, 360)
        Me.Btn_GraphicsBotRight.Name = "Btn_GraphicsBotRight"
        Me.Btn_GraphicsBotRight.Size = New System.Drawing.Size(96, 40)
        Me.Btn_GraphicsBotRight.TabIndex = 9
        Me.Btn_GraphicsBotRight.Text = "TrackBall"
        '
        'Btn_GraphicsTopLeft
        '
        Me.Btn_GraphicsTopLeft.Location = New System.Drawing.Point(24, 360)
        Me.Btn_GraphicsTopLeft.Name = "Btn_GraphicsTopLeft"
        Me.Btn_GraphicsTopLeft.Size = New System.Drawing.Size(96, 40)
        Me.Btn_GraphicsTopLeft.TabIndex = 8
        Me.Btn_GraphicsTopLeft.Text = "TrackBall"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 288)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(15, 23)
        Me.Label7.TabIndex = 5
        Me.Label7.Text = "X"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(192, 248)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(56, 23)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "100 x 150"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(40, 248)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(152, 23)
        Me.Label9.TabIndex = 3
        Me.Label9.Text = "Dimension Ratio"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(8, 32)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(264, 208)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox1.TabIndex = 13
        Me.PictureBox1.TabStop = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.CB_LifterUp)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(8, 760)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(280, 56)
        Me.GroupBox3.TabIndex = 13
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Jack Up Unit"
        '
        'CB_LifterUp
        '
        Me.CB_LifterUp.Location = New System.Drawing.Point(40, 24)
        Me.CB_LifterUp.Name = "CB_LifterUp"
        Me.CB_LifterUp.Size = New System.Drawing.Size(208, 24)
        Me.CB_LifterUp.TabIndex = 0
        Me.CB_LifterUp.Text = "Lift Up In Jack Up Unit"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.CB_ThermalRequired)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(8, 832)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(280, 56)
        Me.GroupBox4.TabIndex = 14
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Thermal"
        '
        'CB_ThermalRequired
        '
        Me.CB_ThermalRequired.Location = New System.Drawing.Point(40, 24)
        Me.CB_ThermalRequired.Name = "CB_ThermalRequired"
        Me.CB_ThermalRequired.Size = New System.Drawing.Size(160, 24)
        Me.CB_ThermalRequired.TabIndex = 0
        Me.CB_ThermalRequired.Text = "Heater Required"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.CB_ReleaseBoard)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(296, 760)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(224, 56)
        Me.GroupBox5.TabIndex = 15
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Board Status"
        '
        'CB_ReleaseBoard
        '
        Me.CB_ReleaseBoard.Location = New System.Drawing.Point(40, 24)
        Me.CB_ReleaseBoard.Name = "CB_ReleaseBoard"
        Me.CB_ReleaseBoard.Size = New System.Drawing.Size(144, 24)
        Me.CB_ReleaseBoard.TabIndex = 0
        Me.CB_ReleaseBoard.Text = "Release Board"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.NewSoftHomeZ)
        Me.GroupBox6.Controls.Add(Me.NewSoftHomeX)
        Me.GroupBox6.Controls.Add(Me.NewSoftHomeY)
        Me.GroupBox6.Controls.Add(Me.Label19)
        Me.GroupBox6.Controls.Add(Me.Label14)
        Me.GroupBox6.Controls.Add(Me.Btn_SoftHomeTrack)
        Me.GroupBox6.Controls.Add(Me.Label15)
        Me.GroupBox6.Controls.Add(Me.CB_SoftHomeUseDefault)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(296, 96)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(225, 216)
        Me.GroupBox6.TabIndex = 18
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "SoftHome"
        '
        'NewSoftHomeZ
        '
        Me.NewSoftHomeZ.Location = New System.Drawing.Point(57, 168)
        Me.NewSoftHomeZ.Name = "NewSoftHomeZ"
        Me.NewSoftHomeZ.Size = New System.Drawing.Size(56, 27)
        Me.NewSoftHomeZ.TabIndex = 17
        '
        'NewSoftHomeX
        '
        Me.NewSoftHomeX.Location = New System.Drawing.Point(58, 72)
        Me.NewSoftHomeX.Name = "NewSoftHomeX"
        Me.NewSoftHomeX.Size = New System.Drawing.Size(56, 27)
        Me.NewSoftHomeX.TabIndex = 16
        '
        'NewSoftHomeY
        '
        Me.NewSoftHomeY.Location = New System.Drawing.Point(58, 120)
        Me.NewSoftHomeY.Name = "NewSoftHomeY"
        Me.NewSoftHomeY.Size = New System.Drawing.Size(56, 27)
        Me.NewSoftHomeY.TabIndex = 15
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(32, 120)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(15, 23)
        Me.Label19.TabIndex = 14
        Me.Label19.Text = "Y"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(32, 168)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(15, 23)
        Me.Label14.TabIndex = 12
        Me.Label14.Text = "Z"
        '
        'Btn_SoftHomeTrack
        '
        Me.Btn_SoftHomeTrack.Location = New System.Drawing.Point(136, 112)
        Me.Btn_SoftHomeTrack.Name = "Btn_SoftHomeTrack"
        Me.Btn_SoftHomeTrack.Size = New System.Drawing.Size(75, 48)
        Me.Btn_SoftHomeTrack.TabIndex = 8
        Me.Btn_SoftHomeTrack.Text = "Track-Ball"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(32, 72)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(15, 23)
        Me.Label15.TabIndex = 5
        Me.Label15.Text = "X"
        '
        'CB_SoftHomeUseDefault
        '
        Me.CB_SoftHomeUseDefault.Location = New System.Drawing.Point(16, 32)
        Me.CB_SoftHomeUseDefault.Name = "CB_SoftHomeUseDefault"
        Me.CB_SoftHomeUseDefault.Size = New System.Drawing.Size(152, 24)
        Me.CB_SoftHomeUseDefault.TabIndex = 1
        Me.CB_SoftHomeUseDefault.Text = "Use Default"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.CB_DipenserLeftHead)
        Me.GroupBox7.Controls.Add(Me.CB_LeftHeadType)
        Me.GroupBox7.Controls.Add(Me.CB_DipenserRightHead)
        Me.GroupBox7.Controls.Add(Me.CB_RightHeadType)
        Me.GroupBox7.Controls.Add(Me.Label13)
        Me.GroupBox7.Controls.Add(Me.Label12)
        Me.GroupBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox7.Location = New System.Drawing.Point(296, 328)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(224, 192)
        Me.GroupBox7.TabIndex = 19
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Dispense Unit"
        '
        'CB_DipenserLeftHead
        '
        Me.CB_DipenserLeftHead.Location = New System.Drawing.Point(8, 32)
        Me.CB_DipenserLeftHead.Name = "CB_DipenserLeftHead"
        Me.CB_DipenserLeftHead.Size = New System.Drawing.Size(112, 24)
        Me.CB_DipenserLeftHead.TabIndex = 1
        Me.CB_DipenserLeftHead.Text = "Left Head"
        '
        'CB_LeftHeadType
        '
        Me.CB_LeftHeadType.Items.AddRange(New Object() {"Time Pressure", "Volumetric", "Positive Displacement", "Jetter"})
        Me.CB_LeftHeadType.Location = New System.Drawing.Point(56, 64)
        Me.CB_LeftHeadType.MaxDropDownItems = 2
        Me.CB_LeftHeadType.Name = "CB_LeftHeadType"
        Me.CB_LeftHeadType.Size = New System.Drawing.Size(160, 31)
        Me.CB_LeftHeadType.TabIndex = 20
        Me.CB_LeftHeadType.Text = "Time Pressure"
        '
        'CB_DipenserRightHead
        '
        Me.CB_DipenserRightHead.Location = New System.Drawing.Point(8, 112)
        Me.CB_DipenserRightHead.Name = "CB_DipenserRightHead"
        Me.CB_DipenserRightHead.Size = New System.Drawing.Size(152, 24)
        Me.CB_DipenserRightHead.TabIndex = 21
        Me.CB_DipenserRightHead.Text = "Right Head"
        '
        'CB_RightHeadType
        '
        Me.CB_RightHeadType.Items.AddRange(New Object() {"Time Pressure", "Volumetric", "Positive Displacement", "Jetter"})
        Me.CB_RightHeadType.Location = New System.Drawing.Point(56, 144)
        Me.CB_RightHeadType.MaxDropDownItems = 2
        Me.CB_RightHeadType.Name = "CB_RightHeadType"
        Me.CB_RightHeadType.Size = New System.Drawing.Size(160, 31)
        Me.CB_RightHeadType.TabIndex = 22
        Me.CB_RightHeadType.Text = "Time Pressure"
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label13.Location = New System.Drawing.Point(8, 144)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(48, 23)
        Me.Label13.TabIndex = 23
        Me.Label13.Text = "Type"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label12.Location = New System.Drawing.Point(8, 64)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(48, 23)
        Me.Label12.TabIndex = 20
        Me.Label12.Text = "Type"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.NewElementTravelSpeed)
        Me.GroupBox8.Controls.Add(Me.ElementTravelSpeedBar)
        Me.GroupBox8.Controls.Add(Me.Label18)
        Me.GroupBox8.Controls.Add(Me.Label16)
        Me.GroupBox8.Controls.Add(Me.Lbl_DefaultElementTravelSpeed)
        Me.GroupBox8.Controls.Add(Me.Btn_ApplyElementTravelSpeed)
        Me.GroupBox8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox8.Location = New System.Drawing.Point(296, 536)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(224, 208)
        Me.GroupBox8.TabIndex = 24
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "Element Travel Speed"
        '
        'NewElementTravelSpeed
        '
        Me.NewElementTravelSpeed.Location = New System.Drawing.Point(84, 64)
        Me.NewElementTravelSpeed.Name = "NewElementTravelSpeed"
        Me.NewElementTravelSpeed.Size = New System.Drawing.Size(56, 27)
        Me.NewElementTravelSpeed.TabIndex = 29
        Me.NewElementTravelSpeed.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'ElementTravelSpeedBar
        '
        Me.ElementTravelSpeedBar.Location = New System.Drawing.Point(16, 104)
        Me.ElementTravelSpeedBar.Name = "ElementTravelSpeedBar"
        Me.ElementTravelSpeedBar.Size = New System.Drawing.Size(192, 42)
        Me.ElementTravelSpeedBar.TabIndex = 28
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(136, 32)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(72, 23)
        Me.Label18.TabIndex = 26
        Me.Label18.Text = "mm/sec"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(16, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 23)
        Me.Label16.TabIndex = 25
        Me.Label16.Text = "Default"
        '
        'Lbl_DefaultElementTravelSpeed
        '
        Me.Lbl_DefaultElementTravelSpeed.Location = New System.Drawing.Point(88, 32)
        Me.Lbl_DefaultElementTravelSpeed.Name = "Lbl_DefaultElementTravelSpeed"
        Me.Lbl_DefaultElementTravelSpeed.Size = New System.Drawing.Size(40, 23)
        Me.Lbl_DefaultElementTravelSpeed.TabIndex = 12
        Me.Lbl_DefaultElementTravelSpeed.Text = "200"
        '
        'Btn_ApplyElementTravelSpeed
        '
        Me.Btn_ApplyElementTravelSpeed.Location = New System.Drawing.Point(76, 152)
        Me.Btn_ApplyElementTravelSpeed.Name = "Btn_ApplyElementTravelSpeed"
        Me.Btn_ApplyElementTravelSpeed.Size = New System.Drawing.Size(75, 40)
        Me.Btn_ApplyElementTravelSpeed.TabIndex = 12
        Me.Btn_ApplyElementTravelSpeed.Text = "Apply"
        '
        'Btn_CancelBasicSetup
        '
        Me.Btn_CancelBasicSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Btn_CancelBasicSetup.Location = New System.Drawing.Point(424, 840)
        Me.Btn_CancelBasicSetup.Name = "Btn_CancelBasicSetup"
        Me.Btn_CancelBasicSetup.Size = New System.Drawing.Size(75, 48)
        Me.Btn_CancelBasicSetup.TabIndex = 30
        Me.Btn_CancelBasicSetup.Text = "Cancel"
        '
        'Btn_ApplyBasicSetup
        '
        Me.Btn_ApplyBasicSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Btn_ApplyBasicSetup.Location = New System.Drawing.Point(320, 840)
        Me.Btn_ApplyBasicSetup.Name = "Btn_ApplyBasicSetup"
        Me.Btn_ApplyBasicSetup.Size = New System.Drawing.Size(75, 48)
        Me.Btn_ApplyBasicSetup.TabIndex = 31
        Me.Btn_ApplyBasicSetup.Text = "Apply"
        '
        'PanelRight
        '
        Me.PanelRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelRight.Controls.Add(Me.Label17)
        Me.PanelRight.Controls.Add(Me.GroupBox3)
        Me.PanelRight.Controls.Add(Me.GroupBox7)
        Me.PanelRight.Controls.Add(Me.Btn_CancelBasicSetup)
        Me.PanelRight.Controls.Add(Me.Btn_ApplyBasicSetup)
        Me.PanelRight.Controls.Add(Me.GroupBox2)
        Me.PanelRight.Controls.Add(Me.GroupBox5)
        Me.PanelRight.Controls.Add(Me.GroupBox6)
        Me.PanelRight.Controls.Add(Me.GroupBox8)
        Me.PanelRight.Controls.Add(Me.GroupBox4)
        Me.PanelRight.Controls.Add(Me.GroupBox1)
        Me.PanelRight.Controls.Add(Me.Label1)
        Me.PanelRight.Controls.Add(Me.Cb_SelectMeasurementUnit)
        Me.PanelRight.Location = New System.Drawing.Point(744, 24)
        Me.PanelRight.Name = "PanelRight"
        Me.PanelRight.Size = New System.Drawing.Size(528, 911)
        Me.PanelRight.TabIndex = 32
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(136, 32)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Basic Setup"
        '
        'PanelLeft
        '
        Me.PanelLeft.BackColor = System.Drawing.Color.LightSteelBlue
        Me.PanelLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelLeft.Controls.Add(Me.CheckBox1)
        Me.PanelLeft.Controls.Add(Me.GroupBox9)
        Me.PanelLeft.Controls.Add(Me.RadioButton8)
        Me.PanelLeft.Controls.Add(Me.RadioButton9)
        Me.PanelLeft.Controls.Add(Me.RadioButton1)
        Me.PanelLeft.Controls.Add(Me.RadioButton6)
        Me.PanelLeft.Controls.Add(Me.RadioButton5)
        Me.PanelLeft.Controls.Add(Me.RadioButton7)
        Me.PanelLeft.Controls.Add(Me.RadioButton4)
        Me.PanelLeft.Controls.Add(Me.RadioButton2)
        Me.PanelLeft.Controls.Add(Me.RadioButton3)
        Me.PanelLeft.Controls.Add(Me.RichTextBox1)
        Me.PanelLeft.Location = New System.Drawing.Point(0, 0)
        Me.PanelLeft.Name = "PanelLeft"
        Me.PanelLeft.Size = New System.Drawing.Size(752, 305)
        Me.PanelLeft.TabIndex = 33
        '
        'CheckBox1
        '
        Me.CheckBox1.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBox1.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.CheckBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(40, 248)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(304, 40)
        Me.CheckBox1.TabIndex = 21
        Me.CheckBox1.Text = "Go To Entire Process Setup"
        Me.CheckBox1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.DomainUpDown3)
        Me.GroupBox9.Controls.Add(Me.LabelStep)
        Me.GroupBox9.Controls.Add(Me.Button3)
        Me.GroupBox9.Controls.Add(Me.Button2)
        Me.GroupBox9.Controls.Add(Me.Button7)
        Me.GroupBox9.Controls.Add(Me.Button6)
        Me.GroupBox9.Controls.Add(Me.Button5)
        Me.GroupBox9.Controls.Add(Me.Button4)
        Me.GroupBox9.Controls.Add(Me.Label3)
        Me.GroupBox9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox9.Location = New System.Drawing.Point(392, 16)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(320, 272)
        Me.GroupBox9.TabIndex = 20
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Stepper"
        '
        'DomainUpDown3
        '
        Me.DomainUpDown3.Location = New System.Drawing.Point(152, 240)
        Me.DomainUpDown3.Name = "DomainUpDown3"
        Me.DomainUpDown3.Size = New System.Drawing.Size(72, 27)
        Me.DomainUpDown3.TabIndex = 1
        Me.DomainUpDown3.Text = "0.002"
        '
        'LabelStep
        '
        Me.LabelStep.Location = New System.Drawing.Point(48, 240)
        Me.LabelStep.Name = "LabelStep"
        Me.LabelStep.Size = New System.Drawing.Size(88, 24)
        Me.LabelStep.TabIndex = 0
        Me.LabelStep.Text = "Fine Step:"
        Me.LabelStep.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button3.Location = New System.Drawing.Point(144, 96)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(80, 48)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "X+"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button2.Location = New System.Drawing.Point(96, 144)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(48, 80)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Y-"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button7.Image = CType(resources.GetObject("Button7.Image"), System.Drawing.Image)
        Me.Button7.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button7.Location = New System.Drawing.Point(96, 16)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(48, 80)
        Me.Button7.TabIndex = 3
        Me.Button7.Text = "Y+"
        Me.Button7.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button6
        '
        Me.Button6.BackColor = System.Drawing.Color.Linen
        Me.Button6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button6.Image = CType(resources.GetObject("Button6.Image"), System.Drawing.Image)
        Me.Button6.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Button6.Location = New System.Drawing.Point(256, 128)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(48, 72)
        Me.Button6.TabIndex = 8
        Me.Button6.Text = "Dn"
        Me.Button6.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.Color.Linen
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button5.Image = CType(resources.GetObject("Button5.Image"), System.Drawing.Image)
        Me.Button5.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button5.Location = New System.Drawing.Point(256, 40)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(48, 72)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Up"
        Me.Button5.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button4.Location = New System.Drawing.Point(16, 96)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(80, 48)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "X-"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(232, 240)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 24)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "mm"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton8
        '
        Me.RadioButton8.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton8.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton8.Location = New System.Drawing.Point(144, 168)
        Me.RadioButton8.Name = "RadioButton8"
        Me.RadioButton8.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton8.TabIndex = 41
        Me.RadioButton8.Text = "Station Settings"
        Me.RadioButton8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton9
        '
        Me.RadioButton9.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton9.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton9.Location = New System.Drawing.Point(248, 168)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton9.TabIndex = 42
        Me.RadioButton9.Text = "Event Failed Actions"
        Me.RadioButton9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton1
        '
        Me.RadioButton1.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton1.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton1.Location = New System.Drawing.Point(40, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton1.TabIndex = 34
        Me.RadioButton1.Text = "Conveyor"
        Me.RadioButton1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton6
        '
        Me.RadioButton6.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton6.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton6.Location = New System.Drawing.Point(248, 96)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton6.TabIndex = 39
        Me.RadioButton6.Text = "Dispenser Units"
        Me.RadioButton6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton5
        '
        Me.RadioButton5.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton5.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton5.Location = New System.Drawing.Point(144, 96)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton5.TabIndex = 38
        Me.RadioButton5.Text = "Thermal Controller"
        Me.RadioButton5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton7
        '
        Me.RadioButton7.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton7.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton7.Location = New System.Drawing.Point(40, 168)
        Me.RadioButton7.Name = "RadioButton7"
        Me.RadioButton7.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton7.TabIndex = 40
        Me.RadioButton7.Text = "SPC Logging"
        Me.RadioButton7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton4
        '
        Me.RadioButton4.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton4.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton4.Location = New System.Drawing.Point(40, 96)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton4.TabIndex = 37
        Me.RadioButton4.Text = "Volume Cal."
        Me.RadioButton4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton2
        '
        Me.RadioButton2.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton2.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton2.Location = New System.Drawing.Point(144, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton2.TabIndex = 35
        Me.RadioButton2.Text = "Lifter and Vacuum"
        Me.RadioButton2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton3
        '
        Me.RadioButton3.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton3.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.RadioButton3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RadioButton3.Location = New System.Drawing.Point(248, 24)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(96, 64)
        Me.RadioButton3.TabIndex = 36
        Me.RadioButton3.Text = "Needle Cal."
        Me.RadioButton3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(40, 24)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(304, 208)
        Me.RichTextBox1.TabIndex = 19
        Me.RichTextBox1.Text = "It is highly recommended to go through the entire process setup for a new pattern" & _
        "." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "For an exsiting pattern, the process setup can also be modified."
        '
        'PSBasicSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1156, 849)
        Me.Controls.Add(Me.PanelLeft)
        Me.Controls.Add(Me.PanelRight)
        Me.Name = "PSBasicSetup"
        Me.Text = "Basic Setup"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ConveyorNewWidth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ConveyorFineAdjust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.GraphicBotRightX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GraphicBotRightY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GraphicTopLeftX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GraphicTopLeftY, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        CType(Me.NewSoftHomeZ, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NewSoftHomeX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NewSoftHomeY, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        CType(Me.NewElementTravelSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ElementTravelSpeedBar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelRight.ResumeLayout(False)
        Me.PanelLeft.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Cb_SelectMeasurementUnit_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cb_SelectMeasurementUnit.SelectedIndexChanged
        If Cb_SelectMeasurementUnit.SelectedItem = "mm" Then
            'Save in pattern file = mm
        Else
            'Save in pattern file = inch
        End If
    End Sub

    Private Sub Cb_SelectMeasurementUnit_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cb_SelectMeasurementUnit.TextChanged
        If Cb_SelectMeasurementUnit.Items.Contains(Cb_SelectMeasurementUnit.Text) Then

        Else
            Cb_SelectMeasurementUnit.SelectedItem = "mm"
        End If
    End Sub

    Private Sub Btn_ConveyorStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_ConveyorStart.Click
        'call conveyor start through RS232
        'ConveyorNewWidth.Value


    End Sub

    Private Sub Btn_ConveyorHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_ConveyorHome.Click
        'call conveyor home 

    End Sub

    Private Sub Btn_WidthReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_WidthReverse.Click
        'Call conveyor width adjust
        'ConveyorFineAdjust.Value
    End Sub

    Private Sub Btn_WidthForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_WidthForward.Click
        'Call conveyor width adjust
        'ConveyorFineAdjust.Value
    End Sub

    Private Sub Btn_GraphicsTopLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_GraphicsTopLeft.Click
        'Call vision and read in value
        'Update GraphicTopLeftX.Value() & GraphicTopLeftY.Value()

    End Sub

    Private Sub Btn_GraphicsBotRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_GraphicsBotRight.Click
        'Call vision and read in value
        'Update GraphicTopLeftX.Value() & GraphicTopLeftY.Value()

    End Sub

    Private Sub CB_LifterUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_LifterUp.CheckedChanged
        'Send command to conveyor to lift up lifter
    End Sub

    Private Sub CB_ThermalRequired_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_ThermalRequired.CheckedChanged
        If CB_ThermalRequired.Checked Then
            'the process required thermal
        End If
    End Sub

    Private Sub CB_ReleaseBoard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_ReleaseBoard.CheckedChanged
        'Release board status - for production
    End Sub

    Private Sub CB_SoftHomeUseDefault_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_SoftHomeUseDefault.CheckedChanged
        If CB_SoftHomeUseDefault.Checked Then
            'retrieve default value from database
            'NewSoftHomeX.Value()
            'NewSoftHomeY.Value()
            'NewSoftHomeZ.Value()
        End If
    End Sub

    Private Sub Btn_SoftHomeTrack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_SoftHomeTrack.Click
        'call vision 
    End Sub

    Private Sub CB_DipenserLeftHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_DipenserLeftHead.CheckedChanged
        If CB_DipenserLeftHead.Checked Then
            CB_LeftHeadType.Enabled = True
        Else
            CB_LeftHeadType.Enabled = False
        End If
    End Sub

    Private Sub CB_LeftHeadType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_LeftHeadType.SelectedIndexChanged
        'Selected Left Dispenser Type = CB_LeftHeadType.SelectedItem
    End Sub

    Private Sub CB_DipenserRightHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_DipenserRightHead.CheckedChanged
        If CB_DipenserRightHead.Checked Then
            CB_RightHeadType.Enabled = True
        Else
            CB_RightHeadType.Enabled = False
        End If
    End Sub

    Private Sub CB_RightHeadType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_RightHeadType.SelectedIndexChanged
        'Selected Right Dispenser Type = CB_LeftHeadType.SelectedItem
    End Sub

    Private Sub ElementTravelSpeedBar_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ElementTravelSpeedBar.Scroll
        NewElementTravelSpeed.Value = ElementTravelSpeedBar.Value
    End Sub

    Private Sub NewElementTravelSpeed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewElementTravelSpeed.ValueChanged
        If NewElementTravelSpeed.Value > ElementTravelSpeedBar.Maximum Then
            NewElementTravelSpeed.Value = ElementTravelSpeedBar.Maximum
        ElseIf NewElementTravelSpeed.Value < ElementTravelSpeedBar.Maximum Then
            NewElementTravelSpeed.Value = ElementTravelSpeedBar.Minimum
        Else
            NewElementTravelSpeed.Value = ElementTravelSpeedBar.Value
        End If
    End Sub

    Private Sub Btn_ApplyBasicSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_ApplyBasicSetup.Click
        'Save all parameters with pattern file
    End Sub

    Private Sub Btn_CancelBasicSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_CancelBasicSetup.Click
        'do not save changes- retrieve previous saved parameters from pattern file
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            RichTextBox1.Hide()
            'CheckBox1.Enabled = False
            CheckBox1.Text = "Back to Basic Setup"
            If RadioButton1.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSConveyor.PanelToBeAdded)
                MyPSConveyor.PanelToBeAdded.BringToFront()
                MyPSConveyor.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton2.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSLifterAndVacuum.PanelToBeAdded)
                MyPSLifterAndVacuum.PanelToBeAdded.BringToFront()
                MyPSLifterAndVacuum.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton3.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSNeedleCal.PanelToBeAdded)
                MyPSNeedleCal.PanelToBeAdded.BringToFront()
                MyPSNeedleCal.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton4.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSVolumeCal.PanelToBeAdded)
                MyPSVolumeCal.PanelToBeAdded.BringToFront()
                MyPSVolumeCal.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton5.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSThermalController.PanelToBeAdded)
                MyPSThermalController.PanelToBeAdded.BringToFront()
                MyPSThermalController.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton6.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSDispenserUnit.PanelToBeAdded)
                MyPSDispenserUnit.PanelToBeAdded.BringToFront()
                MyPSDispenserUnit.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton7.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSSPCLogging.PanelToBeAdded)
                MyPSSPCLogging.PanelToBeAdded.BringToFront()
                MyPSSPCLogging.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton8.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSStationSettings.PanelToBeAdded)
                MyPSStationSettings.PanelToBeAdded.BringToFront()
                MyPSStationSettings.PanelToBeAdded.Location = New Point(0, 0)
            ElseIf RadioButton9.Checked Then
                MyBasicSetup.PanelRight.Controls.Add(MyPSEventFailedActions.PanelToBeAdded)
                MyPSEventFailedActions.PanelToBeAdded.BringToFront()
                MyPSEventFailedActions.PanelToBeAdded.Location = New Point(0, 0)
            End If
        Else
            RichTextBox1.Show()
            CheckBox1.Text = "Go To Entire Process Setup"
            MyBasicSetup.PanelRight.Controls.Remove(MyPSConveyor.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSDispenserUnit.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSEventFailedActions.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSLifterAndVacuum.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSSPCLogging.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSStationSettings.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSThermalController.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSDispenserUnit.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSNeedleCal.PanelToBeAdded)
            MyBasicSetup.PanelRight.Controls.Remove(MyPSVolumeCal.PanelToBeAdded)
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSConveyor.PanelToBeAdded)
            MyPSConveyor.PanelToBeAdded.BringToFront()
            MyPSConveyor.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSConveyor.PanelToBeAdded)
        End If

    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSDispenserUnit.PanelToBeAdded)
            MyPSDispenserUnit.PanelToBeAdded.BringToFront()
            MyPSDispenserUnit.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSDispenserUnit.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSLifterAndVacuum.PanelToBeAdded)
            MyPSLifterAndVacuum.PanelToBeAdded.BringToFront()
            MyPSLifterAndVacuum.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSLifterAndVacuum.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSThermalController.PanelToBeAdded)
            MyPSThermalController.PanelToBeAdded.BringToFront()
            MyPSThermalController.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSThermalController.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSSPCLogging.PanelToBeAdded)
            MyPSSPCLogging.PanelToBeAdded.BringToFront()
            MyPSSPCLogging.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSSPCLogging.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton8.CheckedChanged
        If RadioButton8.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSStationSettings.PanelToBeAdded)
            MyPSStationSettings.PanelToBeAdded.BringToFront()
            MyPSStationSettings.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSStationSettings.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton9.CheckedChanged
        If RadioButton9.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSEventFailedActions.PanelToBeAdded)
            MyPSEventFailedActions.PanelToBeAdded.BringToFront()
            MyPSEventFailedActions.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSEventFailedActions.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSNeedleCal.PanelToBeAdded)
            MyPSNeedleCal.PanelToBeAdded.BringToFront()
            MyPSNeedleCal.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSNeedleCal.PanelToBeAdded)
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            MyBasicSetup.PanelRight.Controls.Add(MyPSVolumeCal.PanelToBeAdded)
            MyPSVolumeCal.PanelToBeAdded.BringToFront()
            MyPSVolumeCal.PanelToBeAdded.Location = New Point(0, 0)
        Else
            MyBasicSetup.PanelRight.Controls.Remove(MyPSVolumeCal.PanelToBeAdded)
        End If
    End Sub
End Class
