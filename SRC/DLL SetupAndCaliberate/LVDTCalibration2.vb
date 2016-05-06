Imports DLL_Export_Service

Public Class LVDTCalibration2
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonBack As System.Windows.Forms.Button
    Friend WithEvents ButtonProbeCalStart As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents NewLVDTZReferencePos As System.Windows.Forms.Label
    Friend WithEvents NewLVDTYOffsetPos As System.Windows.Forms.Label
    Friend WithEvents NewLVDTXOffsetPos As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ButtonLTest As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents LBL_LTestScore As System.Windows.Forms.Label
    Friend WithEvents NB_LThresholdValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_LBrightnessValue As System.Windows.Forms.NumericUpDown
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents RD_LBoW As System.Windows.Forms.RadioButton
    Friend WithEvents RD_LWoB As System.Windows.Forms.RadioButton
    Friend WithEvents NB_Lcompactness As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_LRoughness As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_LminRadius As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_maxRadius As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_Lclose As System.Windows.Forms.NumericUpDown
    Friend WithEvents NB_Lopen As System.Windows.Forms.NumericUpDown
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.NB_Lcompactness = New System.Windows.Forms.NumericUpDown
        Me.NB_LRoughness = New System.Windows.Forms.NumericUpDown
        Me.NB_LminRadius = New System.Windows.Forms.NumericUpDown
        Me.NB_maxRadius = New System.Windows.Forms.NumericUpDown
        Me.NB_Lclose = New System.Windows.Forms.NumericUpDown
        Me.NB_Lopen = New System.Windows.Forms.NumericUpDown
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.ButtonLTest = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.LBL_LTestScore = New System.Windows.Forms.Label
        Me.NB_LThresholdValue = New System.Windows.Forms.NumericUpDown
        Me.NB_LBrightnessValue = New System.Windows.Forms.NumericUpDown
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.RD_LBoW = New System.Windows.Forms.RadioButton
        Me.RD_LWoB = New System.Windows.Forms.RadioButton
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.NewLVDTZReferencePos = New System.Windows.Forms.Label
        Me.NewLVDTYOffsetPos = New System.Windows.Forms.Label
        Me.NewLVDTXOffsetPos = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ButtonProbeCalStart = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonBack = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        CType(Me.NB_Lcompactness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_LRoughness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_LminRadius, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_maxRadius, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_Lclose, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_Lopen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_LThresholdValue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NB_LBrightnessValue, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 35
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox6)
        Me.GroupBox1.Controls.Add(Me.GroupBox4)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ButtonSave)
        Me.GroupBox1.Controls.Add(Me.ButtonRevert)
        Me.GroupBox1.Controls.Add(Me.ButtonBack)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(16, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 760)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Position Offset of LVDT"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.NB_Lcompactness)
        Me.GroupBox6.Controls.Add(Me.NB_LRoughness)
        Me.GroupBox6.Controls.Add(Me.NB_LminRadius)
        Me.GroupBox6.Controls.Add(Me.NB_maxRadius)
        Me.GroupBox6.Controls.Add(Me.NB_Lclose)
        Me.GroupBox6.Controls.Add(Me.NB_Lopen)
        Me.GroupBox6.Controls.Add(Me.Label18)
        Me.GroupBox6.Controls.Add(Me.Label17)
        Me.GroupBox6.Controls.Add(Me.Label16)
        Me.GroupBox6.Controls.Add(Me.Label14)
        Me.GroupBox6.Controls.Add(Me.Label8)
        Me.GroupBox6.Controls.Add(Me.Label2)
        Me.GroupBox6.Controls.Add(Me.Label13)
        Me.GroupBox6.Controls.Add(Me.Label9)
        Me.GroupBox6.Controls.Add(Me.ButtonLTest)
        Me.GroupBox6.Controls.Add(Me.Label3)
        Me.GroupBox6.Controls.Add(Me.Label10)
        Me.GroupBox6.Controls.Add(Me.LBL_LTestScore)
        Me.GroupBox6.Controls.Add(Me.NB_LThresholdValue)
        Me.GroupBox6.Controls.Add(Me.NB_LBrightnessValue)
        Me.GroupBox6.Location = New System.Drawing.Point(8, 192)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(456, 264)
        Me.GroupBox6.TabIndex = 16
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Vision Threshold Setting"
        '
        'NB_Lcompactness
        '
        Me.NB_Lcompactness.DecimalPlaces = 2
        Me.NB_Lcompactness.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_Lcompactness.Location = New System.Drawing.Point(352, 160)
        Me.NB_Lcompactness.Name = "NB_Lcompactness"
        Me.NB_Lcompactness.Size = New System.Drawing.Size(80, 27)
        Me.NB_Lcompactness.TabIndex = 63
        Me.NB_Lcompactness.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_LRoughness
        '
        Me.NB_LRoughness.DecimalPlaces = 2
        Me.NB_LRoughness.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_LRoughness.Location = New System.Drawing.Point(352, 120)
        Me.NB_LRoughness.Name = "NB_LRoughness"
        Me.NB_LRoughness.Size = New System.Drawing.Size(80, 27)
        Me.NB_LRoughness.TabIndex = 64
        Me.NB_LRoughness.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_LminRadius
        '
        Me.NB_LminRadius.DecimalPlaces = 2
        Me.NB_LminRadius.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_LminRadius.Location = New System.Drawing.Point(352, 80)
        Me.NB_LminRadius.Name = "NB_LminRadius"
        Me.NB_LminRadius.Size = New System.Drawing.Size(80, 27)
        Me.NB_LminRadius.TabIndex = 61
        Me.NB_LminRadius.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_maxRadius
        '
        Me.NB_maxRadius.DecimalPlaces = 2
        Me.NB_maxRadius.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_maxRadius.Location = New System.Drawing.Point(352, 40)
        Me.NB_maxRadius.Name = "NB_maxRadius"
        Me.NB_maxRadius.Size = New System.Drawing.Size(80, 27)
        Me.NB_maxRadius.TabIndex = 62
        Me.NB_maxRadius.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_Lclose
        '
        Me.NB_Lclose.DecimalPlaces = 2
        Me.NB_Lclose.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_Lclose.Location = New System.Drawing.Point(120, 160)
        Me.NB_Lclose.Name = "NB_Lclose"
        Me.NB_Lclose.Size = New System.Drawing.Size(80, 27)
        Me.NB_Lclose.TabIndex = 59
        Me.NB_Lclose.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_Lopen
        '
        Me.NB_Lopen.DecimalPlaces = 2
        Me.NB_Lopen.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_Lopen.Location = New System.Drawing.Point(120, 120)
        Me.NB_Lopen.Name = "NB_Lopen"
        Me.NB_Lopen.Size = New System.Drawing.Size(80, 27)
        Me.NB_Lopen.TabIndex = 60
        Me.NB_Lopen.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(240, 160)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(128, 23)
        Me.Label18.TabIndex = 58
        Me.Label18.Text = "Compactness"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(240, 120)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(104, 23)
        Me.Label17.TabIndex = 57
        Me.Label17.Text = "Roughness"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(240, 80)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(104, 23)
        Me.Label16.TabIndex = 56
        Me.Label16.Text = "Min Radius"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(240, 40)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(104, 23)
        Me.Label14.TabIndex = 55
        Me.Label14.Text = "Max Radius"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 23)
        Me.Label8.TabIndex = 54
        Me.Label8.Text = "Open"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 23)
        Me.Label2.TabIndex = 53
        Me.Label2.Text = "Close"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(336, 216)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(24, 23)
        Me.Label13.TabIndex = 52
        Me.Label13.Text = "%"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(224, 216)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 23)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Score"
        '
        'ButtonLTest
        '
        Me.ButtonLTest.Location = New System.Drawing.Point(112, 208)
        Me.ButtonLTest.Name = "ButtonLTest"
        Me.ButtonLTest.Size = New System.Drawing.Size(100, 40)
        Me.ButtonLTest.TabIndex = 46
        Me.ButtonLTest.Text = "Test"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 23)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Brightness"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(24, 80)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 23)
        Me.Label10.TabIndex = 1
        Me.Label10.Text = "Threshold"
        '
        'LBL_LTestScore
        '
        Me.LBL_LTestScore.Location = New System.Drawing.Point(288, 216)
        Me.LBL_LTestScore.Name = "LBL_LTestScore"
        Me.LBL_LTestScore.Size = New System.Drawing.Size(64, 24)
        Me.LBL_LTestScore.TabIndex = 5
        Me.LBL_LTestScore.Text = "68.50"
        '
        'NB_LThresholdValue
        '
        Me.NB_LThresholdValue.DecimalPlaces = 2
        Me.NB_LThresholdValue.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_LThresholdValue.Location = New System.Drawing.Point(120, 80)
        Me.NB_LThresholdValue.Name = "NB_LThresholdValue"
        Me.NB_LThresholdValue.Size = New System.Drawing.Size(80, 27)
        Me.NB_LThresholdValue.TabIndex = 5
        Me.NB_LThresholdValue.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'NB_LBrightnessValue
        '
        Me.NB_LBrightnessValue.DecimalPlaces = 2
        Me.NB_LBrightnessValue.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.NB_LBrightnessValue.Location = New System.Drawing.Point(120, 40)
        Me.NB_LBrightnessValue.Name = "NB_LBrightnessValue"
        Me.NB_LBrightnessValue.Size = New System.Drawing.Size(80, 27)
        Me.NB_LBrightnessValue.TabIndex = 6
        Me.NB_LBrightnessValue.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.RD_LBoW)
        Me.GroupBox4.Controls.Add(Me.RD_LWoB)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 112)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(456, 64)
        Me.GroupBox4.TabIndex = 15
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Vision Background Setup"
        '
        'RD_LBoW
        '
        Me.RD_LBoW.Location = New System.Drawing.Point(264, 32)
        Me.RD_LBoW.Name = "RD_LBoW"
        Me.RD_LBoW.Size = New System.Drawing.Size(168, 24)
        Me.RD_LBoW.TabIndex = 1
        Me.RD_LBoW.Text = "Black on White"
        '
        'RD_LWoB
        '
        Me.RD_LWoB.Checked = True
        Me.RD_LWoB.Location = New System.Drawing.Point(48, 32)
        Me.RD_LWoB.Name = "RD_LWoB"
        Me.RD_LWoB.Size = New System.Drawing.Size(176, 24)
        Me.RD_LWoB.TabIndex = 0
        Me.RD_LWoB.TabStop = True
        Me.RD_LWoB.Text = "White on Black"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.NewLVDTZReferencePos)
        Me.GroupBox2.Controls.Add(Me.NewLVDTYOffsetPos)
        Me.GroupBox2.Controls.Add(Me.NewLVDTXOffsetPos)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.ButtonProbeCalStart)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 472)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(456, 184)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Locate Offset"
        '
        'NewLVDTZReferencePos
        '
        Me.NewLVDTZReferencePos.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.NewLVDTZReferencePos.Location = New System.Drawing.Point(240, 152)
        Me.NewLVDTZReferencePos.Name = "NewLVDTZReferencePos"
        Me.NewLVDTZReferencePos.Size = New System.Drawing.Size(72, 16)
        Me.NewLVDTZReferencePos.TabIndex = 34
        Me.NewLVDTZReferencePos.Text = "0.000"
        '
        'NewLVDTYOffsetPos
        '
        Me.NewLVDTYOffsetPos.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.NewLVDTYOffsetPos.Location = New System.Drawing.Point(240, 120)
        Me.NewLVDTYOffsetPos.Name = "NewLVDTYOffsetPos"
        Me.NewLVDTYOffsetPos.Size = New System.Drawing.Size(72, 16)
        Me.NewLVDTYOffsetPos.TabIndex = 33
        Me.NewLVDTYOffsetPos.Text = "0.000"
        '
        'NewLVDTXOffsetPos
        '
        Me.NewLVDTXOffsetPos.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.NewLVDTXOffsetPos.Location = New System.Drawing.Point(240, 88)
        Me.NewLVDTXOffsetPos.Name = "NewLVDTXOffsetPos"
        Me.NewLVDTXOffsetPos.Size = New System.Drawing.Size(72, 16)
        Me.NewLVDTXOffsetPos.TabIndex = 32
        Me.NewLVDTXOffsetPos.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(64, 152)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(160, 24)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "Current Z Reference"
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(64, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(144, 24)
        Me.Label6.TabIndex = 30
        Me.Label6.Text = "Current Y Offset"
        '
        'Label5
        '
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label5.Location = New System.Drawing.Point(32, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(176, 24)
        Me.Label5.TabIndex = 29
        Me.Label5.Text = "Offset Result:"
        '
        'Label4
        '
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(64, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 24)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "Current X Offset"
        '
        'ButtonProbeCalStart
        '
        Me.ButtonProbeCalStart.Location = New System.Drawing.Point(328, 104)
        Me.ButtonProbeCalStart.Name = "ButtonProbeCalStart"
        Me.ButtonProbeCalStart.Size = New System.Drawing.Size(100, 50)
        Me.ButtonProbeCalStart.TabIndex = 9
        Me.ButtonProbeCalStart.Text = "Start"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label1.Location = New System.Drawing.Point(24, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(400, 48)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "4. Move Camera to locate the Offset Calibrator . Click Start Locate Offset Button" & _
        ""
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(368, 680)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(100, 50)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(200, 680)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(100, 50)
        Me.ButtonRevert.TabIndex = 7
        Me.ButtonRevert.Text = "Cancel"
        '
        'ButtonBack
        '
        Me.ButtonBack.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonBack.Location = New System.Drawing.Point(32, 680)
        Me.ButtonBack.Name = "ButtonBack"
        Me.ButtonBack.Size = New System.Drawing.Size(100, 50)
        Me.ButtonBack.TabIndex = 9
        Me.ButtonBack.Text = "Back"
        '
        'Timer1
        '
        '
        'LVDTCalibration2
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 878)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "LVDTCalibration2"
        Me.Text = "LVDT1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        CType(Me.NB_Lcompactness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_LRoughness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_LminRadius, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_maxRadius, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_Lclose, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_Lopen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_LThresholdValue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NB_LBrightnessValue, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

     

    Dim timer As Boolean

    Private Sub ButtonBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBack.Click
        'GroupBox1.Hide()
        MyLVDTSetup.GroupBox1.Controls.Remove(MyLVDTSetup1.GroupBox1)
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        ''save user inputs to IDS.Data.Hardware
        ''IDSData.HardWare.HeightSensor.TP.OffsetPos.X = NewLVDTXOffsetPos.Text
        ''IDSData.HardWare.HeightSensor.TP.OffestPos.Y = NewLVDTYOffsetPos.Text
        'IDSData.HardWare.HeightSensor.TP.HeightReference = NewLVDTZReferencePos.Text

        'MyLVDTSetup.LVDTOffsetXPos.Text = NewLVDTXOffsetPos.Text
        'MyLVDTSetup.LVDTOffsetYPos.Text = NewLVDTYOffsetPos.Text
        'MyLVDTSetup.LVDTReferenceZPos.Text = NewLVDTZReferencePos.Text
        'ButtonProbeCalStart.Enabled = False

        'Para.RecordID = "FactoryDefault"
        'Para.GroupID = IDS.Data.Admin.User.Group.ID
        'IDS.Data.OpenData()

        IDS.Data.Hardware.HeightSensor.LVDTCalBackGround = RD_LWoB.Checked  'true = W on B, False = B on W
        IDS.Data.Hardware.HeightSensor.LVDTCalBrightness = NB_LBrightnessValue.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalClose = NB_Lclose.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalOpen = NB_Lopen.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalThreshold = NB_LThresholdValue.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalRoughness = NB_LBrightnessValue.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalMaxRadius = NB_maxRadius.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalMinRadius = NB_LminRadius.Value
        IDS.Data.Hardware.HeightSensor.LVDTCalCompactness = NB_Lcompactness.Value

        IDS.Data.Hardware.HeightSensor.TP.HeightReference = NewLVDTZReferencePos.Text
        IDS.Data.Hardware.HeightSensor.TP.OffsetPos.X = NewLVDTXOffsetPos.Text
        IDS.Data.Hardware.HeightSensor.TP.OffsetPos.Y = NewLVDTYOffsetPos.Text

        IDS.Data.SaveData()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        ''retrieve old data from IDS.Data.Hardware
        ''NewLVDTXOffsetPos.Text = IDSData.HardWare.HeightSensor.TP.OffsetPos.X
        ''NewLVDTYOffsetPos.Text = IDSData.HardWare.HeightSensor.TP.OffsetPos.Y
        'NewLVDTZReferencePos.Text = IDSData.HardWare.HeightSensor.TP.HeightReference
        ''Refresh offsetpos
        ''MyLVDTSetup.LVDTOffsetXPos.Text = NewLVDTXOffsetPos.Text
        ''MyLVDTSetup.LVDTOffsetYPos.Text = NewLVDTYOffsetPos.Text
        ''MyLVDTSetup.LVDTReferenceZPos.Text = NewLVDTZReferencePos.Text

        'ButtonProbeCalStart.Enabled = False
        'ButtonLVDTOK.Enabled = True
        'Para.RecordID = "FactoryDefault"
        'Para.GroupID = IDS.Data.Admin.User.Group.ID
        'IDS.Data.OpenData()

        RD_LWoB.Checked = IDS.Data.Hardware.HeightSensor.LVDTCalBackGround
        NB_LBrightnessValue.Text = IDS.Data.Hardware.HeightSensor.LVDTCalBrightness
        NB_Lclose.Text = IDS.Data.Hardware.HeightSensor.LVDTCalClose
        NB_Lopen.Text = IDS.Data.Hardware.HeightSensor.LVDTCalOpen
        NB_LThresholdValue.Text = IDS.Data.Hardware.HeightSensor.LVDTCalThreshold
        NB_LBrightnessValue.Text = IDS.Data.Hardware.HeightSensor.LVDTCalRoughness
        NB_maxRadius.Text = IDS.Data.Hardware.HeightSensor.LVDTCalMaxRadius
        NB_LminRadius.Text = IDS.Data.Hardware.HeightSensor.LVDTCalMinRadius
        NB_Lcompactness.Text = IDS.Data.Hardware.HeightSensor.LVDTCalCompactness

        NewLVDTZReferencePos.Text = IDS.Data.Hardware.HeightSensor.TP.HeightReference
        NewLVDTXOffsetPos.Text = IDS.Data.Hardware.HeightSensor.TP.OffsetPos.X
        NewLVDTYOffsetPos.Text = IDS.Data.Hardware.HeightSensor.TP.OffsetPos.Y

    End Sub




    Private Sub ButtonLVDTOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Get results from Shen Jian's module - X & Y only
        ButtonProbeCalStart.Enabled = True
        'ButtonLVDTOK.Enabled = False
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()

        'IDS.Devices.Vision.IDSV_NC(IDS.__ID, RD_RWoB.Checked, NB_RThresholdValue.Value, NB_RmaxRadius.Value, NB_RminRadius.Value, _
        'NB_Rclose.Value, NB_Ropen.Value, NB_Rroughness.Value, NB_Rcompactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.OFFSET_X, MyNeedleCalibrationSetup1.OFFSET_Y)
        IDS.Devices.Vision.IDSV_NC(RD_LWoB.Checked, NB_LThresholdValue.Value, NB_maxRadius.Value, NB_LminRadius.Value, _
NB_Lclose.Value, NB_Lopen.Value, NB_LRoughness.Value, NB_Lcompactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.OFFSET_X, MyNeedleCalibrationSetup1.OFFSET_Y)

        If timer Then
            Timer1.Start()
        End If
    End Sub


    Private Sub ButtonProbeCalStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonProbeCalStart.Click
        'MessageBox.Show("Probe Calibration in progress...")
        'camera module returns center point - X & Y Only
        'Display offset value using shen jian & camera value
        If IDS.Devices.Vision.IDSV_NC(RD_LWoB.Checked, NB_LThresholdValue.Value, NB_maxRadius.Value, NB_LminRadius.Value, _
    NB_Lclose.Value, NB_Lopen.Value, NB_LRoughness.Value, NB_Lcompactness.Value, 768 / 2, 576 / 2, 700, 550, MyNeedleCalibrationSetup1.OFFSET_X, MyNeedleCalibrationSetup1.OFFSET_Y) Then

            distance(0) = MyNeedleCalibrationSetup1.OFFSET_X  '''
            distance(1) = -MyNeedleCalibrationSetup1.OFFSET_Y   '''

            m_Tri.MoveRelative_XY(distance)

            m_Tri.GetIDSState()
            IDS.Data.Hardware.HeightSensor.TP.OffsetPos.X = m_Tri.XPosition - IDS.Data.Hardware.HeightSensor.TP.CurrentPos.X
            IDS.Data.Hardware.HeightSensor.TP.OffsetPos.Y = m_Tri.YPosition - IDS.Data.Hardware.HeightSensor.TP.CurrentPos.Y

            IDS.Data.SaveData()
            MessageBox.Show("ok.")
        Else
            MessageBox.Show("cannot find the object.")
        End If

    End Sub

    Private Sub ButtonLTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLTest.Click
        If timer Then
            Timer1.Stop()
            timer = False
            ButtonLTest.Text = "TEST"
        Else
            Timer1.Start()
            timer = True
            ButtonLTest.Text = "STOP"
        End If
    End Sub
End Class
