Imports DLL_Export_Service

Public Class GantrySetup
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ButtonSPCancel As System.Windows.Forms.Button
    Friend WithEvents ButtonSPOK As System.Windows.Forms.Button
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents SystemOriginPosZ As System.Windows.Forms.NumericUpDown
    Friend WithEvents SystemOriginPosY As System.Windows.Forms.NumericUpDown
    Friend WithEvents SystemOriginPosX As System.Windows.Forms.NumericUpDown
    Friend WithEvents WorkAreaY As System.Windows.Forms.NumericUpDown
    Friend WithEvents WorkAreaX As System.Windows.Forms.NumericUpDown
    Friend WithEvents WorkAreaZMin As System.Windows.Forms.NumericUpDown
    Friend WithEvents WorkAreaZMax As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown5 As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown6 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents HomeSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents MaxAccLimit As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents HomeSpeedTrackbar As System.Windows.Forms.TrackBar
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents MaxAccLimitTrackbar As System.Windows.Forms.TrackBar
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents MaxSpeedLimit As System.Windows.Forms.NumericUpDown
    Friend WithEvents MaxSpeedLimitTrackbar As System.Windows.Forms.TrackBar
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonTeachPos As System.Windows.Forms.Button
    Friend WithEvents ButtonGotoPositionDefined As System.Windows.Forms.Button
    Friend WithEvents ButtonGotoUserInput As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Vision As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents MoveButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents YPosition As System.Windows.Forms.Label
    Friend WithEvents XPosition As System.Windows.Forms.Label
    Friend WithEvents StationPosition As System.Windows.Forms.ComboBox
    Friend WithEvents SavePositionButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Public WithEvents Needle As System.Windows.Forms.RadioButton
    Public WithEvents RightHead As System.Windows.Forms.RadioButton
    Public WithEvents LeftHead As System.Windows.Forms.RadioButton

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(GantrySetup))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.SavePositionButton = New System.Windows.Forms.Button
        Me.YPosition = New System.Windows.Forms.Label
        Me.XPosition = New System.Windows.Forms.Label
        Me.StationPosition = New System.Windows.Forms.ComboBox
        Me.MoveButton = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.ButtonTeachPos = New System.Windows.Forms.Button
        Me.ButtonGotoUserInput = New System.Windows.Forms.Button
        Me.ButtonGotoPositionDefined = New System.Windows.Forms.Button
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.RightHead = New System.Windows.Forms.RadioButton
        Me.LeftHead = New System.Windows.Forms.RadioButton
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.MaxSpeedLimit = New System.Windows.Forms.NumericUpDown
        Me.HomeSpeed = New System.Windows.Forms.NumericUpDown
        Me.MaxAccLimit = New System.Windows.Forms.NumericUpDown
        Me.Label17 = New System.Windows.Forms.Label
        Me.MaxSpeedLimitTrackbar = New System.Windows.Forms.TrackBar
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.HomeSpeedTrackbar = New System.Windows.Forms.TrackBar
        Me.Label11 = New System.Windows.Forms.Label
        Me.MaxAccLimitTrackbar = New System.Windows.Forms.TrackBar
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.WorkAreaY = New System.Windows.Forms.NumericUpDown
        Me.WorkAreaX = New System.Windows.Forms.NumericUpDown
        Me.Label20 = New System.Windows.Forms.Label
        Me.WorkAreaZMin = New System.Windows.Forms.NumericUpDown
        Me.Label23 = New System.Windows.Forms.Label
        Me.WorkAreaZMax = New System.Windows.Forms.NumericUpDown
        Me.ButtonSPCancel = New System.Windows.Forms.Button
        Me.ButtonSPOK = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.SystemOriginPosX = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SystemOriginPosZ = New System.Windows.Forms.NumericUpDown
        Me.SystemOriginPosY = New System.Windows.Forms.NumericUpDown
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.MaxSpeedLimit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HomeSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxAccLimit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxSpeedLimitTrackbar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.HomeSpeedTrackbar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.MaxAccLimitTrackbar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.WorkAreaY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WorkAreaX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WorkAreaZMin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WorkAreaZMax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.SystemOriginPosX, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SystemOriginPosZ, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SystemOriginPosY, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox6)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox3)
        Me.PanelToBeAdded.Controls.Add(Me.Label24)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSPCancel)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSPOK)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox4)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 0
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.SavePositionButton)
        Me.GroupBox6.Controls.Add(Me.YPosition)
        Me.GroupBox6.Controls.Add(Me.XPosition)
        Me.GroupBox6.Controls.Add(Me.StationPosition)
        Me.GroupBox6.Controls.Add(Me.MoveButton)
        Me.GroupBox6.Controls.Add(Me.GroupBox7)
        Me.GroupBox6.Location = New System.Drawing.Point(8, 152)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(496, 320)
        Me.GroupBox6.TabIndex = 68
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Station Positions"
        '
        'SavePositionButton
        '
        Me.SavePositionButton.Location = New System.Drawing.Point(254, 208)
        Me.SavePositionButton.Name = "SavePositionButton"
        Me.SavePositionButton.Size = New System.Drawing.Size(168, 48)
        Me.SavePositionButton.TabIndex = 69
        Me.SavePositionButton.Text = "Save Current Position"
        '
        'YPosition
        '
        Me.YPosition.Location = New System.Drawing.Point(260, 176)
        Me.YPosition.Name = "YPosition"
        Me.YPosition.Size = New System.Drawing.Size(128, 23)
        Me.YPosition.TabIndex = 67
        Me.YPosition.Text = "0"
        Me.YPosition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'XPosition
        '
        Me.XPosition.Location = New System.Drawing.Point(108, 176)
        Me.XPosition.Name = "XPosition"
        Me.XPosition.Size = New System.Drawing.Size(128, 23)
        Me.XPosition.TabIndex = 66
        Me.XPosition.Text = "0"
        Me.XPosition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'StationPosition
        '
        Me.StationPosition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.StationPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.StationPosition.Items.AddRange(New Object() {"Park Position", "Purge Position", "Clean Position", "Change Syringe Position"})
        Me.StationPosition.Location = New System.Drawing.Point(112, 120)
        Me.StationPosition.Name = "StationPosition"
        Me.StationPosition.Size = New System.Drawing.Size(272, 28)
        Me.StationPosition.TabIndex = 64
        '
        'MoveButton
        '
        Me.MoveButton.Location = New System.Drawing.Point(74, 208)
        Me.MoveButton.Name = "MoveButton"
        Me.MoveButton.Size = New System.Drawing.Size(168, 48)
        Me.MoveButton.TabIndex = 65
        Me.MoveButton.Text = "Move to Saved Station Position"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.GroupBox5)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(184, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(128, 24)
        Me.GroupBox4.TabIndex = 61
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Mode"
        Me.GroupBox4.Visible = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.ButtonTeachPos)
        Me.GroupBox5.Controls.Add(Me.ButtonGotoUserInput)
        Me.GroupBox5.Controls.Add(Me.ButtonGotoPositionDefined)
        Me.GroupBox5.Location = New System.Drawing.Point(-16, 64)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(480, 248)
        Me.GroupBox5.TabIndex = 76
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "GroupBox5"
        '
        'ButtonTeachPos
        '
        Me.ButtonTeachPos.Location = New System.Drawing.Point(16, 160)
        Me.ButtonTeachPos.Name = "ButtonTeachPos"
        Me.ButtonTeachPos.Size = New System.Drawing.Size(80, 64)
        Me.ButtonTeachPos.TabIndex = 75
        Me.ButtonTeachPos.Text = "Save Current StationPosition"
        '
        'ButtonGotoUserInput
        '
        Me.ButtonGotoUserInput.Location = New System.Drawing.Point(112, 160)
        Me.ButtonGotoUserInput.Name = "ButtonGotoUserInput"
        Me.ButtonGotoUserInput.Size = New System.Drawing.Size(108, 64)
        Me.ButtonGotoUserInput.TabIndex = 57
        Me.ButtonGotoUserInput.Text = "Move to Saved StationPosition"
        '
        'ButtonGotoPositionDefined
        '
        Me.ButtonGotoPositionDefined.Location = New System.Drawing.Point(240, 160)
        Me.ButtonGotoPositionDefined.Name = "ButtonGotoPositionDefined"
        Me.ButtonGotoPositionDefined.Size = New System.Drawing.Size(80, 64)
        Me.ButtonGotoPositionDefined.TabIndex = 72
        Me.ButtonGotoPositionDefined.Text = "Revert"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.RightHead)
        Me.GroupBox7.Controls.Add(Me.LeftHead)
        Me.GroupBox7.Controls.Add(Me.GroupBox8)
        Me.GroupBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox7.Location = New System.Drawing.Point(112, 40)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(272, 64)
        Me.GroupBox7.TabIndex = 61
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Current Head"
        '
        'RightHead
        '
        Me.RightHead.Location = New System.Drawing.Point(112, 32)
        Me.RightHead.Name = "RightHead"
        Me.RightHead.Size = New System.Drawing.Size(72, 24)
        Me.RightHead.TabIndex = 3
        Me.RightHead.Text = "Right"
        '
        'LeftHead
        '
        Me.LeftHead.Checked = True
        Me.LeftHead.Location = New System.Drawing.Point(24, 32)
        Me.LeftHead.Name = "LeftHead"
        Me.LeftHead.Size = New System.Drawing.Size(64, 24)
        Me.LeftHead.TabIndex = 2
        Me.LeftHead.TabStop = True
        Me.LeftHead.Text = "Left"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.Button1)
        Me.GroupBox8.Controls.Add(Me.Button2)
        Me.GroupBox8.Controls.Add(Me.Button3)
        Me.GroupBox8.Location = New System.Drawing.Point(-16, 64)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(480, 248)
        Me.GroupBox8.TabIndex = 76
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "GroupBox5"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(16, 160)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 64)
        Me.Button1.TabIndex = 75
        Me.Button1.Text = "Save Current StationPosition"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(112, 160)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(108, 64)
        Me.Button2.TabIndex = 57
        Me.Button2.Text = "Move to Saved StationPosition"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(240, 160)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(80, 64)
        Me.Button3.TabIndex = 72
        Me.Button3.Text = "Revert"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.MaxSpeedLimit)
        Me.GroupBox3.Controls.Add(Me.HomeSpeed)
        Me.GroupBox3.Controls.Add(Me.MaxAccLimit)
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.MaxSpeedLimitTrackbar)
        Me.GroupBox3.Controls.Add(Me.Label18)
        Me.GroupBox3.Controls.Add(Me.Label15)
        Me.GroupBox3.Controls.Add(Me.HomeSpeedTrackbar)
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.MaxAccLimitTrackbar)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(8, 664)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(496, 200)
        Me.GroupBox3.TabIndex = 29
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Motion Parameters"
        '
        'MaxSpeedLimit
        '
        Me.MaxSpeedLimit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaxSpeedLimit.Location = New System.Drawing.Point(344, 152)
        Me.MaxSpeedLimit.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.MaxSpeedLimit.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.MaxSpeedLimit.Name = "MaxSpeedLimit"
        Me.MaxSpeedLimit.Size = New System.Drawing.Size(72, 27)
        Me.MaxSpeedLimit.TabIndex = 23
        Me.MaxSpeedLimit.Value = New Decimal(New Integer() {500, 0, 0, 0})
        '
        'HomeSpeed
        '
        Me.HomeSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HomeSpeed.Location = New System.Drawing.Point(344, 96)
        Me.HomeSpeed.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.HomeSpeed.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.HomeSpeed.Name = "HomeSpeed"
        Me.HomeSpeed.Size = New System.Drawing.Size(72, 27)
        Me.HomeSpeed.TabIndex = 22
        Me.HomeSpeed.Value = New Decimal(New Integer() {500, 0, 0, 0})
        '
        'MaxAccLimit
        '
        Me.MaxAccLimit.DecimalPlaces = 2
        Me.MaxAccLimit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaxAccLimit.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.MaxAccLimit.Location = New System.Drawing.Point(344, 40)
        Me.MaxAccLimit.Maximum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.MaxAccLimit.Minimum = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.MaxAccLimit.Name = "MaxAccLimit"
        Me.MaxAccLimit.Size = New System.Drawing.Size(72, 27)
        Me.MaxAccLimit.TabIndex = 20
        Me.MaxAccLimit.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.Location = New System.Drawing.Point(424, 152)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(70, 22)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "mm/sec"
        '
        'MaxSpeedLimitTrackbar
        '
        Me.MaxSpeedLimitTrackbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaxSpeedLimitTrackbar.LargeChange = 500
        Me.MaxSpeedLimitTrackbar.Location = New System.Drawing.Point(144, 152)
        Me.MaxSpeedLimitTrackbar.Maximum = 99900
        Me.MaxSpeedLimitTrackbar.Minimum = 100
        Me.MaxSpeedLimitTrackbar.Name = "MaxSpeedLimitTrackbar"
        Me.MaxSpeedLimitTrackbar.Size = New System.Drawing.Size(200, 45)
        Me.MaxSpeedLimitTrackbar.SmallChange = 100
        Me.MaxSpeedLimitTrackbar.TabIndex = 0
        Me.MaxSpeedLimitTrackbar.TabStop = False
        Me.MaxSpeedLimitTrackbar.TickFrequency = 5000
        Me.MaxSpeedLimitTrackbar.Value = 50000
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(8, 152)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(96, 24)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "Max Speed"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(424, 96)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(70, 22)
        Me.Label15.TabIndex = 0
        Me.Label15.Text = "mm/sec"
        '
        'HomeSpeedTrackbar
        '
        Me.HomeSpeedTrackbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HomeSpeedTrackbar.Location = New System.Drawing.Point(144, 96)
        Me.HomeSpeedTrackbar.Maximum = 500
        Me.HomeSpeedTrackbar.Minimum = 1
        Me.HomeSpeedTrackbar.Name = "HomeSpeedTrackbar"
        Me.HomeSpeedTrackbar.Size = New System.Drawing.Size(200, 45)
        Me.HomeSpeedTrackbar.TabIndex = 0
        Me.HomeSpeedTrackbar.TabStop = False
        Me.HomeSpeedTrackbar.TickFrequency = 10
        Me.HomeSpeedTrackbar.Value = 500
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(8, 96)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(128, 24)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Homing Speed"
        '
        'MaxAccLimitTrackbar
        '
        Me.MaxAccLimitTrackbar.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaxAccLimitTrackbar.Location = New System.Drawing.Point(144, 32)
        Me.MaxAccLimitTrackbar.Maximum = 100
        Me.MaxAccLimitTrackbar.Minimum = 1
        Me.MaxAccLimitTrackbar.Name = "MaxAccLimitTrackbar"
        Me.MaxAccLimitTrackbar.Size = New System.Drawing.Size(200, 45)
        Me.MaxAccLimitTrackbar.TabIndex = 0
        Me.MaxAccLimitTrackbar.TabStop = False
        Me.MaxAccLimitTrackbar.TickFrequency = 10
        Me.MaxAccLimitTrackbar.Value = 1
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(8, 40)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(144, 24)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Max Acceleration"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(424, 40)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(27, 22)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "G"
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label24.Location = New System.Drawing.Point(0, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(152, 32)
        Me.Label24.TabIndex = 54
        Me.Label24.Text = "Gantry Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 0
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label32)
        Me.GroupBox2.Controls.Add(Me.Label33)
        Me.GroupBox2.Controls.Add(Me.Label29)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.Label22)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.WorkAreaY)
        Me.GroupBox2.Controls.Add(Me.WorkAreaX)
        Me.GroupBox2.Controls.Add(Me.Label20)
        Me.GroupBox2.Controls.Add(Me.WorkAreaZMin)
        Me.GroupBox2.Controls.Add(Me.Label23)
        Me.GroupBox2.Controls.Add(Me.WorkAreaZMax)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 488)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(240, 168)
        Me.GroupBox2.TabIndex = 28
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Work Area"
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label32.Location = New System.Drawing.Point(168, 128)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(40, 23)
        Me.Label32.TabIndex = 40
        Me.Label32.Text = "mm"
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label33.Location = New System.Drawing.Point(168, 96)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(40, 23)
        Me.Label33.TabIndex = 39
        Me.Label33.Text = "mm"
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(168, 64)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(40, 23)
        Me.Label29.TabIndex = 37
        Me.Label29.Text = "mm"
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(8, 64)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(64, 23)
        Me.Label21.TabIndex = 21
        Me.Label21.Text = "Length"
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.Location = New System.Drawing.Point(8, 32)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(56, 23)
        Me.Label22.TabIndex = 20
        Me.Label22.Text = "Width"
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(168, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(40, 23)
        Me.Label16.TabIndex = 0
        Me.Label16.Text = "mm"
        '
        'WorkAreaY
        '
        Me.WorkAreaY.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.WorkAreaY.Location = New System.Drawing.Point(88, 64)
        Me.WorkAreaY.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.WorkAreaY.Name = "WorkAreaY"
        Me.WorkAreaY.Size = New System.Drawing.Size(80, 27)
        Me.WorkAreaY.TabIndex = 25
        Me.WorkAreaY.Value = New Decimal(New Integer() {370, 0, 0, 0})
        '
        'WorkAreaX
        '
        Me.WorkAreaX.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.WorkAreaX.Location = New System.Drawing.Point(88, 32)
        Me.WorkAreaX.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.WorkAreaX.Name = "WorkAreaX"
        Me.WorkAreaX.Size = New System.Drawing.Size(80, 27)
        Me.WorkAreaX.TabIndex = 24
        Me.WorkAreaX.Value = New Decimal(New Integer() {400, 0, 0, 0})
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(8, 96)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(56, 23)
        Me.Label20.TabIndex = 0
        Me.Label20.Text = "Z max"
        '
        'WorkAreaZMin
        '
        Me.WorkAreaZMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WorkAreaZMin.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.WorkAreaZMin.Location = New System.Drawing.Point(88, 128)
        Me.WorkAreaZMin.Maximum = New Decimal(New Integer() {200, 0, 0, 0})
        Me.WorkAreaZMin.Minimum = New Decimal(New Integer() {200, 0, 0, -2147483648})
        Me.WorkAreaZMin.Name = "WorkAreaZMin"
        Me.WorkAreaZMin.Size = New System.Drawing.Size(80, 27)
        Me.WorkAreaZMin.TabIndex = 29
        Me.WorkAreaZMin.Value = New Decimal(New Integer() {95, 0, 0, -2147483648})
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.Location = New System.Drawing.Point(8, 128)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(56, 23)
        Me.Label23.TabIndex = 28
        Me.Label23.Text = "Z min"
        '
        'WorkAreaZMax
        '
        Me.WorkAreaZMax.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WorkAreaZMax.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.WorkAreaZMax.Location = New System.Drawing.Point(88, 96)
        Me.WorkAreaZMax.Maximum = New Decimal(New Integer() {200, 0, 0, 0})
        Me.WorkAreaZMax.Minimum = New Decimal(New Integer() {200, 0, 0, -2147483648})
        Me.WorkAreaZMax.Name = "WorkAreaZMax"
        Me.WorkAreaZMax.Size = New System.Drawing.Size(80, 27)
        Me.WorkAreaZMax.TabIndex = 1
        Me.WorkAreaZMax.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'ButtonSPCancel
        '
        Me.ButtonSPCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCancel.Location = New System.Drawing.Point(392, 880)
        Me.ButtonSPCancel.Name = "ButtonSPCancel"
        Me.ButtonSPCancel.Size = New System.Drawing.Size(90, 48)
        Me.ButtonSPCancel.TabIndex = 19
        Me.ButtonSPCancel.Text = "Revert"
        '
        'ButtonSPOK
        '
        Me.ButtonSPOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPOK.Location = New System.Drawing.Point(280, 880)
        Me.ButtonSPOK.Name = "ButtonSPOK"
        Me.ButtonSPOK.Size = New System.Drawing.Size(90, 48)
        Me.ButtonSPOK.TabIndex = 18
        Me.ButtonSPOK.Text = "Save"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.SystemOriginPosX)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.SystemOriginPosZ)
        Me.GroupBox1.Controls.Add(Me.SystemOriginPosY)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(264, 488)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(240, 168)
        Me.GroupBox1.TabIndex = 28
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Origin"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 96)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(16, 23)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Z"
        '
        'SystemOriginPosX
        '
        Me.SystemOriginPosX.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.SystemOriginPosX.Location = New System.Drawing.Point(40, 32)
        Me.SystemOriginPosX.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.SystemOriginPosX.Minimum = New Decimal(New Integer() {500, 0, 0, -2147483648})
        Me.SystemOriginPosX.Name = "SystemOriginPosX"
        Me.SystemOriginPosX.Size = New System.Drawing.Size(80, 27)
        Me.SystemOriginPosX.TabIndex = 10
        Me.SystemOriginPosX.Value = New Decimal(New Integer() {190, 0, 0, -2147483648})
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Y"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "X"
        '
        'SystemOriginPosZ
        '
        Me.SystemOriginPosZ.Enabled = False
        Me.SystemOriginPosZ.Increment = New Decimal(New Integer() {0, 0, 0, 0})
        Me.SystemOriginPosZ.Location = New System.Drawing.Point(40, 96)
        Me.SystemOriginPosZ.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.SystemOriginPosZ.Name = "SystemOriginPosZ"
        Me.SystemOriginPosZ.Size = New System.Drawing.Size(80, 27)
        Me.SystemOriginPosZ.TabIndex = 12
        '
        'SystemOriginPosY
        '
        Me.SystemOriginPosY.Increment = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.SystemOriginPosY.Location = New System.Drawing.Point(40, 64)
        Me.SystemOriginPosY.Maximum = New Decimal(New Integer() {500, 0, 0, 0})
        Me.SystemOriginPosY.Minimum = New Decimal(New Integer() {500, 0, 0, -2147483648})
        Me.SystemOriginPosY.Name = "SystemOriginPosY"
        Me.SystemOriginPosY.Size = New System.Drawing.Size(80, 27)
        Me.SystemOriginPosY.TabIndex = 11
        Me.SystemOriginPosY.Value = New Decimal(New Integer() {380, 0, 0, -2147483648})
        '
        'GantrySetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "GantrySetup"
        Me.Text = "SystemParameters"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        CType(Me.MaxSpeedLimit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HomeSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxAccLimit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxSpeedLimitTrackbar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.HomeSpeedTrackbar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.MaxAccLimitTrackbar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.WorkAreaY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WorkAreaX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WorkAreaZMin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WorkAreaZMax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.SystemOriginPosX, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SystemOriginPosZ, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SystemOriginPosY, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

     

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
        If Not m_Tri.Move_Z(0) Then Exit Sub
    End Sub

    Private Sub MaxAccLimitTrackbar_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxAccLimitTrackbar.Scroll
        MaxAccLimit.Value = (MaxAccLimitTrackbar.Value / 100)
    End Sub

    Private Sub HomeSpeedTrackbar_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HomeSpeedTrackbar.Scroll
        HomeSpeed.Value = HomeSpeedTrackbar.Value
    End Sub

    Private Sub ElementXYSpeedTrackbar_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxSpeedLimitTrackbar.Scroll
        MaxSpeedLimit.Value = MaxSpeedLimitTrackbar.Value / 100
    End Sub

    Private Sub MaxAccLimit_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxAccLimit.ValueChanged
        MaxAccLimitTrackbar.Value = MaxAccLimit.Value * 100
    End Sub

    Private Sub HomeSpeed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HomeSpeed.ValueChanged
        HomeSpeedTrackbar.Value = HomeSpeed.Value
    End Sub

    Private Sub ElementXYSpeed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MaxSpeedLimit.ValueChanged
        MaxSpeedLimitTrackbar.Value = MaxSpeedLimit.Value * 100
    End Sub

    Private Sub SaveGantrySettings()

        'If Vision.Checked Then
        '    offset_x = 0
        '    offset_y = 0
        'Else
            With IDS.Data.Hardware.Needle
                If LeftHead.Checked Then
                    offset_x = .Left.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
                    offset_y = .Left.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
                ElseIf RightHead.Checked Then
                    offset_x = .Right.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
                    offset_y = .Right.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
                End If
            End With
        'End If 'yy

        m_Tri.GetIDSState()
        x = m_Tri.XPosition + offset_x
        y = m_Tri.YPosition + offset_y

        If StationPosition.SelectedItem = "Park Position" Then
            IDS.Data.Hardware.Gantry.ParkPosition.X = x
            IDS.Data.Hardware.Gantry.ParkPosition.Y = y

        ElseIf StationPosition.SelectedItem = "Purge Position" Then
            IDS.Data.Hardware.Gantry.PurgePosition.X = x
            IDS.Data.Hardware.Gantry.PurgePosition.Y = y

        ElseIf StationPosition.SelectedItem = "Clean Position" Then
            IDS.Data.Hardware.Gantry.CleanPosition.X = x
            IDS.Data.Hardware.Gantry.CleanPosition.Y = y

        ElseIf StationPosition.SelectedItem = "Change Syringe Position" Then
            IDS.Data.Hardware.Gantry.ChangeSyringePosition.X = x
            IDS.Data.Hardware.Gantry.ChangeSyringePosition.Y = y

            'ElseIf StationPosition.SelectedItem = "Volume Calibration Position" Then
            '    IDS.Data.Hardware.Gantry.WeighingScalePosition.X = x
            '    IDS.Data.Hardware.Gantry.WeighingScalePosition.Y = y
            'Else

            '    If LeftHead.Checked Then
            '        If StationPosition.SelectedItem = "Needle Calibration First Row First Dot Position" Then
            '            IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X = x
            '            IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y = y
            '        ElseIf StationPosition.SelectedItem = "Needle Calibration Last Row Last Dot Position" Then
            '            IDS.Data.Hardware.Needle.Left.ArrayDotPos3.X = x
            '            IDS.Data.Hardware.Needle.Left.ArrayDotPos3.Y = y
            '        End If
            '    ElseIf RightHead.Checked Then
            '        If StationPosition.SelectedItem = "Needle Calibration First Row First Dot Position" Then
            '            IDS.Data.Hardware.Needle.Right.ArrayDotPos1.X = x
            '            IDS.Data.Hardware.Needle.Right.ArrayDotPos1.Y = y
            '        ElseIf StationPosition.SelectedItem = "Needle Calibration Last Row Last Dot Position" Then
            '            IDS.Data.Hardware.Needle.Right.ArrayDotPos3.X = x
            '            IDS.Data.Hardware.Needle.Right.ArrayDotPos3.Y = y
            '        End If
            '    End If 'yy

        End If

        XPosition.Text = "X: " + CStr(x)
        YPosition.Text = "Y: " + CStr(y)

    End Sub

    Private Sub ButtonSPOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPOK.Click
        SaveData()
    End Sub

    Private Sub ButtonSPCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPCancel.Click
        RevertData()
    End Sub

    Public Sub SaveData()
        'motion parameters - homing speed & acceleration left out
        IDS.Data.Hardware.Gantry.MaxSpeedLimit = MaxSpeedLimit.Text

        'sys origin
        IDS.Data.Hardware.Gantry.SystemOriginPos.X = SystemOriginPosX.Text
        IDS.Data.Hardware.Gantry.SystemOriginPos.Y = SystemOriginPosY.Text
        IDS.Data.Hardware.Gantry.SystemOriginPos.Z = SystemOriginPosZ.Text

        'work area
        IDS.Data.Hardware.Gantry.WorkArea.Z.Max = WorkAreaZMax.Text
        IDS.Data.Hardware.Gantry.WorkArea.Z.Min = WorkAreaZMin.Text
        IDS.Data.Hardware.Gantry.WorkArea.X = WorkAreaX.Text
        IDS.Data.Hardware.Gantry.WorkArea.Y = WorkAreaY.Text

        IDS.Data.SaveLocalData()
    End Sub

    Public Sub RevertData()

        IDS.Data.OpenData()

        SystemOriginPosX.Text = IDS.Data.Hardware.Gantry.SystemOriginPos.X
        SystemOriginPosY.Text = IDS.Data.Hardware.Gantry.SystemOriginPos.Y
        SystemOriginPosZ.Text = IDS.Data.Hardware.Gantry.SystemOriginPos.Z

        WorkAreaZMax.Text = IDS.Data.Hardware.Gantry.WorkArea.Z.Max
        WorkAreaZMin.Text = IDS.Data.Hardware.Gantry.WorkArea.Z.Min
        WorkAreaX.Text = IDS.Data.Hardware.Gantry.WorkArea.X
        WorkAreaY.Text = IDS.Data.Hardware.Gantry.WorkArea.Y
        MaxAccLimitTrackbar.Value = MaxAccLimit.Value * 100

    End Sub

    Private Sub MoveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveButton.Click

        SetServiceSpeed()
        If Not m_Tri.Move_Z(0) Then Exit Sub

        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem = "Park Position" Then
                x = .ParkPosition.X
                y = .ParkPosition.Y
            ElseIf StationPosition.SelectedItem = "Purge Position" Then
                x = .PurgePosition.X
                y = .PurgePosition.Y
            ElseIf StationPosition.SelectedItem = "Clean Position" Then
                x = .CleanPosition.X
                y = .CleanPosition.Y
            ElseIf StationPosition.SelectedItem = "Change Syringe Position" Then
                x = .ChangeSyringePosition.X
                y = .ChangeSyringePosition.Y
                'ElseIf StationPosition.SelectedItem = "Volume Calibration Position" Then
                '    x = .WeighingScalePosition.X
                '    y = .WeighingScalePosition.Y 'yy
            End If
        End With

        'With IDS.Data.Hardware.Needle
        '    If LeftHead.Checked Then
        '        offset_x = .Left.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
        '        offset_y = .Left.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
        '        If StationPosition.SelectedItem = "Needle Calibration First Row First Dot Position" Then
        '            x = .Left.ArrayDotPos1.X
        '            y = .Left.ArrayDotPos1.Y
        '        ElseIf StationPosition.SelectedItem = "Needle Calibration Last Row Last Dot Position" Then
        '            x = .Left.ArrayDotPos3.X
        '            y = .Left.ArrayDotPos3.Y
        '        End If
        '    ElseIf RightHead.Checked Then
        '        offset_x = .Right.CalibratorPos.X - .Right.NeedleCalibrationPosition.X
        '        offset_y = .Right.CalibratorPos.Y - .Right.NeedleCalibrationPosition.Y
        '        If StationPosition.SelectedItem = "Needle Calibration First Row First Dot Position" Then
        '            x = .Right.ArrayDotPos1.X
        '            y = .Right.ArrayDotPos1.Y
        '        ElseIf StationPosition.SelectedItem = "Needle Calibration Last Row Last Dot Position" Then
        '            x = .Right.ArrayDotPos3.X
        '            y = .Right.ArrayDotPos3.Y
        '        End If
        '    End If
        'End With 'yy

        'If Vision.Checked Then
        '    position(0) = x
        '    position(1) = y
        'ElseIf Needle.Checked Then
            position(0) = x - offset_x
            position(1) = y - offset_y
        'End If

        If Not m_Tri.Move_XY(position) Then Exit Sub

    End Sub

    Private Sub SavePositionButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SavePositionButton.Click

        'station settings
        SaveGantrySettings()
        IDS.Data.SaveLocalData()

    End Sub

    Private Sub StationPosition_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StationPosition.SelectedIndexChanged

        If StationPosition.SelectedItem = "Needle Calibration Position" Then
            'Vision.Enabled = False
            'If Vision.Checked Then
            'Vision.Checked = False
            LeftHead.Checked = True
            'End If
            'Else
            '    Vision.Enabled = True 'yy
        End If

        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem = "Park Position" Then
                XPosition.Text = "X: " + CStr(.ParkPosition.X)
                YPosition.Text = "Y: " + CStr(.ParkPosition.Y)
            ElseIf StationPosition.SelectedItem = "Purge Position" Then
                XPosition.Text = "X: " + CStr(.PurgePosition.X)
                YPosition.Text = "Y: " + CStr(.PurgePosition.Y)
            ElseIf StationPosition.SelectedItem = "Clean Position" Then
                XPosition.Text = "X: " + CStr(.CleanPosition.X)
                YPosition.Text = "Y: " + CStr(.CleanPosition.Y)
            ElseIf StationPosition.SelectedItem = "Change Syringe Position" Then
                XPosition.Text = "X: " + CStr(.ChangeSyringePosition.X)
                YPosition.Text = "Y: " + CStr(.ChangeSyringePosition.Y)
                'ElseIf StationPosition.SelectedItem = "Volume Calibration Position" Then
                '    XPosition.Text = "X: " + CStr(.WeighingScalePosition.X)
                '    YPosition.Text = "Y: " + CStr(.WeighingScalePosition.Y)
                'ElseIf StationPosition.SelectedItem = "Needle Calibration First Row First Dot Position" Then
                '    If LeftHead.Checked Then
                '        XPosition.Text = "X: " + CStr(IDS.Data.Hardware.Needle.Left.ArrayDotPos1.X)
                '        YPosition.Text = "Y: " + CStr(IDS.Data.Hardware.Needle.Left.ArrayDotPos1.Y)
                '    ElseIf RightHead.Checked Then
                '        XPosition.Text = "X: " + CStr(IDS.Data.Hardware.Needle.Right.ArrayDotPos1.X)
                '        YPosition.Text = "Y: " + CStr(IDS.Data.Hardware.Needle.Right.ArrayDotPos1.Y)
                '    End If
                'ElseIf StationPosition.SelectedItem = "Needle Calibration Last Row Last Dot Position" Then
                '    If LeftHead.Checked Then
                '        XPosition.Text = "X: " + CStr(IDS.Data.Hardware.Needle.Left.ArrayDotPos3.X)
                '        YPosition.Text = "Y: " + CStr(IDS.Data.Hardware.Needle.Left.ArrayDotPos3.Y)
                '    ElseIf RightHead.Checked Then
                '        XPosition.Text = "X: " + CStr(IDS.Data.Hardware.Needle.Right.ArrayDotPos3.X)
                '        YPosition.Text = "Y: " + CStr(IDS.Data.Hardware.Needle.Right.ArrayDotPos3.Y)
                '    End If 'yy
            End If
        End With

    End Sub

    'Private Sub Vision_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Vision.Click
    '    Needle.Checked = Not Vision.Checked
    'End Sub

    'Private Sub Needle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Needle.Click
    '    Vision.Checked = Not Needle.Checked
    'End Sub 'yy

    Private Sub LeftHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeftHead.Click
        LeftHead.Checked = Not RightHead.Checked
    End Sub

    Private Sub RightHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RightHead.Click
        RightHead.Checked = Not LeftHead.Checked
    End Sub
End Class
