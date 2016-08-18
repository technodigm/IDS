Imports DLL_Export_Service

Public Class GantrySettings
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
    Friend WithEvents CB_UseSystemDefault As System.Windows.Forms.CheckBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents gpbDualHead As System.Windows.Forms.GroupBox
    Public WithEvents chkDualHead As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents SavePositionButton As System.Windows.Forms.Button
    Friend WithEvents StationPosition As System.Windows.Forms.ComboBox
    Friend WithEvents MoveButton As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Public WithEvents Vision As System.Windows.Forms.RadioButton
    Public WithEvents Needle As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonTeachPos As System.Windows.Forms.Button
    Friend WithEvents ButtonGotoUserInput As System.Windows.Forms.Button
    Friend WithEvents ButtonGotoPositionDefined As System.Windows.Forms.Button
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Public WithEvents RightHead As System.Windows.Forms.RadioButton
    Public WithEvents LeftHead As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ZPosition As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ElementXYSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents ElementZSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ServiceZSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents ServiceXYSpeed As System.Windows.Forms.NumericUpDown
    Friend WithEvents btMoveXYOnly As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(GantrySettings))
        Me.CB_UseSystemDefault = New System.Windows.Forms.CheckBox
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.SavePositionButton = New System.Windows.Forms.Button
        Me.ZPosition = New System.Windows.Forms.Label
        Me.StationPosition = New System.Windows.Forms.ComboBox
        Me.MoveButton = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.Vision = New System.Windows.Forms.RadioButton
        Me.Needle = New System.Windows.Forms.RadioButton
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
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ElementZSpeed = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.ElementXYSpeed = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ServiceZSpeed = New System.Windows.Forms.NumericUpDown
        Me.ServiceXYSpeed = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.gpbDualHead = New System.Windows.Forms.GroupBox
        Me.chkDualHead = New System.Windows.Forms.CheckBox
        Me.btMoveXYOnly = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.ElementZSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ElementXYSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ServiceZSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ServiceXYSpeed, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpbDualHead.SuspendLayout()
        Me.SuspendLayout()
        '
        'CB_UseSystemDefault
        '
        Me.CB_UseSystemDefault.Location = New System.Drawing.Point(112, 952)
        Me.CB_UseSystemDefault.Name = "CB_UseSystemDefault"
        Me.CB_UseSystemDefault.Size = New System.Drawing.Size(168, 24)
        Me.CB_UseSystemDefault.TabIndex = 29
        Me.CB_UseSystemDefault.Text = "Use System Default Position"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox6)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(40, 24)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 17
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.btMoveXYOnly)
        Me.GroupBox6.Controls.Add(Me.SavePositionButton)
        Me.GroupBox6.Controls.Add(Me.ZPosition)
        Me.GroupBox6.Controls.Add(Me.StationPosition)
        Me.GroupBox6.Controls.Add(Me.MoveButton)
        Me.GroupBox6.Controls.Add(Me.GroupBox4)
        Me.GroupBox6.Controls.Add(Me.GroupBox7)
        Me.GroupBox6.Location = New System.Drawing.Point(8, 400)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(496, 352)
        Me.GroupBox6.TabIndex = 69
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Station Z Position Setup"
        '
        'SavePositionButton
        '
        Me.SavePositionButton.Location = New System.Drawing.Point(164, 288)
        Me.SavePositionButton.Name = "SavePositionButton"
        Me.SavePositionButton.Size = New System.Drawing.Size(168, 48)
        Me.SavePositionButton.TabIndex = 69
        Me.SavePositionButton.Text = "Save Current Position"
        '
        'ZPosition
        '
        Me.ZPosition.Location = New System.Drawing.Point(16, 176)
        Me.ZPosition.Name = "ZPosition"
        Me.ZPosition.Size = New System.Drawing.Size(464, 23)
        Me.ZPosition.TabIndex = 66
        Me.ZPosition.Text = "0"
        Me.ZPosition.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'StationPosition
        '
        Me.StationPosition.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.StationPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.StationPosition.Items.AddRange(New Object() {"Park Z Position", "Purge Z Position", "Clean Z Position", "Change Syringe Z Position", "Volume Calibration Z Position"})
        Me.StationPosition.Location = New System.Drawing.Point(72, 120)
        Me.StationPosition.Name = "StationPosition"
        Me.StationPosition.Size = New System.Drawing.Size(368, 28)
        Me.StationPosition.TabIndex = 64
        '
        'MoveButton
        '
        Me.MoveButton.Location = New System.Drawing.Point(272, 224)
        Me.MoveButton.Name = "MoveButton"
        Me.MoveButton.Size = New System.Drawing.Size(168, 48)
        Me.MoveButton.TabIndex = 65
        Me.MoveButton.Text = "Move to Saved Station XYZ"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Vision)
        Me.GroupBox4.Controls.Add(Me.Needle)
        Me.GroupBox4.Controls.Add(Me.GroupBox5)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(24, 40)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(200, 64)
        Me.GroupBox4.TabIndex = 61
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Mode"
        '
        'Vision
        '
        Me.Vision.Checked = True
        Me.Vision.Location = New System.Drawing.Point(112, 32)
        Me.Vision.Name = "Vision"
        Me.Vision.Size = New System.Drawing.Size(80, 24)
        Me.Vision.TabIndex = 3
        Me.Vision.TabStop = True
        Me.Vision.Text = "Vision"
        '
        'Needle
        '
        Me.Needle.Location = New System.Drawing.Point(16, 32)
        Me.Needle.Name = "Needle"
        Me.Needle.Size = New System.Drawing.Size(88, 24)
        Me.Needle.TabIndex = 2
        Me.Needle.Text = "Needle"
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
        Me.GroupBox7.Location = New System.Drawing.Point(272, 40)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(200, 64)
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
        Me.Label17.Size = New System.Drawing.Size(176, 32)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Gantry Settings"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ElementZSpeed)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ElementXYSpeed)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.ServiceZSpeed)
        Me.GroupBox1.Controls.Add(Me.ServiceXYSpeed)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 208)
        Me.GroupBox1.TabIndex = 69
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Gantry Speed"
        '
        'ElementZSpeed
        '
        Me.ElementZSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ElementZSpeed.Location = New System.Drawing.Point(273, 40)
        Me.ElementZSpeed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.ElementZSpeed.Name = "ElementZSpeed"
        Me.ElementZSpeed.Size = New System.Drawing.Size(72, 27)
        Me.ElementZSpeed.TabIndex = 75
        Me.ElementZSpeed.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(81, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(168, 22)
        Me.Label2.TabIndex = 73
        Me.Label2.Text = "Dispensing Z Speed"
        '
        'ElementXYSpeed
        '
        Me.ElementXYSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ElementXYSpeed.Location = New System.Drawing.Point(273, 80)
        Me.ElementXYSpeed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.ElementXYSpeed.Name = "ElementXYSpeed"
        Me.ElementXYSpeed.Size = New System.Drawing.Size(72, 27)
        Me.ElementXYSpeed.TabIndex = 75
        Me.ElementXYSpeed.Value = New Decimal(New Integer() {500, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(81, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(184, 22)
        Me.Label3.TabIndex = 73
        Me.Label3.Text = "Dispensing XY Speed"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(345, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 22)
        Me.Label1.TabIndex = 73
        Me.Label1.Text = "mm/sec"
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(345, 80)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(70, 22)
        Me.Label15.TabIndex = 73
        Me.Label15.Text = "mm/sec"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(345, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(70, 22)
        Me.Label4.TabIndex = 73
        Me.Label4.Text = "mm/sec"
        '
        'ServiceZSpeed
        '
        Me.ServiceZSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ServiceZSpeed.Location = New System.Drawing.Point(273, 120)
        Me.ServiceZSpeed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.ServiceZSpeed.Name = "ServiceZSpeed"
        Me.ServiceZSpeed.Size = New System.Drawing.Size(72, 27)
        Me.ServiceZSpeed.TabIndex = 75
        Me.ServiceZSpeed.Value = New Decimal(New Integer() {50, 0, 0, 0})
        '
        'ServiceXYSpeed
        '
        Me.ServiceXYSpeed.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ServiceXYSpeed.Location = New System.Drawing.Point(273, 160)
        Me.ServiceXYSpeed.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.ServiceXYSpeed.Name = "ServiceXYSpeed"
        Me.ServiceXYSpeed.Size = New System.Drawing.Size(72, 27)
        Me.ServiceXYSpeed.TabIndex = 75
        Me.ServiceXYSpeed.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(81, 160)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(152, 22)
        Me.Label5.TabIndex = 73
        Me.Label5.Text = "Service XY Speed"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(345, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 22)
        Me.Label6.TabIndex = 73
        Me.Label6.Text = "mm/sec"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(81, 120)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(136, 22)
        Me.Label7.TabIndex = 73
        Me.Label7.Text = "Service Z Speed"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(392, 320)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(88, 48)
        Me.ButtonRevert.TabIndex = 77
        Me.ButtonRevert.Text = "Revert"
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(280, 320)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(88, 48)
        Me.ButtonSave.TabIndex = 76
        Me.ButtonSave.Text = "Save"
        '
        'gpbDualHead
        '
        Me.gpbDualHead.Controls.Add(Me.chkDualHead)
        Me.gpbDualHead.Location = New System.Drawing.Point(280, 944)
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
        'btMoveXYOnly
        '
        Me.btMoveXYOnly.Location = New System.Drawing.Point(56, 224)
        Me.btMoveXYOnly.Name = "btMoveXYOnly"
        Me.btMoveXYOnly.Size = New System.Drawing.Size(168, 48)
        Me.btMoveXYOnly.TabIndex = 70
        Me.btMoveXYOnly.Text = "Move to  Saved Station XY Only"
        '
        'GantrySettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Controls.Add(Me.CB_UseSystemDefault)
        Me.Controls.Add(Me.gpbDualHead)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "GantrySettings"
        Me.Text = "GantrySettings"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ElementZSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ElementXYSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ServiceZSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ServiceXYSpeed, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpbDualHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub SaveGantrySettings()

        If Vision.Checked Then
            offset_x = 0
            offset_y = 0
        Else
            With IDS.Data.Hardware.Needle
                If LeftHead.Checked Then
                    offset_z = .Left.NeedleCalibrationPosition.Z
                ElseIf RightHead.Checked Then
                    offset_z = .Right.NeedleCalibrationPosition.Z
                End If
            End With
        End If

        m_Tri.GetIDSState()
        z = m_Tri.ZPosition - offset_z

        If StationPosition.SelectedItem.ToString = "Park Z Position" Then
            IDS.Data.Hardware.Gantry.ParkPosition.Z = z
        ElseIf StationPosition.SelectedItem.ToString = "Purge Z Position" Then
            IDS.Data.Hardware.Gantry.PurgePosition.Z = z
        ElseIf StationPosition.SelectedItem.ToString = "Clean Z Position" Then
            IDS.Data.Hardware.Gantry.CleanPosition.Z = z
        ElseIf StationPosition.SelectedItem.ToString = "Change Syringe Z Position" Then
            IDS.Data.Hardware.Gantry.ChangeSyringePosition.Z = z
        ElseIf StationPosition.SelectedItem.ToString = "Volume Calibration Z Position" Then
            IDS.Data.Hardware.Gantry.WeighingScalePosition.Z = z
        End If

        'ZPosition.Text = "Current Z:" & m_Tri.ZPosition & " - Needle Z Offset:" & offset_z & " = Calibrated Z: " + CStr(z)
        'Park Z Position
        'Purge Z Position
        'Clean Z Position
        'Change Syringe Z Position
        'Volume Calibration Z Position
        Dim zOff As Double = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z
        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem = "Park Z Position" Then
                ZPosition.Text = "Z: " + (.ParkPosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem = "Purge Z Position" Then
                ZPosition.Text = "Z: " + (.PurgePosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem = "Clean Z Position" Then
                ZPosition.Text = "Z: " + (.CleanPosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem = "Change Syringe Z Position" Then
                ZPosition.Text = "Z: " + (.ChangeSyringePosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem = "Volume Calibration Z Position" Then
                ZPosition.Text = "Z: " + (.WeighingScalePosition.Z + zOff).ToString("0.000")
            End If
        End With
    End Sub

    Private Sub MoveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveButton.Click

        SetServiceSpeed()
        MoveButton.Enabled = False
        btMoveXYOnly.Enabled = False
        If Not m_Tri.Move_Z(0) Then GoTo Reset

        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem.ToString = "Park Z Position" Then
                x = .ParkPosition.X
                y = .ParkPosition.Y
                z = .ParkPosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Purge Z Position" Then
                x = .PurgePosition.X
                y = .PurgePosition.Y
                z = .PurgePosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Clean Z Position" Then
                x = .CleanPosition.X
                y = .CleanPosition.Y
                z = .CleanPosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Change Syringe Z Position" Then
                x = .ChangeSyringePosition.X
                y = .ChangeSyringePosition.Y
                z = .ChangeSyringePosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Volume Calibration Z Position" Then
                x = .WeighingScalePosition.X
                y = .WeighingScalePosition.Y
                z = .WeighingScalePosition.Z
            End If
        End With

        With IDS.Data.Hardware.Needle
            If LeftHead.Checked Then
                offset_x = .Left.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
                offset_y = .Left.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
                offset_z = .Left.NeedleCalibrationPosition.Z
            ElseIf RightHead.Checked Then
                offset_x = .Right.CalibratorPos.X - .Right.NeedleCalibrationPosition.X
                offset_y = .Right.CalibratorPos.Y - .Right.NeedleCalibrationPosition.Y
                offset_z = .Right.NeedleCalibrationPosition.Z
            End If
        End With

        If Vision.Checked Then
            position(0) = x
            position(1) = y
            position(2) = z
        ElseIf Needle.Checked Then
            position(0) = x - offset_x
            position(1) = y - offset_y
            position(2) = z + offset_z
        End If

        If Not m_Tri.Move_Z(0) Then GoTo Reset
        If Not m_Tri.Move_XY(position) Then GoTo Reset
        If Needle.Checked Then
            If Not m_Tri.Move_Z(position(2)) Then GoTo Reset
        End If
Reset:
        MoveButton.Enabled = True
        btMoveXYOnly.Enabled = True
    End Sub

    Private Sub SavePositionButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SavePositionButton.Click

        'station settings
        SaveGantrySettings()
        IDS.Data.SaveData()

    End Sub

    Private Sub StationPosition_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StationPosition.SelectedIndexChanged

        'If StationPosition.SelectedItem.ToString = "Needle Calibration Z Position" Then
        '    Vision.Enabled = False
        '    If Vision.Checked Then
        '        Vision.Checked = False
        '        LeftHead.Checked = True
        '    End If
        'Else
        '    Vision.Enabled = True
        'End If
        Dim zOff As Double = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z
        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem.ToString = "Park Z Position" Then
                ZPosition.Text = "Z: " + (.ParkPosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem.ToString = "Purge Z Position" Then
                ZPosition.Text = "Z: " + (.PurgePosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem.ToString = "Clean Z Position" Then
                ZPosition.Text = "Z: " + (.CleanPosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem.ToString = "Change Syringe Z Position" Then
                ZPosition.Text = "Z: " + (.ChangeSyringePosition.Z + zOff).ToString("0.000")
            ElseIf StationPosition.SelectedItem.ToString = "Volume Calibration Z Position" Then
                ZPosition.Text = "Z: " + (.WeighingScalePosition.Z + zOff).ToString("0.000")
            End If
        End With

    End Sub

    Private Sub Vision_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Needle.Checked = Not Vision.Checked
        SavePositionButton.Enabled = Not Vision.Checked
    End Sub

    Private Sub Needle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Needle.Click
        Vision.Checked = Not Needle.Checked
        SavePositionButton.Enabled = Not Vision.Checked
    End Sub

    Private Sub LeftHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LeftHead.Click
        LeftHead.Checked = Not RightHead.Checked
    End Sub

    Private Sub RightHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RightHead.Click
        RightHead.Checked = Not LeftHead.Checked
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click

        RemovePanel(CurrentControl)
        If Not m_Tri.Move_Z(0) Then Exit Sub

    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        SaveData()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub

    Public Sub SaveData()
        IDS.Data.Hardware.Gantry.ElementXYSpeed = ElementXYSpeed.Value
        IDS.Data.Hardware.Gantry.ElementZSpeed = ElementZSpeed.Value
        IDS.Data.Hardware.Gantry.ServiceXYSpeed = ServiceXYSpeed.Value
        IDS.Data.Hardware.Gantry.ServiceZSpeed = ServiceZSpeed.Value
        IDS.Data.SaveData()
    End Sub

    Public Sub RevertData()
        IDS.Data.OpenData()
        ElementXYSpeed.Text = IDS.Data.Hardware.Gantry.ElementXYSpeed
        ElementZSpeed.Text = IDS.Data.Hardware.Gantry.ElementZSpeed
        ServiceXYSpeed.Text = IDS.Data.Hardware.Gantry.ServiceXYSpeed
        ServiceZSpeed.Text = IDS.Data.Hardware.Gantry.ServiceZSpeed
    End Sub

    Private Sub ElementXYSpeed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ElementXYSpeed.ValueChanged

    End Sub

    Private Sub Vision_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Vision.Click
        Needle.Checked = Not Vision.Checked
        SavePositionButton.Enabled = Not Vision.Checked
        MoveButton.Enabled = Not Vision.Checked
    End Sub

    Private Sub Vision_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Vision.CheckedChanged
        Needle.Checked = Not Vision.Checked
        SavePositionButton.Enabled = Not Vision.Checked
        MoveButton.Enabled = Not Vision.Checked
    End Sub

    Private Sub btMoveXYOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btMoveXYOnly.Click

        SetServiceSpeed()
        btMoveXYOnly.Enabled = False
        MoveButton.Enabled = False
        With IDS.Data.Hardware.Gantry
            If StationPosition.SelectedItem.ToString = "Park Z Position" Then
                x = .ParkPosition.X
                y = .ParkPosition.Y
                z = .ParkPosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Purge Z Position" Then
                x = .PurgePosition.X
                y = .PurgePosition.Y
                z = .PurgePosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Clean Z Position" Then
                x = .CleanPosition.X
                y = .CleanPosition.Y
                z = .CleanPosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Change Syringe Z Position" Then
                x = .ChangeSyringePosition.X
                y = .ChangeSyringePosition.Y
                z = .ChangeSyringePosition.Z
            ElseIf StationPosition.SelectedItem.ToString = "Volume Calibration Z Position" Then
                x = .WeighingScalePosition.X
                y = .WeighingScalePosition.Y
                z = .WeighingScalePosition.Z
            End If
        End With

        With IDS.Data.Hardware.Needle
            If LeftHead.Checked Then
                offset_x = .Left.CalibratorPos.X - .Left.NeedleCalibrationPosition.X
                offset_y = .Left.CalibratorPos.Y - .Left.NeedleCalibrationPosition.Y
                offset_z = .Left.NeedleCalibrationPosition.Z
            ElseIf RightHead.Checked Then
                offset_x = .Right.CalibratorPos.X - .Right.NeedleCalibrationPosition.X
                offset_y = .Right.CalibratorPos.Y - .Right.NeedleCalibrationPosition.Y
                offset_z = .Right.NeedleCalibrationPosition.Z
            End If
        End With

        If Vision.Checked Then
            position(0) = x
            position(1) = y
            position(2) = z
        ElseIf Needle.Checked Then
            position(0) = x - offset_x
            position(1) = y - offset_y
            position(2) = z + offset_z
        End If

        If Not m_Tri.Move_Z(0) Then GoTo Reset
        If Not m_Tri.Move_XY(position) Then GoTo Reset
Reset:
        btMoveXYOnly.Enabled = True
        MoveButton.Enabled = True
    End Sub
End Class
