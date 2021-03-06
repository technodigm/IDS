Imports DLL_Export_Service

Public Class ConveyorSettings
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
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents gpbDualHead As System.Windows.Forms.GroupBox
    Public WithEvents chkDualHead As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents WidthDisplay As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ButtonConveyorHome As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorMoveRev As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorMoveForward As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorStart As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TimeOut As System.Windows.Forms.NumericUpDown
    Friend WithEvents Speed As System.Windows.Forms.NumericUpDown
    Friend WithEvents WidthMoveStep As System.Windows.Forms.NumericUpDown
    Friend WithEvents Width As System.Windows.Forms.NumericUpDown
    Friend WithEvents PositionTimer As System.Timers.Timer
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents upstreamTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents downstreamTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(ConveyorSettings))
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.downstreamTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.upstreamTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TimeOut = New System.Windows.Forms.NumericUpDown
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Speed = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.WidthDisplay = New System.Windows.Forms.Label
        Me.WidthMoveStep = New System.Windows.Forms.NumericUpDown
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ButtonConveyorHome = New System.Windows.Forms.Button
        Me.ButtonConveyorMoveRev = New System.Windows.Forms.Button
        Me.ButtonConveyorMoveForward = New System.Windows.Forms.Button
        Me.ButtonConveyorStart = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.Width = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.gpbDualHead = New System.Windows.Forms.GroupBox
        Me.chkDualHead = New System.Windows.Forms.CheckBox
        Me.PositionTimer = New System.Timers.Timer
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.downstreamTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.upstreamTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TimeOut, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Speed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.WidthMoveStep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Width, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpbDualHead.SuspendLayout()
        CType(Me.PositionTimer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!)
        Me.ButtonRevert.Location = New System.Drawing.Point(408, 464)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(88, 48)
        Me.ButtonRevert.TabIndex = 20
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!)
        Me.ButtonSave.Location = New System.Drawing.Point(296, 464)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(88, 48)
        Me.ButtonSave.TabIndex = 19
        Me.ButtonSave.Text = "Save"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 5
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(408, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 48
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.downstreamTimeout)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.upstreamTimeout)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TimeOut)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.Speed)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.WidthDisplay)
        Me.GroupBox1.Controls.Add(Me.WidthMoveStep)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.ButtonConveyorHome)
        Me.GroupBox1.Controls.Add(Me.ButtonConveyorMoveRev)
        Me.GroupBox1.Controls.Add(Me.ButtonConveyorMoveForward)
        Me.GroupBox1.Controls.Add(Me.ButtonConveyorStart)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Width)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 88)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 360)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Conveyor Information"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label12.Location = New System.Drawing.Point(256, 296)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(24, 23)
        Me.Label12.TabIndex = 53
        Me.Label12.Text = "s"
        '
        'downstreamTimeout
        '
        Me.downstreamTimeout.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.downstreamTimeout.Location = New System.Drawing.Point(192, 296)
        Me.downstreamTimeout.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.downstreamTimeout.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.downstreamTimeout.Name = "downstreamTimeout"
        Me.downstreamTimeout.Size = New System.Drawing.Size(64, 27)
        Me.downstreamTimeout.TabIndex = 52
        Me.downstreamTimeout.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label13.Location = New System.Drawing.Point(16, 296)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(176, 23)
        Me.Label13.TabIndex = 54
        Me.Label13.Text = "downstream Timeout"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(256, 264)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(24, 23)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "s"
        '
        'upstreamTimeout
        '
        Me.upstreamTimeout.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.upstreamTimeout.Location = New System.Drawing.Point(192, 264)
        Me.upstreamTimeout.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.upstreamTimeout.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.upstreamTimeout.Name = "upstreamTimeout"
        Me.upstreamTimeout.Size = New System.Drawing.Size(64, 27)
        Me.upstreamTimeout.TabIndex = 49
        Me.upstreamTimeout.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(16, 264)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(168, 23)
        Me.Label8.TabIndex = 51
        Me.Label8.Text = "Upstream Timeout"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(256, 232)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 23)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "s"
        '
        'TimeOut
        '
        Me.TimeOut.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TimeOut.Location = New System.Drawing.Point(192, 232)
        Me.TimeOut.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.TimeOut.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.TimeOut.Name = "TimeOut"
        Me.TimeOut.Size = New System.Drawing.Size(64, 27)
        Me.TimeOut.TabIndex = 46
        Me.TimeOut.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 232)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(144, 23)
        Me.Label9.TabIndex = 48
        Me.Label9.Text = "Retrieve Timeout"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(256, 184)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 24)
        Me.Label11.TabIndex = 45
        Me.Label11.Text = "mm/s"
        '
        'Speed
        '
        Me.Speed.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Speed.Location = New System.Drawing.Point(192, 184)
        Me.Speed.Maximum = New Decimal(New Integer() {400, 0, 0, 0})
        Me.Speed.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.Speed.Name = "Speed"
        Me.Speed.Size = New System.Drawing.Size(64, 27)
        Me.Speed.TabIndex = 44
        Me.Speed.Value = New Decimal(New Integer() {300, 0, 0, 0})
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 184)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 23)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "Speed"
        '
        'WidthDisplay
        '
        Me.WidthDisplay.BackColor = System.Drawing.Color.Transparent
        Me.WidthDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.WidthDisplay.Location = New System.Drawing.Point(208, 40)
        Me.WidthDisplay.Name = "WidthDisplay"
        Me.WidthDisplay.Size = New System.Drawing.Size(272, 23)
        Me.WidthDisplay.TabIndex = 18
        '
        'WidthMoveStep
        '
        Me.WidthMoveStep.DecimalPlaces = 1
        Me.WidthMoveStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.WidthMoveStep.Increment = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.WidthMoveStep.Location = New System.Drawing.Point(168, 136)
        Me.WidthMoveStep.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.WidthMoveStep.Minimum = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.WidthMoveStep.Name = "WidthMoveStep"
        Me.WidthMoveStep.Size = New System.Drawing.Size(64, 27)
        Me.WidthMoveStep.TabIndex = 21
        Me.WidthMoveStep.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Current Width:"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 136)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 23)
        Me.Label4.TabIndex = 20
        Me.Label4.Text = "Relative Move"
        '
        'ButtonConveyorHome
        '
        Me.ButtonConveyorHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonConveyorHome.Location = New System.Drawing.Point(392, 80)
        Me.ButtonConveyorHome.Name = "ButtonConveyorHome"
        Me.ButtonConveyorHome.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorHome.TabIndex = 18
        Me.ButtonConveyorHome.Text = "Home"
        '
        'ButtonConveyorMoveRev
        '
        Me.ButtonConveyorMoveRev.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonConveyorMoveRev.Location = New System.Drawing.Point(392, 128)
        Me.ButtonConveyorMoveRev.Name = "ButtonConveyorMoveRev"
        Me.ButtonConveyorMoveRev.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorMoveRev.TabIndex = 23
        Me.ButtonConveyorMoveRev.Text = ">>--<<"
        '
        'ButtonConveyorMoveForward
        '
        Me.ButtonConveyorMoveForward.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonConveyorMoveForward.Location = New System.Drawing.Point(288, 128)
        Me.ButtonConveyorMoveForward.Name = "ButtonConveyorMoveForward"
        Me.ButtonConveyorMoveForward.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorMoveForward.TabIndex = 22
        Me.ButtonConveyorMoveForward.Text = "<<-->>"
        '
        'ButtonConveyorStart
        '
        Me.ButtonConveyorStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonConveyorStart.Location = New System.Drawing.Point(288, 80)
        Me.ButtonConveyorStart.Name = "ButtonConveyorStart"
        Me.ButtonConveyorStart.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorStart.TabIndex = 19
        Me.ButtonConveyorStart.Text = "Move"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 88)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 23)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Conveyor Width"
        '
        'Width
        '
        Me.Width.DecimalPlaces = 1
        Me.Width.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Width.Increment = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.Width.Location = New System.Drawing.Point(168, 88)
        Me.Width.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.Width.Minimum = New Decimal(New Integer() {25, 0, 0, 0})
        Me.Width.Name = "Width"
        Me.Width.Size = New System.Drawing.Size(64, 27)
        Me.Width.TabIndex = 42
        Me.Width.Value = New Decimal(New Integer() {150, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(232, 136)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 24)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "mm"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(232, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 24)
        Me.Label5.TabIndex = 45
        Me.Label5.Text = "mm"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(200, 32)
        Me.Label17.TabIndex = 33
        Me.Label17.Text = "Conveyor Settings"
        '
        'gpbDualHead
        '
        Me.gpbDualHead.Controls.Add(Me.chkDualHead)
        Me.gpbDualHead.Location = New System.Drawing.Point(16, 936)
        Me.gpbDualHead.Name = "gpbDualHead"
        Me.gpbDualHead.Size = New System.Drawing.Size(176, 42)
        Me.gpbDualHead.TabIndex = 36
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
        'PositionTimer
        '
        Me.PositionTimer.Enabled = True
        Me.PositionTimer.SynchronizingObject = Me
        '
        'ConveyorSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 912)
        Me.Controls.Add(Me.gpbDualHead)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "ConveyorSettings"
        Me.Text = "Conveyor"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.downstreamTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.upstreamTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TimeOut, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Speed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.WidthMoveStep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Width, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpbDualHead.ResumeLayout(False)
        CType(Me.PositionTimer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim ButtonConveyorHome_Clicked As Boolean = False
    Dim CurrentWidthValue As Double
    Friend WithEvents T1 As Timer = IDS.T1
    Private Testing As Boolean = True

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        With IDS.Data.Hardware.Conveyor
            .Speed = Speed.Value
            .TimeOut = TimeOut.Value
            .upstreamTimeout = upstreamTimeout.Value
            .downstreamTimeout = downstreamTimeout.Value
            .Width = Conveyor.WidthPosition ' Width.Value
            .WidthMoveStep = WidthMoveStep.Value
        End With
        'upstreamTimeout.Value
        'downstreamTimeout.Value
        Width.Value = Conveyor.WidthPosition
        IDS.Data.SaveData()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        ' do not save to pattern file, retrieve old value from pattern file
        RevertData()
    End Sub

    Public Sub RevertData(Optional ByVal hideexit As Boolean = False)
        ButtonExit.Visible = Not hideexit
        IDS.Data.OpenData()
        With IDS.Data.Hardware.Conveyor
            Speed.Text = .Speed
            TimeOut.Text = .TimeOut
            Width.Text = .Width
            WidthMoveStep.Text = .WidthMoveStep
            Me.upstreamTimeout.Text = .upstreamTimeout
            Me.downstreamTimeout.Text = .downstreamTimeout
        End With
    End Sub

    '   Use "LostFocus" method to update and match the value changed between "SpeedEnter" and "SpeedTestEnter".     '
    Private Sub SpeedEnter_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If Not IsNumeric(Speed.Value) Then
            MsgBox("Please a numeric value!")
            Exit Sub
        End If
    End Sub

    Public Sub CloseConveyorSetup()
        If Conveyor Is Nothing Then Return
        Conveyor.Command("Normal Mode")
        Conveyor.Command("Start Mode")
        Conveyor.Command("Manual Mode")
        Conveyor.Command("Non Production Mode")
        Conveyor.Command("Reset PLC Logic")
    End Sub

    Private Sub ConveyorMove(ByVal width As Double)
        'Dim init_position As Double = Conveyor.WidthPosition
        'start_time = Now
        'Do
        '    'haven't move from initial position, so call again.
        '    If init_position = Conveyor.WidthPosition Then Conveyor.MoveTo(width)
        '    'wait 5s
        '    For i As Integer = 1 To 100
        '        Sleep(50)
        '        TraceDoEvents()
        '    Next
        '    stop_time = Now
        '    elapsed_time = stop_time.Subtract(start_time)
        'Loop Until Conveyor.WidthPosition = width Or elapsed_time.TotalSeconds > 20 Or m_Tri.EStopActivated
    End Sub

    Public Function InitializeConveyorSetup() As Boolean
        If (Conveyor Is Nothing) Then
            Return False
        End If

        Dim speed As Integer
        Dim success As Boolean = True
        'reset so you don't quit the loop too fast.
        Conveyor.ConveyorError = "No Error"

        Conveyor.Command("Width Mode")
        Conveyor.MoveToWidth = IDS.Data.Hardware.Conveyor.Width
        'Console.WriteLine("Conveyor width: " & IDS.Data.Hardware.Conveyor.Width)
        'wait abit to get position value from positiontimer_tick otherwise widthposition = 0
        For i As Integer = 1 To 5
            Sleep(50)
            TraceDoEvents()
        Next i
        success = Conveyor.MoveTo(IDS.Data.Hardware.Conveyor.Width)
        If success Then
            start_time = Now
            Do
                Sleep(50)
                TraceDoEvents()
                stop_time = Now
                elapsed_time = stop_time.Subtract(start_time)
            Loop Until Conveyor.WidthPosition = IDS.Data.Hardware.Conveyor.Width Or m_Tri.EStopActivated Or elapsed_time.TotalSeconds > 10 Or Conveyor.GetError() = "Width Adjustment Failed"
        End If
        If Conveyor.GetError() = "Width Adjustment Failed" Then
            success = False
        End If
        Conveyor.Command("Normal Mode")
        speed = Fix((IDS.Data.Hardware.Conveyor.Speed / (21.5 * 3.1416)) * 100)
        Conveyor.SetCommand("Conveyor Speed", speed)
        Conveyor.SetCommand("Retrieve Timer", IDS.Data.Hardware.Conveyor.TimeOut)
        Conveyor.Command("Production Mode")
        Return success
        'Conveyor.PositionTimer.Stop()
    End Function

    Friend Sub Conveyor_T1_Tick()

        If Form_Service.NoActionToExecute Then
            If Conveyor.GetError() <> "No Error" Then
                Form_Service.DisplayErrorMessage(Conveyor.GetError())
            End If
        End If

        '1003201 - Width Adjustment Failed
        'OK pressed
        If IDS.__ID = "1013201" Then
            If Not (Conveyor Is Nothing) Then
                Conveyor.ActivePositionTimer(True)
            End If
            Conveyor.Command("Width Mode")
            start_time = Now
            Do
                Sleep(50)
                TraceDoEvents()
                stop_time = Now
                elapsed_time = stop_time.Subtract(start_time)
            Loop Until Conveyor.GetError() = "No Error" Or elapsed_time.TotalSeconds > 5 Or m_Tri.EStopActivated
            If Conveyor.GetError = "No Error" Then
                Conveyor.Command("Normal Mode")
                Conveyor.MoveTo(IDS.Data.Hardware.Conveyor.Width)
                'Conveyor.PositionTimer.Stop()
            End If
            Form_Service.ResetEventCode()

            '1003201 - Width Adjustment Failed
            'SKIP pressed
        ElseIf IDS.__ID = "1023201" Then
            Conveyor.Command("Normal Mode")
            Conveyor.Command("Reset Width Error") 'reset width mode error (M90)
            Conveyor.ConveyorError = "No Error"
            'Conveyor.PositionTimer.Stop()
            Form_Service.ResetEventCode()

            '1013210/1013211 - PLC Communication Error or PLC No Signal Conveyor
            'OK pressed
        ElseIf IDS.__ID = "1013210" Or IDS.__ID = "1013211" Then
            Form_Service.ResetEventCode()

            'SKIP pressed
        ElseIf IDS.__ID = "1023202" Or IDS.__ID = "1023203" Or IDS.__ID = "1023204" Or IDS.__ID = "1023206" Or IDS.__ID = "1023208" Then
            '       "Lifter 1/2/3" errors [No air supply (or) Hardware, valve]      '
            '       "Up-Stream to S1 Jam" or "No board is being come"               '
            '       or "S3 to Down-Stream Jam" or "No board is going out"           '
            '   Respond:                                                            
            '    User presses "Skip" button.                                   
            '   (Reset "Error" and "Respond")
            Conveyor.Reset()
            MyConveyorSettings.InitializeConveyorSetup()
            Form_Service.ResetEventCode()

            'RESET pressed
        ElseIf IDS.__ID > 1123200 And IDS.__ID < 1123300 Then
            Conveyor.Reset()

            'Exit Do
            Form_Service.ResetEventCode()

            'RELEASE pressed
        ElseIf IDS.__ID > 1133200 And IDS.__ID < 1133300 Then
            Conveyor.Command("Release")
            Conveyor.Reset()
            'reset Up - S1
            'reset S1 - S2
            'reset S2 - S3
            'reset S3 - Down
            'reset S2
            'reset Lifter Limit Error
            Form_Service.ResetEventCode()
        End If

    End Sub

    Friend Sub Camera_T1_Tick()

        If Form_Service.NoActionToExecute Then
            'If Not IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetter" Then
            If Vision.FrmVision.IsThereCameraError Then Form_Service.DisplayErrorMessage("Camera Signal Failed")
            'End If
        End If

        If IDS.__ID = "1016299" Then
            Form_Service.ResetEventCode()
        End If

    End Sub

    Friend Sub Regulator_T1_Tick()

        'If Form_Service.NoActionToExecute Then
        '    Form_Service.SetEventCode("1001200")
        '    IDS.Devices.Regulator.getRegulatorErr(IDS.__ID)
        'End If

        'If IDS.__ID > 1011210 And IDS.__ID < 1011217 Then
        '    'Exit Do
        '    Form_Service.SetEventCode("1001200")
        'End If

    End Sub

    Dim pressure_counter As Integer = 0 'use this temporarily to record how many times pressure_t1_tick is called before re-showing error pop up if pressureerror_flag = true all the time

    Public Sub Pressure_T1_Tick()
        If Testing Then Return
        If IDS.Devices.DIO.DIO.pressureerror_flag = True And Form_Service.NoActionToExecute Then
            If pressure_counter = 0 Then
                pressure_counter = 100
                Form_Service.DisplayErrorMessage("Low Incoming Air Pressure")
            Else
                pressure_counter -= 1
            End If
        End If

        If IDS.__ID = "1013212" Then
            Form_Service.ResetEventCode()
            IDS.Devices.DIO.DIO.pressureerror_flag = False
        End If

    End Sub
    Public Sub e_stop_T1_Tick()

        If m_Tri.EStopActivated() Then
            If Form_Service.Visible = False And IDS.__ID <> "1006205" Then
                Form_Service.DisplayErrorMessage("E-Stop")
                IDS.Devices.Vision.FrmVision.SwitchCamera("T_Camera")
                IDS.Devices.Vision.FrmVision.SetBrightness(IDS.Data.Hardware.Camera.Brightness)
            End If
        End If

    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        'Conveyor.PositionTimer.Stop()
        RemovePanel(CurrentControl)
    End Sub

    Private Sub ButtonConveyorStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorStart.Click
        Conveyor.MoveTo(Width.Value)
    End Sub


    Private Sub ButtonConveyorHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorHome.Click
        Conveyor.Command("Width Mode")
        Conveyor.Command("Homing")
        Conveyor.Command("Normal Mode")
    End Sub

    Private Sub ButtonConveyorMoveForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorMoveForward.Click
        Conveyor.Command("Width Mode")
        Conveyor.SetCommand("Jog Out", WidthMoveStep.Value * 2)
        Conveyor.Command("Normal Mode")
    End Sub

    Private Sub ButtonConveyorMoveRev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorMoveRev.Click
        Conveyor.Command("Width Mode")
        Conveyor.SetCommand("Jog In", WidthMoveStep.Value * 2)
        Conveyor.Command("Normal Mode")
    End Sub

    Private Overloads Sub MyConveyorSettings_Tick(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles PositionTimer.Elapsed
        If Not (Conveyor Is Nothing) Then
            WidthDisplay.Text = CStr(Conveyor.WidthPosition)
        End If

    End Sub
End Class
