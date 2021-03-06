Imports System.IO
Public Class Conveyor
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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorMoveForward As System.Windows.Forms.Button
    Friend WithEvents ConveyorFineAdjust As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonConveyorMoveReverse As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TimeOutEnter As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents SpeedEnter As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButtonConveyorStart As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ConveyorWidthEnter As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents ConveyorWidth As System.Windows.Forms.Label
    Friend WithEvents BoxToBeAdded As System.Windows.Forms.GroupBox
    Public WithEvents AxMSComm1 As AxMSCommLib.AxMSComm
    Friend WithEvents BT_Retrieve As System.Windows.Forms.Button
    Friend WithEvents BT_Release As System.Windows.Forms.Button
    Friend WithEvents ResetPLCLogic As System.Windows.Forms.Button
    Public WithEvents PositionTimer As System.Windows.Forms.Timer
    Friend WithEvents ButtonHome As System.Windows.Forms.Button
    Friend WithEvents ButtonApply As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents upstreamTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents downstreamTimeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Conveyor))
        Me.AxMSComm1 = New AxMSCommLib.AxMSComm
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.ButtonConveyorMoveReverse = New System.Windows.Forms.Button
        Me.ButtonConveyorMoveForward = New System.Windows.Forms.Button
        Me.ConveyorFineAdjust = New System.Windows.Forms.NumericUpDown
        Me.PositionTimer = New System.Windows.Forms.Timer(Me.components)
        Me.BoxToBeAdded = New System.Windows.Forms.GroupBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.downstreamTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.upstreamTimeout = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.TimeOutEnter = New System.Windows.Forms.NumericUpDown
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.SpeedEnter = New System.Windows.Forms.NumericUpDown
        Me.Label10 = New System.Windows.Forms.Label
        Me.ButtonApply = New System.Windows.Forms.Button
        Me.ConveyorWidth = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonConveyorStart = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.ConveyorWidthEnter = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.BT_Retrieve = New System.Windows.Forms.Button
        Me.BT_Release = New System.Windows.Forms.Button
        Me.ResetPLCLogic = New System.Windows.Forms.Button
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ConveyorFineAdjust, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.BoxToBeAdded.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.downstreamTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.upstreamTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TimeOutEnter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SpeedEnter, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ConveyorWidthEnter, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AxMSComm1
        '
        Me.AxMSComm1.Enabled = True
        Me.AxMSComm1.Location = New System.Drawing.Point(488, 392)
        Me.AxMSComm1.Name = "AxMSComm1"
        Me.AxMSComm1.OcxState = CType(resources.GetObject("AxMSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMSComm1.Size = New System.Drawing.Size(38, 38)
        Me.AxMSComm1.TabIndex = 0
        '
        'ButtonHome
        '
        Me.ButtonHome.Location = New System.Drawing.Point(392, 80)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(72, 40)
        Me.ButtonHome.TabIndex = 1
        Me.ButtonHome.Text = "Home"
        '
        'ButtonConveyorMoveReverse
        '
        Me.ButtonConveyorMoveReverse.Location = New System.Drawing.Point(392, 128)
        Me.ButtonConveyorMoveReverse.Name = "ButtonConveyorMoveReverse"
        Me.ButtonConveyorMoveReverse.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorMoveReverse.TabIndex = 25
        Me.ButtonConveyorMoveReverse.Text = ">>--<<"
        '
        'ButtonConveyorMoveForward
        '
        Me.ButtonConveyorMoveForward.Location = New System.Drawing.Point(288, 128)
        Me.ButtonConveyorMoveForward.Name = "ButtonConveyorMoveForward"
        Me.ButtonConveyorMoveForward.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorMoveForward.TabIndex = 24
        Me.ButtonConveyorMoveForward.Text = "<<-->>"
        '
        'ConveyorFineAdjust
        '
        Me.ConveyorFineAdjust.DecimalPlaces = 1
        Me.ConveyorFineAdjust.Increment = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.ConveyorFineAdjust.Location = New System.Drawing.Point(168, 136)
        Me.ConveyorFineAdjust.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.ConveyorFineAdjust.Minimum = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.ConveyorFineAdjust.Name = "ConveyorFineAdjust"
        Me.ConveyorFineAdjust.Size = New System.Drawing.Size(64, 27)
        Me.ConveyorFineAdjust.TabIndex = 47
        Me.ConveyorFineAdjust.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'PositionTimer
        '
        Me.PositionTimer.Interval = 800
        '
        'BoxToBeAdded
        '
        Me.BoxToBeAdded.Controls.Add(Me.GroupBox1)
        Me.BoxToBeAdded.Controls.Add(Me.ConveyorWidth)
        Me.BoxToBeAdded.Controls.Add(Me.Label2)
        Me.BoxToBeAdded.Controls.Add(Me.Label5)
        Me.BoxToBeAdded.Controls.Add(Me.ButtonConveyorStart)
        Me.BoxToBeAdded.Controls.Add(Me.Label7)
        Me.BoxToBeAdded.Controls.Add(Me.ConveyorWidthEnter)
        Me.BoxToBeAdded.Controls.Add(Me.Label8)
        Me.BoxToBeAdded.Controls.Add(Me.Label12)
        Me.BoxToBeAdded.Controls.Add(Me.ButtonConveyorMoveReverse)
        Me.BoxToBeAdded.Controls.Add(Me.ButtonConveyorMoveForward)
        Me.BoxToBeAdded.Controls.Add(Me.ButtonHome)
        Me.BoxToBeAdded.Controls.Add(Me.ConveyorFineAdjust)
        Me.BoxToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.BoxToBeAdded.Location = New System.Drawing.Point(16, 16)
        Me.BoxToBeAdded.Name = "BoxToBeAdded"
        Me.BoxToBeAdded.Size = New System.Drawing.Size(496, 360)
        Me.BoxToBeAdded.TabIndex = 49
        Me.BoxToBeAdded.TabStop = False
        Me.BoxToBeAdded.Text = "Conveyor Information"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.downstreamTimeout)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.upstreamTimeout)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.TimeOutEnter)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.SpeedEnter)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.ButtonApply)
        Me.GroupBox1.Location = New System.Drawing.Point(16, 168)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(464, 168)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Conveyor Settings"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(264, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 23)
        Me.Label4.TabIndex = 53
        Me.Label4.Text = "s"
        '
        'downstreamTimeout
        '
        Me.downstreamTimeout.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.downstreamTimeout.Location = New System.Drawing.Point(192, 128)
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
        Me.Label13.Location = New System.Drawing.Point(8, 128)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(176, 23)
        Me.Label13.TabIndex = 54
        Me.Label13.Text = "downstream Timeout"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(264, 96)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(24, 23)
        Me.Label1.TabIndex = 50
        Me.Label1.Text = "s"
        '
        'upstreamTimeout
        '
        Me.upstreamTimeout.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.upstreamTimeout.Location = New System.Drawing.Point(192, 96)
        Me.upstreamTimeout.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.upstreamTimeout.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.upstreamTimeout.Name = "upstreamTimeout"
        Me.upstreamTimeout.Size = New System.Drawing.Size(64, 27)
        Me.upstreamTimeout.TabIndex = 49
        Me.upstreamTimeout.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(168, 23)
        Me.Label3.TabIndex = 51
        Me.Label3.Text = "Upstream Timeout"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(264, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(24, 23)
        Me.Label6.TabIndex = 47
        Me.Label6.Text = "s"
        '
        'TimeOutEnter
        '
        Me.TimeOutEnter.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TimeOutEnter.Location = New System.Drawing.Point(192, 72)
        Me.TimeOutEnter.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.TimeOutEnter.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.TimeOutEnter.Name = "TimeOutEnter"
        Me.TimeOutEnter.Size = New System.Drawing.Size(64, 27)
        Me.TimeOutEnter.TabIndex = 46
        Me.TimeOutEnter.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(40, 72)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(144, 23)
        Me.Label9.TabIndex = 48
        Me.Label9.Text = "Retrieve Timeout"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(264, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 24)
        Me.Label11.TabIndex = 45
        Me.Label11.Text = "mm/s"
        '
        'SpeedEnter
        '
        Me.SpeedEnter.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.SpeedEnter.Location = New System.Drawing.Point(192, 40)
        Me.SpeedEnter.Maximum = New Decimal(New Integer() {400, 0, 0, 0})
        Me.SpeedEnter.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.SpeedEnter.Name = "SpeedEnter"
        Me.SpeedEnter.Size = New System.Drawing.Size(64, 27)
        Me.SpeedEnter.TabIndex = 44
        Me.SpeedEnter.Value = New Decimal(New Integer() {400, 0, 0, 0})
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(120, 40)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 23)
        Me.Label10.TabIndex = 43
        Me.Label10.Text = "Speed"
        '
        'ButtonApply
        '
        Me.ButtonApply.Location = New System.Drawing.Point(336, 64)
        Me.ButtonApply.Name = "ButtonApply"
        Me.ButtonApply.Size = New System.Drawing.Size(112, 56)
        Me.ButtonApply.TabIndex = 19
        Me.ButtonApply.Text = "Download to Conveyor"
        '
        'ConveyorWidth
        '
        Me.ConveyorWidth.BackColor = System.Drawing.Color.Transparent
        Me.ConveyorWidth.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ConveyorWidth.Location = New System.Drawing.Point(208, 40)
        Me.ConveyorWidth.Name = "ConveyorWidth"
        Me.ConveyorWidth.Size = New System.Drawing.Size(272, 23)
        Me.ConveyorWidth.TabIndex = 18
        Me.ConveyorWidth.Text = "Front Sensor Hit. Please Wait..."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(128, 23)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Current Width:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Fine Adjust"
        '
        'ButtonConveyorStart
        '
        Me.ButtonConveyorStart.Location = New System.Drawing.Point(288, 80)
        Me.ButtonConveyorStart.Name = "ButtonConveyorStart"
        Me.ButtonConveyorStart.Size = New System.Drawing.Size(75, 40)
        Me.ButtonConveyorStart.TabIndex = 19
        Me.ButtonConveyorStart.Text = "Move"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(136, 23)
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Conveyor Width"
        '
        'ConveyorWidthEnter
        '
        Me.ConveyorWidthEnter.DecimalPlaces = 1
        Me.ConveyorWidthEnter.Increment = New Decimal(New Integer() {5, 0, 0, 65536})
        Me.ConveyorWidthEnter.Location = New System.Drawing.Point(168, 88)
        Me.ConveyorWidthEnter.Maximum = New Decimal(New Integer() {300, 0, 0, 0})
        Me.ConveyorWidthEnter.Minimum = New Decimal(New Integer() {25, 0, 0, 0})
        Me.ConveyorWidthEnter.Name = "ConveyorWidthEnter"
        Me.ConveyorWidthEnter.Size = New System.Drawing.Size(64, 27)
        Me.ConveyorWidthEnter.TabIndex = 42
        Me.ConveyorWidthEnter.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(232, 136)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(40, 24)
        Me.Label8.TabIndex = 45
        Me.Label8.Text = "mm"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label12.Location = New System.Drawing.Point(232, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(40, 24)
        Me.Label12.TabIndex = 45
        Me.Label12.Text = "mm"
        '
        'BT_Retrieve
        '
        Me.BT_Retrieve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Retrieve.Location = New System.Drawing.Point(52, 392)
        Me.BT_Retrieve.Name = "BT_Retrieve"
        Me.BT_Retrieve.Size = New System.Drawing.Size(88, 48)
        Me.BT_Retrieve.TabIndex = 50
        Me.BT_Retrieve.Text = "Retrieve"
        '
        'BT_Release
        '
        Me.BT_Release.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Release.Location = New System.Drawing.Point(164, 392)
        Me.BT_Release.Name = "BT_Release"
        Me.BT_Release.Size = New System.Drawing.Size(88, 48)
        Me.BT_Release.TabIndex = 51
        Me.BT_Release.Text = "Release"
        '
        'ResetPLCLogic
        '
        Me.ResetPLCLogic.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResetPLCLogic.Location = New System.Drawing.Point(276, 392)
        Me.ResetPLCLogic.Name = "ResetPLCLogic"
        Me.ResetPLCLogic.Size = New System.Drawing.Size(200, 48)
        Me.ResetPLCLogic.TabIndex = 54
        Me.ResetPLCLogic.Text = "Reset PLC Logic"
        '
        'Conveyor
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(528, 456)
        Me.Controls.Add(Me.BT_Retrieve)
        Me.Controls.Add(Me.BT_Release)
        Me.Controls.Add(Me.ResetPLCLogic)
        Me.Controls.Add(Me.BoxToBeAdded)
        Me.Controls.Add(Me.AxMSComm1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Conveyor"
        Me.Text = "Conveyor"
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ConveyorFineAdjust, System.ComponentModel.ISupportInitialize).EndInit()
        Me.BoxToBeAdded.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.downstreamTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.upstreamTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TimeOutEnter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SpeedEnter, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ConveyorWidthEnter, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'exposed to other modules
    Public MinWidth As Double = 25
    Public WidthPosition As Double = 0
    Public ConveyorError As String = ""
    Public MoveToWidth As Double = -1
  
    Public Function GetError() As String
        If ConveyorError = "Width" Then
            Return "Width Adjustment Failed"
        ElseIf ConveyorError = "Lifter 1 Blocked" Then
            Return "Lifter 1 Blocked"
        ElseIf ConveyorError = "Lifter 2 Blocked" Then
            Return "Lifter 2 Blocked"
        ElseIf ConveyorError = "Lifter 3 Blocked" Then
            Return "Lifter 3 Blocked"
        ElseIf ConveyorError = "S1 to S2 Jam" Then
            Return "Retrieve Timeout"
        ElseIf ConveyorError = "Up-Stream to S1 Jam" Then
            Return "Stage 1 Travel Timeout"
        ElseIf ConveyorError = "S2 to S3 Jam" Then
            Return "Release Travel Timeout"
        ElseIf ConveyorError = "S3 to Down-Stream Jam" Then
            Return "Stage 3 Travel Timeout"
        ElseIf ConveyorError = "Alignment Wrong" Then
            Return "Board Not Aligned"
        ElseIf ConveyorError = "Others" Then
            Return "PLC Communication Error"
        Else
            'If ConveyorError = "No Error" Then
            Return "No Error"
        End If
    End Function

    Public Function ReportWidthAdjustFailure() As String
        Return "Tried to move to " + MoveToWidth.ToString + " mm from " + WidthPosition.ToString + " mm."
    End Function

    Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm
        Dim raw_result As Object
        Dim result As Integer
        Dim byte_result As Byte()

        Try
            With AxMSComm1
                Select Case .CommEvent
                    Case 2
                        .RThreshold = 0
                        raw_result = AxMSComm1.Input
                        byte_result = CType(raw_result, Byte())
                    Case Else
                End Select
            End With
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try

        'conveyor at home position is 300mm
        Dim displacement_from_home As Double
        displacement_from_home = (Convert.ToInt32(byte_result(1)) * 256 + Convert.ToInt32(byte_result(2))) / 144
        WidthPosition = 300 - displacement_from_home
        ConveyorWidth.Text = CStr(WidthPosition)

        If byte_result(0) = &H4 And byte_result(3) = &H3 Then
            If MoveToWidth = WidthPosition Then
                ConveyorError = "No Error"
                Command("Reset Width Error")
            Else
                If Not (MoveToWidth = -1) Then
                    ConveyorError = "Width"
                    'Console.WriteLine(DateTime.Now.ToLongTimeString() + " Width error received.")
                Else
                    'Console.WriteLine(DateTime.Now.ToLongTimeString() + " Width error received before any moveto command was sent")
                End If

            End If
        ElseIf byte_result(0) = &H77 Or byte_result(0) = &H2 And byte_result(3) = &H3 Then
            ConveyorError = "No Error"
        ElseIf byte_result(0) = &H51 And byte_result(3) = &H3 Then
            ConveyorError = "Lifter 1 Blocked"
        ElseIf byte_result(0) = &H52 And byte_result(3) = &H3 Then
            ConveyorError = "Lifter 2 Blocked"
        ElseIf byte_result(0) = &H53 And byte_result(3) = &H3 Then
            ConveyorError = "Lifter 3 Blocked"
        ElseIf byte_result(0) = &H41 And byte_result(3) = &H3 Then
            ConveyorError = "S1 to S2 Jam"
        ElseIf byte_result(0) = &H42 And byte_result(3) = &H3 Then
            ConveyorError = "Up-Stream to S1 Jam"
        ElseIf byte_result(0) = &H45 And byte_result(3) = &H3 Then
            ConveyorError = "S2 to S3 Jam"
        ElseIf byte_result(0) = &H46 And byte_result(3) = &H3 Then
            ConveyorError = "S3 to Down-Stream Jam"
        ElseIf byte_result(0) = &H43 And byte_result(3) = &H3 Then
            ConveyorError = "Alignment Wrong"
        Else
            ConveyorError = "Others"
        End If

        AxMSComm1.RThreshold = 1

    End Sub

    Public Sub OpenPort()
        If AxMSComm1.PortOpen = True Then
            Exit Sub
        End If
        Try
            With AxMSComm1
                .CommPort = ConveyorPort
                .Settings = "9600,N,8,1"
                .InBufferSize = 4
                .InputLen = 4
                .RThreshold = 1
                .InputMode = MSCommLib.InputModeConstants.comInputModeBinary
                .InBufferCount = 0
                If Not .PortOpen Then
                    .PortOpen = True
                End If
            End With
        Catch ex As System.Runtime.InteropServices.COMException
            ExceptionDisplay(ex)
        End Try
    End Sub

    Public Sub ClosePort()
        If AxMSComm1.PortOpen Then
            AxMSComm1.PortOpen = False
        End If
    End Sub


    Public Sub Command(ByVal command As String)
        Dim output_str As String = Chr(&H2)
        'If AxMSComm1.InBufferCount < 4 Then
        '    Exit Sub
        'End If
        Select Case command
            Case "Read Position"
                output_str = output_str + Chr(&H13)   'read position
            Case "Homing"
                output_str = output_str + Chr(&H10)   'homing
            Case "Reset Up-S1"
                output_str = output_str + Chr(&H42)   'reset Up - S1
            Case "Reset S1-S2"
                output_str = output_str + Chr(&H41)   'reset S1 - S2
            Case "Reset S2-S3"
                output_str = output_str + Chr(&H45)   'reset S2 - S3
            Case "Reset S3-Down"
                output_str = output_str + Chr(&H46)   'reset S3 - Down
            Case "Reset S2"
                output_str = output_str + Chr(&H43)   'reset S2
            Case "Reset Lifter Limit Error"
                output_str = output_str + Chr(&H50)   'reset Lifter Limit Error
            Case "Reset Width Error"
                output_str = output_str + Chr(&H55)   'reset width mode error (M90)
            Case "Reset PLC Logic"
                output_str = output_str + Chr(&HFF)   'reset conveyor status
            Case "Width Mode"
                output_str = output_str + Chr(&H9)   'width mode
            Case "Normal Mode"
                output_str = output_str + Chr(&H8)   'normal mode
            Case "Start Mode"
                output_str = output_str + Chr(&H3)   'Start
            Case "Stop Mode"
                output_str = output_str + Chr(&H4)   'Stop
            Case "Auto Mode"
                output_str = output_str + Chr(&H2)   'auto mode
            Case "Manual Mode"
                output_str = output_str + Chr(&H1)   'manual mode
            Case "Production Mode"
                output_str = output_str + Chr(&HC)   'Production running
            Case "Non Production Mode"
                output_str = output_str + Chr(&HD)   'Non Production running
            Case "Retrieve"
                output_str = output_str + Chr(&HA)   'retrieve
            Case "Release"
                output_str = output_str + Chr(&HB)   'release
            Case "Left Lifter"
                output_str = output_str + Chr(&H1A)   'Left Lifter
            Case "Center Lifter"
                output_str = output_str + Chr(&H1B)   'Center Lifter
            Case "Right Lifter"
                output_str = output_str + Chr(&H1C)   'Right Lifter
            Case "Reset Conveyor Status"
                output_str = output_str + Chr(&H3F)    'Reset Conveyor Status
        End Select

        output_str = output_str + Chr(&H3)
        'check if we are in a position where we are reading conveyor values.
        'if yes, then reenable after temporarily stopping read
        'if no, then dont enable at all.
        Dim currently_reading As Boolean = PositionTimer.Enabled
        If currently_reading Then
            PositionTimer.Stop()
        End If

        If AxMSComm1.PortOpen Then AxMSComm1.Output = output_str
        Sleep(50)
        If currently_reading Then
            PositionTimer.Start()
        End If


    End Sub

    Public Sub Reset()
        Command("Reset Up-S1")
        Command("Reset S1-S2")
        Command("Reset S2-S3")
        Command("Reset S3-Down")
        Command("Reset S2")
        Command("Reset Lifter Limit Error")
        ConveyorError = "No Error"
    End Sub

    Public Sub SetCommand(ByVal command As String, ByVal value As Decimal)

        Dim output_str As String = Chr(&H2)
        Dim command_temp As String

        'user should never be able to see this
        If value < 0 And value > 999 Then
            MsgBox("Step movement out of bounds")
            Exit Sub
        End If

        If value < 999 And value > 99 Then
            command_temp = CStr(value)
        ElseIf value < 100 And value > 9 Then
            command_temp = "0" & CStr(value)
        ElseIf value < 10 And value > 0 Then
            command_temp = "00" & CStr(value)
        Else
            command_temp = "Invalid value of " + value.ToString + " for setting conveyor command"
            MsgBox(command_temp)
            Exit Sub
        End If

        Select Case command
            Case "Jog In"
                output_str = output_str + Chr(&H11)
            Case "Jog Out"
                output_str = output_str + Chr(&H12)
            Case "Conveyor Speed"
                output_str = output_str + Chr(&H7)
            Case "Left Timer"
                output_str = output_str + Chr(&H6)
            Case "Right Timer"
                output_str = output_str + Chr(&H60)
            Case "Retrieve Timer"
                output_str = output_str + Chr(&H5)
        End Select

        output_str = output_str & Chr(Val(command_temp.Chars(0))) & Chr(Val(command_temp.Chars(1))) & Chr(Val(command_temp.Chars(2)))

        output_str = output_str + Chr(&H3)
        Dim PositionTimerStatus As Boolean = PositionTimer.Enabled
        If PositionTimerStatus Then
            PositionTimer.Stop()
        End If

        If AxMSComm1.PortOpen Then AxMSComm1.Output = output_str
        Sleep(50)
        If PositionTimerStatus Then
            PositionTimer.Start()
        End If


    End Sub

    Private Sub Conveyor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OpenPort()
    End Sub


    Public Function MoveTo(ByVal width As Double) As Boolean
        'Console.WriteLine("Conveyor command move to " & width)
        MoveToWidth = width
        Dim DesiredAbsPost As Double = 0.0
        Dim distance As Decimal
        If MoveToWidth < MinWidth Then 'Cannot move smaller than MinWidth otherwise the sensor will be block and cannot move the conveyor again, cos PLC assume products is on the conveyor.
            Return False
        End If
        If WidthPosition > width Then
            distance = CDec(WidthPosition - width)
            Command("Width Mode")
            SetCommand("Jog In", distance * 2)
            Command("Normal Mode")
        ElseIf WidthPosition < width Then
            distance = CDec(width - WidthPosition)
            Command("Width Mode")
            SetCommand("Jog Out", distance * 2)
            Command("Normal Mode")
        End If
        Return True
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHome.Click
        Command("Width Mode")
        Command("Homing")
        Command("Normal Mode")
    End Sub

    Private Sub ButtonConveyorMoveForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorMoveForward.Click
        Command("Width Mode")
        SetCommand("Jog Out", ConveyorFineAdjust.Value * 2)
        Command("Normal Mode")
    End Sub

    Private Sub ButtonConveyorMoveReverse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorMoveReverse.Click
        Command("Width Mode")
        SetCommand("Jog In", ConveyorFineAdjust.Value * 2)
        Command("Normal Mode")
    End Sub

    Private Sub PositionTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PositionTimer.Tick
        Command("Read Position")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonApply.Click
        SetCommand("Retrieve Timer", TimeOutEnter.Value)
        SetCommand("Left Timer", upstreamTimeout.Value)
        SetCommand("Right Timer", downstreamTimeout.Value)
        SetCommand("speed", SpeedEnter.Value)
    End Sub

    Private Sub ButtonConveyorStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorStart.Click
        MoveTo(ConveyorWidthEnter.Value)
    End Sub


    Private Sub ResetPLCLogic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPLCLogic.Click
        Command("Reset Conveyor Status")
    End Sub


    Private Sub BT_Retrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Retrieve.Click
        Command("Retrieve")
    End Sub

    Private Sub BT_Release_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Release.Click
        Command("Release")
    End Sub

    Private Sub Conveyor_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        e.Cancel = True
        Hide()
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

    End Sub
End Class
