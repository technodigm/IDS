Imports DLL_Export_Service
Imports DLL_Export_Device_Motor
Imports System.Threading

Public Class LaserCalibration
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ButtonLSStart As System.Windows.Forms.Button
    Friend WithEvents NewLaserZReference As System.Windows.Forms.Label
    Friend WithEvents NewLaserYOffset As System.Windows.Forms.Label
    Friend WithEvents NewLaserXOffset As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetX As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetY As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents LaserReading As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents LasertoCameraOffset As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetZ As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(LaserCalibration))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.NewLaserZReference = New System.Windows.Forms.Label
        Me.NewLaserYOffset = New System.Windows.Forms.Label
        Me.NewLaserXOffset = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.LasertoCameraOffset = New System.Windows.Forms.Label
        Me.ButtonLSStart = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.LaserReading = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label11 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.LaserOffsetX = New System.Windows.Forms.Label
        Me.LaserOffsetY = New System.Windows.Forms.Label
        Me.LaserOffsetZ = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.ButtonRevert)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ButtonSave)
        Me.GroupBox1.Controls.Add(Me.LaserReading)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.LaserOffsetY)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.LaserOffsetZ)
        Me.GroupBox1.Controls.Add(Me.LaserOffsetX)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(472, 792)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Position Offset of Laser Sensing Point"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.NewLaserZReference)
        Me.GroupBox2.Controls.Add(Me.NewLaserYOffset)
        Me.GroupBox2.Controls.Add(Me.NewLaserXOffset)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.LasertoCameraOffset)
        Me.GroupBox2.Controls.Add(Me.ButtonLSStart)
        Me.GroupBox2.Location = New System.Drawing.Point(40, 280)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(392, 376)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Locate Offset"
        '
        'Label4
        '
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(64, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(176, 25)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "New Laser X Offset:"
        '
        'NewLaserZReference
        '
        Me.NewLaserZReference.Location = New System.Drawing.Point(264, 240)
        Me.NewLaserZReference.Name = "NewLaserZReference"
        Me.NewLaserZReference.Size = New System.Drawing.Size(88, 24)
        Me.NewLaserZReference.TabIndex = 18
        Me.NewLaserZReference.Text = "0.000"
        '
        'NewLaserYOffset
        '
        Me.NewLaserYOffset.Location = New System.Drawing.Point(264, 192)
        Me.NewLaserYOffset.Name = "NewLaserYOffset"
        Me.NewLaserYOffset.Size = New System.Drawing.Size(88, 24)
        Me.NewLaserYOffset.TabIndex = 17
        Me.NewLaserYOffset.Text = "0.000"
        '
        'NewLaserXOffset
        '
        Me.NewLaserXOffset.Location = New System.Drawing.Point(264, 144)
        Me.NewLaserXOffset.Name = "NewLaserXOffset"
        Me.NewLaserXOffset.Size = New System.Drawing.Size(88, 24)
        Me.NewLaserXOffset.TabIndex = 16
        Me.NewLaserXOffset.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(16, 240)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(232, 26)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "New Laser Reference Point:"
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(64, 192)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(168, 30)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "New Laser Y Offset:"
        '
        'LasertoCameraOffset
        '
        Me.LasertoCameraOffset.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.LasertoCameraOffset.Location = New System.Drawing.Point(8, 48)
        Me.LasertoCameraOffset.Name = "LasertoCameraOffset"
        Me.LasertoCameraOffset.Size = New System.Drawing.Size(376, 32)
        Me.LasertoCameraOffset.TabIndex = 13
        Me.LasertoCameraOffset.Text = "Laser to camera offset"
        Me.LasertoCameraOffset.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ButtonLSStart
        '
        Me.ButtonLSStart.Location = New System.Drawing.Point(144, 304)
        Me.ButtonLSStart.Name = "ButtonLSStart"
        Me.ButtonLSStart.Size = New System.Drawing.Size(100, 50)
        Me.ButtonLSStart.TabIndex = 13
        Me.ButtonLSStart.Text = "Start"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Location = New System.Drawing.Point(384, 680)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(75, 40)
        Me.ButtonRevert.TabIndex = 12
        Me.ButtonRevert.Text = "Revert"
        Me.ButtonRevert.Visible = False
        '
        'Label2
        '
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label2.Location = New System.Drawing.Point(32, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(222, 30)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "2. Click Start Button"
        '
        'Label3
        '
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(32, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(432, 48)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "1. Jog laser sensing point to the center of Height Reference Block."
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(296, 680)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 40)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        Me.ButtonSave.Visible = False
        '
        'LaserReading
        '
        Me.LaserReading.Location = New System.Drawing.Point(296, 241)
        Me.LaserReading.Name = "LaserReading"
        Me.LaserReading.Size = New System.Drawing.Size(88, 24)
        Me.LaserReading.TabIndex = 16
        '
        'Label9
        '
        Me.Label9.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label9.Location = New System.Drawing.Point(32, 240)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(216, 25)
        Me.Label9.TabIndex = 12
        Me.Label9.Text = "Current Laser Reading  :"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(-102, -8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 21)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Sensor"
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 892)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(704, 20)
        Me.StatusBar1.TabIndex = 33
        Me.StatusBar1.Text = "Ready"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label11)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Panel2)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(112, 8)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 31
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label11.Location = New System.Drawing.Point(0, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(280, 32)
        Me.Label11.TabIndex = 39
        Me.Label11.Text = "Laser Sensor Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 38
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(12, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(488, 848)
        Me.Panel2.TabIndex = 34
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(224, 23)
        Me.Label5.TabIndex = 81
        Me.Label5.Text = "Existing Laser Reference :"
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(16, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(184, 23)
        Me.Label16.TabIndex = 35
        Me.Label16.Text = "Existing Laser Offset :"
        '
        'LaserOffsetX
        '
        Me.LaserOffsetX.Location = New System.Drawing.Point(352, 32)
        Me.LaserOffsetX.Name = "LaserOffsetX"
        Me.LaserOffsetX.Size = New System.Drawing.Size(96, 23)
        Me.LaserOffsetX.TabIndex = 74
        Me.LaserOffsetX.Text = "100.000"
        '
        'LaserOffsetY
        '
        Me.LaserOffsetY.Location = New System.Drawing.Point(224, 32)
        Me.LaserOffsetY.Name = "LaserOffsetY"
        Me.LaserOffsetY.Size = New System.Drawing.Size(120, 23)
        Me.LaserOffsetY.TabIndex = 77
        Me.LaserOffsetY.Text = "100.000"
        '
        'LaserOffsetZ
        '
        Me.LaserOffsetZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LaserOffsetZ.Location = New System.Drawing.Point(312, 64)
        Me.LaserOffsetZ.Name = "LaserOffsetZ"
        Me.LaserOffsetZ.Size = New System.Drawing.Size(72, 23)
        Me.LaserOffsetZ.TabIndex = 79
        Me.LaserOffsetZ.Text = "100.000"
        '
        'Timer1
        '
        Me.Timer1.Interval = 10
        '
        'LaserCalibration
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.ClientSize = New System.Drawing.Size(704, 912)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "LaserCalibration"
        Me.Text = "Laser Sensor"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private Sub ButtonLSSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        With IDS.Data.Hardware.HeightSensor.Laser
            .HeightReference = NewLaserZReference.Text
            .OffsetPos.X = NewLaserXOffset.Text - IDS.Data.Hardware.Camera.ReferencePos.X
            .OffsetPos.Y = NewLaserYOffset.Text - IDS.Data.Hardware.Camera.ReferencePos.Y
            'Manual calculation override (why??? kr sept 2015)
            '.OffsetPos.X = 52.819 - IDS.Data.Hardware.Camera.ReferencePos.X
            '.OffsetPos.Y = 127.271 - IDS.Data.Hardware.Camera.ReferencePos.Y
        End With

        IDS.Data.SaveData()
    End Sub
  

    Private Sub ButtonLSStart_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLSStart.Click

        If ButtonLSStart.Text = "Start" Then
            ButtonLSStart.Text = "Stop"
            StartLaserBlockCalibration()
        Else
            ExitLaserCalibration()
            ButtonLSStart.Text = "Start"
        End If
        Return
        Timer1.Start()
        m_Tri.SetMachineRun()
        Laser.EnableContinuousRead()

        Dim counter As Integer = 0
        Dim calibrator_plate_height, block_height, difference As Double
        Dim position(3) As Double

        Laser.LASER_Reading = 0
        WaitLoop()
        block_height = Laser.LASER_Reading
        NewLaserZReference.Text = block_height
        IDS.Data.Hardware.HeightSensor.Laser.HeightReference = block_height
        IDS.Data.Hardware.HeightSensor.Laser.CurrentPos.Z = block_height
        IDS.Data.SaveData()

        m_Tri.Set_XY_Speed(10)

        '1) move to the top right position outside block
        distance(0) = 7
        distance(1) = 7
        m_Tri.MoveRelative_XY(distance)

        '2) move to the center right position outside block
        'MySleep(100)
        difference = Math.Abs(Laser.LASER_Reading - block_height)
        If difference > 0.95 And difference < 2 Then
            distance(0) = 0
            distance(1) = -7
            m_Tri.MoveRelative_XY(distance)
            WaitLoop()
            calibrator_plate_height = Laser.LASER_Reading
        Else
            GoTo CalibrationError
        End If

        '3) move from center right position outside block to
        '   the right side of the block edge
        While Math.Abs(calibrator_plate_height - Laser.LASER_Reading) < 0.95 And counter <= 400
            distance(0) = -0.01
            distance(1) = 0
            counter = counter + 1
            LaserReading.Text = Laser.LASER_Reading
            m_Tri.MoveRelative_XY(distance)
        End While

        '4) move to the center x position inside block
        'MySleep(100)
        difference = Math.Abs(Laser.LASER_Reading - calibrator_plate_height)
        If difference > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            position(0) = m_Tri.XPosition
            distance(0) = -9
            distance(1) = 0
            m_Tri.MoveRelative_XY(distance)
            WaitLoop()
            block_height = Laser.LASER_Reading
        Else
            GoTo CalibrationError
        End If

        '5) move from center x position inside block to
        '   left side of the block edge
        counter = 0
        While Math.Abs(block_height - Laser.LASER_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            distance(0) = -0.01
            distance(1) = 0
            LaserReading.Text = Laser.LASER_Reading
            m_Tri.MoveRelative_XY(distance)
        End While

        '6) move from left side of the block edge 
        '   to top center position outside of block
        'MySleep(100)
        difference = Math.Abs(Laser.LASER_Reading - block_height)
        If difference > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            position(1) = m_Tri.XPosition
            distance(0) = 7
            distance(1) = 7
            m_Tri.MoveRelative_XY(distance)
            WaitLoop()
            calibrator_plate_height = Laser.LASER_Reading
        Else
            GoTo CalibrationError
        End If

        '7) move from top center position outside of block
        '   to top edge of the block
        While Math.Abs(calibrator_plate_height - Laser.LASER_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            distance(0) = 0
            distance(1) = -0.01
            LaserReading.Text = Laser.LASER_Reading
            m_Tri.MoveRelative_XY(distance)
        End While

        '8) move from top edge of the block
        '   to the center y position of the block
        'MySleep(100)
        difference = Math.Abs(Laser.LASER_Reading - calibrator_plate_height)
        If difference > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            position(2) = m_Tri.YPosition
            distance(0) = 0
            distance(1) = -9
            m_Tri.MoveRelative_XY(distance)
            WaitLoop()
            block_height = Laser.LASER_Reading
        Else
            GoTo CalibrationError
        End If

        '9) move from the center y position of the block
        '   to the bottom edge of the block
        counter = 0
        While Math.Abs(block_height - Laser.LASER_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            distance(0) = 0
            distance(1) = -0.01
            LaserReading.Text = Laser.LASER_Reading
            m_Tri.MoveRelative_XY(distance)
        End While

        '10) get the y position of the bottom edge of the block
        difference = Math.Abs(Laser.LASER_Reading - block_height)
        If difference > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            position(3) = m_Tri.YPosition
        Else
            GoTo CalibrationError
        End If

        'finish
        TraceDoEvents()
        distance(0) = (position(0) - position(1)) / 2 + position(1)
        distance(1) = (position(2) - position(3)) / 2 + position(3)
        distance(2) = 0
        m_Tri.Set_XY_Speed(10)
        m_Tri.Move_XY(distance)
        MessageBox.Show("finish")
        NewLaserXOffset.Text = distance(0)
        NewLaserYOffset.Text = distance(1)
        IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X = distance(0)
        IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y = distance(1)
        IDS.Data.SaveData()
        'after successful setup disable the following
        Laser.DisableContinuousRead()
        LaserReading.Text = ""
        m_Tri.ResetCalibrationFlag()
        m_Tri.SetMachineStop()
        Exit Sub

CalibrationError:
        MessageBox.Show("Error")
        LaserReading.Text = ""
        Laser.DisableContinuousRead()
        m_Tri.ResetCalibrationFlag()
        m_Tri.SetMachineStop()

    End Sub

    Private Sub ButtonLSCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        If Not TrioMotionCalibrating() Then Exit Sub
        RevertData()
    End Sub

    Public Sub RevertData()
        NewLaserZReference.Text = 0
        NewLaserXOffset.Text = 0
        NewLaserYOffset.Text = 0
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'LaserReading.Text = CStr(Laser.LASER_Reading)
    End Sub

    Private Sub WaitLoop()
        Dim CurrentTime, StartTime As DateTime
        Dim ElapsedTime As TimeSpan
        StartTime = Now
        Dim ReadingReceived As Boolean = False
        While ReadingReceived = False
            TraceDoEvents()
            CurrentTime = Now
            ElapsedTime = CurrentTime.Subtract(StartTime)
            If ElapsedTime.TotalSeconds > 2 Then
                ReadingReceived = True
            End If
        End While
    End Sub

    Private startInit As AutoResetEvent = New AutoResetEvent(False)
    Private errorState As AutoResetEvent = New AutoResetEvent(False)
    Private abortInit As AutoResetEvent = New AutoResetEvent(False)
    Private waitHandles() As WaitHandle = {abortInit, errorState, startInit}
    Public isInited As Boolean = False
    Public isExited As Boolean = False
    Public errorMessage As String = ""
    Private InitThread As Threading.Thread
    Public Sub StartLaserBlockCalibration()
        InitThread = New System.Threading.Thread(AddressOf ControllerThread)
        InitThread.Priority = ThreadPriority.Normal
        InitThread.Start()
        startInit.Set()
    End Sub
    Public Sub ExitLaserCalibration()
        abortInit.Set()
    End Sub
    'Thread
    Public Sub ControllerThread()
        Dim stepCnt As Integer = 0
        Dim counter As Integer = 0
        Dim calibrator_plate_height, block_height, difference As Double
        Dim position(3) As Double
        Dim center(1) As Double
        Dim bigStep(1) As Double
        Dim smallStep(1) As Double
        Dim stepSize As Double = 6.0
        Dim index As Integer = 0
        While (1)
            Dim i As Integer = WaitHandle.WaitAny(waitHandles)
            Select Case i
                Case 0
                    'Abort
                    Console.WriteLine("Laser blob calibration Thread exit #1")
                    abortInit.Reset()
                    isExited = True
                    InitThread.Abort() 'Kill the thread if sucess
                    Return
                Case 1 'Error occured
                    stepCnt = 0
                    counter = 0
                    startInit.Reset()
                    errorState.Reset()
                    MessageBox.Show("Error when calibrate laser:" + errorMessage)
                Case 2
                    startInit.Set()
                    If stepCnt = 0 Then
                        index = 0
                        counter = 0
                        m_Tri.SetMachineRun()
                        Laser.EnableContinuousRead()
                        Laser.LASER_Reading = 0
                        WaitLoop()
                        block_height = Laser.LASER_Reading
                        NewLaserZReference.Text = block_height
                        IDS.Data.Hardware.HeightSensor.Laser.HeightReference = block_height
                        IDS.Data.Hardware.HeightSensor.Laser.CurrentPos.Z = block_height
                        IDS.Data.SaveData()
                        m_Tri.Set_XY_Speed(10)
                        center(0) = m_Tri.XPosition
                        center(1) = m_Tri.YPosition
                        bigStep(0) = stepSize 'Move + X 
                        bigStep(1) = 0
                        m_Tri.MoveRelative_XY(bigStep)
                        WaitLoop()
                        calibrator_plate_height = Laser.LASER_Reading
                        smallStep(0) = -0.01
                        smallStep(1) = 0
                        stepCnt += 1
                    ElseIf stepCnt = 1 Then
                        If Math.Abs(calibrator_plate_height - Laser.LASER_Reading) < 0.95 And counter <= 400 Then
                            counter = counter + 1
                            LaserReading.Text = Laser.LASER_Reading
                            m_Tri.MoveRelative_XY(smallStep)
                        ElseIf Math.Abs(calibrator_plate_height - Laser.LASER_Reading) > 0.95 Then 'sharp 
                            'ElseIf Math.Abs(9.8) > 0.95 Then 'sharp drop
                            m_Tri.GetIDSState()
                            If (index < 2) Then  'Convert to system coordinate
                                position(index) = m_Tri.XPosition
                                bigStep(0) *= -1
                                bigStep(1) = 0
                                smallStep(0) *= -1
                                smallStep(1) = 0
                                If (index = 1) Then
                                    bigStep(0) = 0
                                    bigStep(1) = stepSize 'next step is Y offset from center
                                    smallStep(0) = 0
                                    smallStep(1) = -0.01
                                End If
                            Else                 'Convert to system coordinate
                                position(index) = m_Tri.YPosition
                                bigStep(0) = 0
                                bigStep(1) *= -1
                                smallStep(0) = 0
                                smallStep(1) *= -1
                            End If
                            index += 1
                            stepCnt += 1
                        End If
                        If counter >= 400 Then
                            errorMessage = "Counter exceed limit but sharp drop is not detected!"
                            errorState.Set()
                            counter = 0
                        End If
                    ElseIf stepCnt = 2 Then
                        If index < 4 Then
                            counter = 0
                            m_Tri.Move_XY(center)
                            m_Tri.MoveRelative_XY(bigStep)
                            WaitLoop()
                        End If
                        If index = 4 Then 'Success
                            stepCnt += 1
                        Else
                            stepCnt -= 1
                        End If
                    ElseIf stepCnt = 3 Then
                        distance(0) = (position(0) - position(1)) / 2 + position(1)
                        distance(1) = (position(2) - position(3)) / 2 + position(3)
                        distance(2) = 0
                        m_Tri.Set_XY_Speed(10)
                        m_Tri.Move_XY(distance)
                        NewLaserXOffset.Text = distance(0)
                        NewLaserYOffset.Text = distance(1)
                        IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X = distance(0) - IDS.Data.Hardware.Camera.ReferencePos.X
                        IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y = distance(1) - IDS.Data.Hardware.Camera.ReferencePos.Y
                        NewLaserYOffset.Text = RoundTo3DecimalPoints(IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X)
                        NewLaserXOffset.Text = RoundTo3DecimalPoints(IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y)
                        LaserOffsetX.Text = "X:" & RoundTo3DecimalPoints(IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X)
                        LaserOffsetY.Text = "Y:" & RoundTo3DecimalPoints(IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y)
                        'LasertoCameraOffset.Text = "Laser to camera offset X:" & IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.X & " Y:" & IDS.Data.Hardware.HeightSensor.Laser.OffsetPos.Y
                        IDS.Data.SaveData()
                        'after successful setup disable the following
                        Laser.DisableContinuousRead()
                        LaserReading.Text = ""
                        m_Tri.ResetCalibrationFlag()
                        m_Tri.SetMachineStop()
                        stepCnt = 0
                        counter = 0
                        index = 0
                        ButtonLSStart.Text = "Start"
                        startInit.Reset()
                        MessageBox.Show("Laser to camera calibration done and data was saved. Laser offset X:" & LaserOffsetX.Text & " Y:" & LaserOffsetY.Text)
                        InitThread.Abort() 'Kill the thread if sucess
                    End If
                    'Init
            End Select
        End While
    End Sub
    Public Function RoundTo3DecimalPoints(ByVal value As Double) As Double
        value = CInt(value * 1000.0) / 1000.0
        Return value
    End Function
    Private Sub HardToSystem_X(ByVal pt As Double, ByRef newPt As Double)
        newPt = pt - IDS.Data.Hardware.Gantry.SystemOriginPos.X
    End Sub

    Private Sub LaserOffsetZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaserOffsetZ.Click

    End Sub
End Class
