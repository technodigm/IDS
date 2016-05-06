Imports DLL_Export_Service

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
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
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
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetX As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetY As System.Windows.Forms.Label
    Friend WithEvents LaserOffsetZ As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(LaserCalibration))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Label4 = New System.Windows.Forms.Label
        Me.NewLaserZReference = New System.Windows.Forms.Label
        Me.NewLaserYOffset = New System.Windows.Forms.Label
        Me.NewLaserXOffset = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonLSStart = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label11 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.LaserOffsetX = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.LaserOffsetY = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.LaserOffsetZ = New System.Windows.Forms.Label
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
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(472, 760)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Position Offset of Laser Sensing Point"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.ProgressBar1)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.NewLaserZReference)
        Me.GroupBox2.Controls.Add(Me.NewLaserYOffset)
        Me.GroupBox2.Controls.Add(Me.NewLaserXOffset)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.ButtonLSStart)
        Me.GroupBox2.Location = New System.Drawing.Point(40, 224)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(392, 376)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Locate Offset"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(40, 96)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(320, 22)
        Me.ProgressBar1.TabIndex = 0
        Me.ProgressBar1.Value = 50
        Me.ProgressBar1.Visible = False
        '
        'Label4
        '
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(80, 144)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 25)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Current X"
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
        Me.Label7.Location = New System.Drawing.Point(80, 240)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(160, 26)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Current Z"
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(80, 192)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 30)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Current Y"
        '
        'Label5
        '
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label5.Location = New System.Drawing.Point(32, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(139, 32)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "Offset Result"
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
        '
        'Label2
        '
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label2.Location = New System.Drawing.Point(32, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(222, 30)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "2. Click Start Button"
        '
        'Label3
        '
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(32, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(432, 48)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "1. Jog laser sensing point to the approximate center of Height Reference Block"
        '
        'ButtonSave
        '
        Me.ButtonSave.Location = New System.Drawing.Point(296, 680)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(75, 40)
        Me.ButtonSave.TabIndex = 6
        Me.ButtonSave.Text = "Save"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(-102, -8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 21)
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
        Me.Panel2.Controls.Add(Me.Label12)
        Me.Panel2.Controls.Add(Me.Label15)
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.Label17)
        Me.Panel2.Controls.Add(Me.LaserOffsetX)
        Me.Panel2.Controls.Add(Me.Label19)
        Me.Panel2.Controls.Add(Me.LaserOffsetY)
        Me.Panel2.Controls.Add(Me.Label21)
        Me.Panel2.Controls.Add(Me.LaserOffsetZ)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(12, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(488, 848)
        Me.Panel2.TabIndex = 34
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(368, 32)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(16, 23)
        Me.Label12.TabIndex = 82
        Me.Label12.Text = ","
        Me.Label12.Visible = False
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(272, 32)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(16, 23)
        Me.Label15.TabIndex = 81
        Me.Label15.Text = ","
        Me.Label15.Visible = False
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(16, 32)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(168, 23)
        Me.Label16.TabIndex = 35
        Me.Label16.Text = "Laser Position Offset :"
        Me.Label16.Visible = False
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(464, 32)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(16, 23)
        Me.Label17.TabIndex = 80
        Me.Label17.Text = ")"
        Me.Label17.Visible = False
        '
        'LaserOffsetX
        '
        Me.LaserOffsetX.Location = New System.Drawing.Point(400, 32)
        Me.LaserOffsetX.Name = "LaserOffsetX"
        Me.LaserOffsetX.Size = New System.Drawing.Size(72, 23)
        Me.LaserOffsetX.TabIndex = 74
        Me.LaserOffsetX.Text = "100.000"
        Me.LaserOffsetX.Visible = False
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(184, 32)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(16, 23)
        Me.Label19.TabIndex = 75
        Me.Label19.Text = "("
        Me.Label19.Visible = False
        '
        'LaserOffsetY
        '
        Me.LaserOffsetY.Location = New System.Drawing.Point(296, 32)
        Me.LaserOffsetY.Name = "LaserOffsetY"
        Me.LaserOffsetY.Size = New System.Drawing.Size(72, 23)
        Me.LaserOffsetY.TabIndex = 77
        Me.LaserOffsetY.Text = "100.000"
        Me.LaserOffsetY.Visible = False
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(322, 32)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(16, 23)
        Me.Label21.TabIndex = 78
        Me.Label21.Text = ","
        Me.Label21.Visible = False
        '
        'LaserOffsetZ
        '
        Me.LaserOffsetZ.Location = New System.Drawing.Point(392, 32)
        Me.LaserOffsetZ.Name = "LaserOffsetZ"
        Me.LaserOffsetZ.Size = New System.Drawing.Size(72, 23)
        Me.LaserOffsetZ.TabIndex = 79
        Me.LaserOffsetZ.Text = "100.000"
        Me.LaserOffsetZ.Visible = False
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
            .OffsetPos.X = 52.819 - IDS.Data.Hardware.Camera.ReferencePos.X
            .OffsetPos.Y = 127.271 - IDS.Data.Hardware.Camera.ReferencePos.Y
        End With

        IDS.Data.SaveData()
    End Sub


    Private Sub ButtonLSStart_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLSStart.Click

        Dim counter As Integer = 0
        Dim calibrator_plate_height, block_height As Double

        m_Tri.SetMachineRun()
        OnLaser()
        If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        block_height = Laser.MM_Reading
        NewLaserZReference.Text = block_height
        IDS.Data.Hardware.HeightSensor.Laser.HeightReference = block_height
        IDS.Data.Hardware.HeightSensor.Laser.CurrentPos.Z = block_height
        IDS.Data.SaveData()

        m_Tri.Set_XY_Speed(10)

        distance(0) = 7
        distance(1) = 7
        m_Tri.MoveRelative_XY(distance)
        If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        If (Laser.MM_Reading - block_height) > 0.95 And (Laser.MM_Reading - block_height) < 2 Then
            distance(0) = 0
            distance(1) = -7
            m_Tri.MoveRelative_XY(distance)
        Else
            GoTo CalibrationError
        End If

        calibrator_plate_height = Laser.MM_Reading
        distance(0) = -0.01
        distance(1) = 0
        While Math.Abs(calibrator_plate_height - Laser.MM_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        End While

        distance(0) = -9
        distance(1) = 0
        distance(2) = 0
        If Math.Abs(calibrator_plate_height - Laser.MM_Reading) > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            m_Tri.m_TriCtrl.SetTable(101, 1, m_Tri.XPosition) 'X rising pos
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        Else
            GoTo CalibrationError
        End If

        counter = 0
        If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        block_height = Laser.MM_Reading
        distance(0) = -0.01
        distance(1) = 0
        While Math.Abs(block_height - Laser.MM_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        End While

        distance(0) = 5
        distance(1) = 6
        If Math.Abs(block_height - Laser.MM_Reading) > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            m_Tri.m_TriCtrl.SetTable(102, 1, m_Tri.XPosition) 'X rising pos
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        Else
            GoTo CalibrationError
        End If

        distance(0) = 0
        distance(1) = -0.01
        While Math.Abs(calibrator_plate_height - Laser.MM_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        End While

        distance(0) = 0
        distance(1) = -9
        If Math.Abs(calibrator_plate_height - Laser.MM_Reading) > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            m_Tri.m_TriCtrl.SetTable(103, 1, m_Tri.YPosition) 'X rising pos
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        Else
            GoTo CalibrationError
        End If

        counter = 0
        distance(0) = 0
        distance(1) = -0.01
        While Math.Abs(block_height - Laser.MM_Reading) < 0.95 And counter <= 400
            counter = counter + 1
            m_Tri.MoveRelative_XY(distance)
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        End While

        If Math.Abs(block_height - Laser.MM_Reading) > 0.95 And counter < 400 Then
            m_Tri.GetIDSState()
            m_Tri.m_TriCtrl.SetTable(104, 1, m_Tri.YPosition) 'X rising pos
            If Not Laser.WaitForReadingToStabilize() Then GoTo calibrationerror
        Else
            GoTo CalibrationError
        End If

        'finish
        TraceDoEvents()
        Dim position(3) As Double
        m_Tri.m_TriCtrl.GetTable(101, 4, position)
        distance(0) = (position(0) - position(1)) / 2 + position(1)
        distance(1) = (position(2) - position(3)) / 2 + position(3)
        distance(2) = 0
        m_Tri.Set_XY_Speed(10)
        m_Tri.Move_XY(distance)
        MessageBox.Show("finish")
        NewLaserXOffset.Text = distance(0)
        NewLaserYOffset.Text = distance(1)
        m_Tri.SetMachineStop()
        Exit Sub

CalibrationError:
        OffLaser()
        m_Tri.ResetCalibrationFlag()
        MessageBox.Show("Error")
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

End Class
