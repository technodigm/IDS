Imports DLL_Export_Service

Public Class NeedleCalibrationSetup3
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonStart As System.Windows.Forms.Button
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents ButtonBack As System.Windows.Forms.Button
    Friend WithEvents NSensorZPos As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonStart = New System.Windows.Forms.Button
        Me.ButtonBack = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonNext = New System.Windows.Forms.Button
        Me.NSensorZPos = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 32)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(632, 944)
        Me.PanelToBeAdded.TabIndex = 46
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(272, 32)
        Me.Label9.TabIndex = 47
        Me.Label9.Text = "Needle Calibration Setup"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ButtonStart)
        Me.GroupBox1.Controls.Add(Me.ButtonBack)
        Me.GroupBox1.Controls.Add(Me.ButtonRevert)
        Me.GroupBox1.Controls.Add(Me.ButtonNext)
        Me.GroupBox1.Controls.Add(Me.NSensorZPos)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 184)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 672)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Needle Calibrator Sensor - Z Reference Level"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label5.Location = New System.Drawing.Point(24, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(456, 48)
        Me.Label5.TabIndex = 47
        Me.Label5.Text = "1. Move needle to approximately touching the center of z sensor"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(32, 120)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(432, 200)
        Me.GroupBox2.TabIndex = 46
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Note: "
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(16, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(396, 24)
        Me.Label6.TabIndex = 50
        Me.Label6.Text = "Upon clicking Start ..."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(16, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(408, 72)
        Me.Label4.TabIndex = 49
        Me.Label4.Text = "Needle will extend very slowly while checking the touch sensor output."
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(24, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(456, 32)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "2. Click Start Button to find the touch sensor Z Level"
        '
        'ButtonStart
        '
        Me.ButtonStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStart.Location = New System.Drawing.Point(200, 480)
        Me.ButtonStart.Name = "ButtonStart"
        Me.ButtonStart.Size = New System.Drawing.Size(100, 50)
        Me.ButtonStart.TabIndex = 9
        Me.ButtonStart.Text = "Start and Save"
        '
        'ButtonBack
        '
        Me.ButtonBack.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonBack.Location = New System.Drawing.Point(32, 600)
        Me.ButtonBack.Name = "ButtonBack"
        Me.ButtonBack.Size = New System.Drawing.Size(100, 50)
        Me.ButtonBack.TabIndex = 45
        Me.ButtonBack.Text = "Back"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(200, 600)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(100, 50)
        Me.ButtonRevert.TabIndex = 7
        Me.ButtonRevert.Text = "Cancel"
        Me.ButtonRevert.Visible = False
        '
        'ButtonNext
        '
        Me.ButtonNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNext.Location = New System.Drawing.Point(368, 600)
        Me.ButtonNext.Name = "ButtonNext"
        Me.ButtonNext.Size = New System.Drawing.Size(100, 50)
        Me.ButtonNext.TabIndex = 6
        Me.ButtonNext.Text = "Next"
        '
        'NSensorZPos
        '
        Me.NSensorZPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NSensorZPos.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.NSensorZPos.Location = New System.Drawing.Point(288, 376)
        Me.NSensorZPos.Name = "NSensorZPos"
        Me.NSensorZPos.Size = New System.Drawing.Size(144, 16)
        Me.NSensorZPos.TabIndex = 17
        Me.NSensorZPos.Visible = False
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(72, 376)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(208, 24)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Touch Sensor Z Position "
        Me.Label7.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, -32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(296, 32)
        Me.Label1.TabIndex = 48
        Me.Label1.Text = "Needle Calibrator Sensor Position Offset: (100,100,0)"
        '
        'NeedleCalibrationSetup3
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 912)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSetup3"
        Me.Text = "NeedleCalibrationSetup3"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBack.Click
        RemovePanel(PanelToBeAdded)
        AddPanel(MySetup.PanelRight, MyNeedleCalibrationSetup2.PanelToBeAdded)
    End Sub

    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNext.Click
        RemovePanel(PanelToBeAdded)
        OnLaser()
        MyNeedleCalibrationSettings.RevertData()
        AddPanel(MySetup.PanelRight, MyNeedleCalibrationSettings.PanelToBeAdded)
    End Sub

    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click

        distance(0) = 0
        distance(1) = 0
        distance(2) = 10.0
        m_Tri.MoveRelative_XYZ(distance)

        'this is a fixed xy value from the middle of the calibration block to the touch sensor
        'distance(0) = -76.005
        'distance(1) = -4.284

        'm_Tri.MoveRelative_XY(distance)

        distance(0) = 0
        distance(1) = 0
        distance(2) = -9.6 'kr jan 26 original -10.2
        m_Tri.MoveRelative_XYZ(distance)

        m_Tri.m_TriCtrl.SetTable(109, 1, 2)

        If Not TrioMotionCalibrating() Then Exit Sub

        Dim NeedleZCalFlag As Integer = m_Tri.GetCalibrationFlag
        If NeedleZCalFlag = 0 Then
            TraceDoEvents()
            m_Tri.m_TriCtrl.GetTable(101, 2, TableValues)
            IDS.Data.Hardware.Needle.Right.TouchSensorZPosition = TableValues(0)
            IDS.Data.Hardware.Needle.Left.TouchSensorZPosition = TableValues(0)
            NSensorZPos.Text = IDS.Data.Hardware.Needle.Left.TouchSensorZPosition
            IDS.Data.SaveData()
        ElseIf NeedleZCalFlag = 2 Then
            m_Tri.ResetCalibrationFlag()
            MessageBox.Show("nothing sensed")
        End If
        SetServiceSpeed()
        If Not m_Tri.Move_Z(0) Then Exit Sub
        MyNeedleCalibrationSetup3.NSensorZPos.Text = IDS.Data.Hardware.Needle.Right.TouchSensorZPosition - IDS.Data.Hardware.Needle.Left.CalibratorPos.Z

    End Sub

End Class
