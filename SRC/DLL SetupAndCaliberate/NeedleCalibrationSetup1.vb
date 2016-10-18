Imports DLL_Export_Service

Public Class NeedleCalibrationSetup1
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
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonCapture As System.Windows.Forms.Button
    Friend WithEvents NSensorXPos As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents NSensorYPos As System.Windows.Forms.Label
    Friend WithEvents NSensorZPos As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents LNeedleZPos As System.Windows.Forms.Label
    Friend WithEvents LNeedleYPos As System.Windows.Forms.Label
    Friend WithEvents LNeedleXPos As System.Windows.Forms.Label
    Friend WithEvents RNeedleZPos As System.Windows.Forms.Label
    Friend WithEvents RNeedleYPos As System.Windows.Forms.Label
    Friend WithEvents RNeedleXPos As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblBrightness As System.Windows.Forms.Label
    Friend WithEvents txtBrightness As System.Windows.Forms.TextBox
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSetup1))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label16 = New System.Windows.Forms.Label
        Me.RNeedleZPos = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.RNeedleYPos = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.RNeedleXPos = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.LNeedleZPos = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.LNeedleYPos = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.LNeedleXPos = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.NSensorZPos = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.NSensorYPos = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.NSensorXPos = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtBrightness = New System.Windows.Forms.TextBox
        Me.lblBrightness = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonCapture = New System.Windows.Forms.Button
        Me.ButtonNext = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Panel2)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(8, 24)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 39
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(272, 32)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Needle Calibration Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(408, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 45
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label16)
        Me.Panel2.Controls.Add(Me.RNeedleZPos)
        Me.Panel2.Controls.Add(Me.Label18)
        Me.Panel2.Controls.Add(Me.RNeedleYPos)
        Me.Panel2.Controls.Add(Me.Label20)
        Me.Panel2.Controls.Add(Me.Label21)
        Me.Panel2.Controls.Add(Me.RNeedleXPos)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.LNeedleZPos)
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.LNeedleYPos)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.LNeedleXPos)
        Me.Panel2.Controls.Add(Me.Label11)
        Me.Panel2.Controls.Add(Me.NSensorZPos)
        Me.Panel2.Controls.Add(Me.Label7)
        Me.Panel2.Controls.Add(Me.NSensorYPos)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.NSensorXPos)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(512, 864)
        Me.Panel2.TabIndex = 44
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(480, 128)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(16, 23)
        Me.Label16.TabIndex = 80
        Me.Label16.Text = ")"
        Me.Label16.Visible = False
        '
        'RNeedleZPos
        '
        Me.RNeedleZPos.Location = New System.Drawing.Point(416, 128)
        Me.RNeedleZPos.Name = "RNeedleZPos"
        Me.RNeedleZPos.Size = New System.Drawing.Size(72, 23)
        Me.RNeedleZPos.TabIndex = 79
        Me.RNeedleZPos.Text = "0"
        Me.RNeedleZPos.Visible = False
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(392, 128)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(16, 23)
        Me.Label18.TabIndex = 78
        Me.Label18.Text = ","
        Me.Label18.Visible = False
        '
        'RNeedleYPos
        '
        Me.RNeedleYPos.Location = New System.Drawing.Point(328, 128)
        Me.RNeedleYPos.Name = "RNeedleYPos"
        Me.RNeedleYPos.Size = New System.Drawing.Size(72, 23)
        Me.RNeedleYPos.TabIndex = 77
        Me.RNeedleYPos.Text = "0"
        Me.RNeedleYPos.Visible = False
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(304, 128)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(16, 23)
        Me.Label20.TabIndex = 76
        Me.Label20.Text = ","
        Me.Label20.Visible = False
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(224, 128)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(16, 23)
        Me.Label21.TabIndex = 75
        Me.Label21.Text = "("
        Me.Label21.Visible = False
        '
        'RNeedleXPos
        '
        Me.RNeedleXPos.Location = New System.Drawing.Point(240, 128)
        Me.RNeedleXPos.Name = "RNeedleXPos"
        Me.RNeedleXPos.Size = New System.Drawing.Size(72, 23)
        Me.RNeedleXPos.TabIndex = 74
        Me.RNeedleXPos.Text = "0"
        Me.RNeedleXPos.Visible = False
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(480, 72)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(16, 23)
        Me.Label8.TabIndex = 73
        Me.Label8.Text = ")"
        Me.Label8.Visible = False
        '
        'LNeedleZPos
        '
        Me.LNeedleZPos.Location = New System.Drawing.Point(416, 72)
        Me.LNeedleZPos.Name = "LNeedleZPos"
        Me.LNeedleZPos.Size = New System.Drawing.Size(72, 23)
        Me.LNeedleZPos.TabIndex = 72
        Me.LNeedleZPos.Text = "0"
        Me.LNeedleZPos.Visible = False
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(392, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(16, 23)
        Me.Label10.TabIndex = 71
        Me.Label10.Text = ","
        Me.Label10.Visible = False
        '
        'LNeedleYPos
        '
        Me.LNeedleYPos.Location = New System.Drawing.Point(328, 72)
        Me.LNeedleYPos.Name = "LNeedleYPos"
        Me.LNeedleYPos.Size = New System.Drawing.Size(72, 23)
        Me.LNeedleYPos.TabIndex = 70
        Me.LNeedleYPos.Text = "0"
        Me.LNeedleYPos.Visible = False
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(304, 72)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(16, 23)
        Me.Label13.TabIndex = 69
        Me.Label13.Text = ","
        Me.Label13.Visible = False
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(224, 72)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(16, 23)
        Me.Label14.TabIndex = 68
        Me.Label14.Text = "("
        Me.Label14.Visible = False
        '
        'LNeedleXPos
        '
        Me.LNeedleXPos.Location = New System.Drawing.Point(240, 72)
        Me.LNeedleXPos.Name = "LNeedleXPos"
        Me.LNeedleXPos.Size = New System.Drawing.Size(72, 23)
        Me.LNeedleXPos.TabIndex = 67
        Me.LNeedleXPos.Text = "0"
        Me.LNeedleXPos.Visible = False
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(480, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(16, 23)
        Me.Label11.TabIndex = 66
        Me.Label11.Text = ")"
        Me.Label11.Visible = False
        '
        'NSensorZPos
        '
        Me.NSensorZPos.Location = New System.Drawing.Point(416, 16)
        Me.NSensorZPos.Name = "NSensorZPos"
        Me.NSensorZPos.Size = New System.Drawing.Size(72, 23)
        Me.NSensorZPos.TabIndex = 64
        Me.NSensorZPos.Text = "0"
        Me.NSensorZPos.Visible = False
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(392, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(16, 23)
        Me.Label7.TabIndex = 63
        Me.Label7.Text = ","
        Me.Label7.Visible = False
        '
        'NSensorYPos
        '
        Me.NSensorYPos.Location = New System.Drawing.Point(328, 16)
        Me.NSensorYPos.Name = "NSensorYPos"
        Me.NSensorYPos.Size = New System.Drawing.Size(72, 23)
        Me.NSensorYPos.TabIndex = 62
        Me.NSensorYPos.Text = "0"
        Me.NSensorYPos.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(304, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 23)
        Me.Label4.TabIndex = 61
        Me.Label4.Text = ","
        Me.Label4.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(224, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 23)
        Me.Label2.TabIndex = 60
        Me.Label2.Text = "("
        Me.Label2.Visible = False
        '
        'NSensorXPos
        '
        Me.NSensorXPos.Location = New System.Drawing.Point(240, 16)
        Me.NSensorXPos.Name = "NSensorXPos"
        Me.NSensorXPos.Size = New System.Drawing.Size(72, 23)
        Me.NSensorXPos.TabIndex = 59
        Me.NSensorXPos.Text = "0"
        Me.NSensorXPos.Visible = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(24, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(200, 40)
        Me.Label6.TabIndex = 58
        Me.Label6.Text = "Right Needle X, Y, Z Position Offset: "
        Me.Label6.Visible = False
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(24, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(200, 40)
        Me.Label5.TabIndex = 57
        Me.Label5.Text = "Left Needle X, Y, Z Position Offset: "
        Me.Label5.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 41)
        Me.Label1.TabIndex = 43
        Me.Label1.Text = "Needle Calibrator Sensor Position Offset: "
        Me.Label1.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtBrightness)
        Me.GroupBox1.Controls.Add(Me.lblBrightness)
        Me.GroupBox1.Controls.Add(Me.PictureBox1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ButtonCapture)
        Me.GroupBox1.Controls.Add(Me.ButtonNext)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 184)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 672)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Locate Needle Calibrator Sensor"
        '
        'txtBrightness
        '
        Me.txtBrightness.Location = New System.Drawing.Point(208, 120)
        Me.txtBrightness.Name = "txtBrightness"
        Me.txtBrightness.Size = New System.Drawing.Size(80, 27)
        Me.txtBrightness.TabIndex = 14
        Me.txtBrightness.Text = "31"
        '
        'lblBrightness
        '
        Me.lblBrightness.Location = New System.Drawing.Point(96, 120)
        Me.lblBrightness.Name = "lblBrightness"
        Me.lblBrightness.Size = New System.Drawing.Size(112, 16)
        Me.lblBrightness.TabIndex = 13
        Me.lblBrightness.Text = "Brightness"
        '
        'PictureBox1
        '
        Me.PictureBox1.AccessibleName = ""
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(152, 184)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(191, 159)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'Label3
        '
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(24, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(460, 48)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "1. Jog camera to the approximate center of Needle Calibrator Sensor and click Cap" & _
        "ture Image Button."
        '
        'ButtonCapture
        '
        Me.ButtonCapture.Location = New System.Drawing.Point(200, 368)
        Me.ButtonCapture.Name = "ButtonCapture"
        Me.ButtonCapture.Size = New System.Drawing.Size(100, 50)
        Me.ButtonCapture.TabIndex = 9
        Me.ButtonCapture.Text = "Capture and Save"
        '
        'ButtonNext
        '
        Me.ButtonNext.Location = New System.Drawing.Point(368, 600)
        Me.ButtonNext.Name = "ButtonNext"
        Me.ButtonNext.Size = New System.Drawing.Size(100, 50)
        Me.ButtonNext.TabIndex = 6
        Me.ButtonNext.Text = "Next"
        '
        'NeedleCalibrationSetup1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSetup1"
        Me.Text = "NeedleCalibrationSetup1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Friend Shared Offset_X, Offset_Y As Double
    Shared vParam As DLL_Export_Device_Vision.QC.QCParam

    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNext.Click
        RemovePanel(PanelToBeAdded)
        AddPanel(MySetup.PanelRight, MyNeedleCalibrationSetup2.PanelToBeAdded)
    End Sub

    Friend Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub

    Private Sub ButtonCapture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCapture.Click
        If (CInt(txtBrightness.Text) <= 0) Then
            Exit Sub
        End If

        IDS.Devices.Vision.IDSV_SetBrightness(CInt(txtBrightness.Text))

        If IDS.Devices.Vision.IDSV_NC(False, 200, 5, 1, 3, 3, 3, 3, 768 / 2, 576 / 2, 766, 572, Offset_X, Offset_Y) Then
            Vision.FrmVision.ClearDisplay()
        Else
            Vision.FrmVision.ClearDisplay()
            MsgBox("Failed to capture the touch sensor.")
            Exit Sub
        End If

        m_Tri.GetIDSState()
        Dim z As Integer = 1 + 1
        IDS.Data.Hardware.Needle.Right.CalibratorPos.X = m_Tri.XPosition + Offset_X
        IDS.Data.Hardware.Needle.Right.CalibratorPos.Y = m_Tri.YPosition - Offset_Y
        IDS.Data.Hardware.Needle.Left.CalibratorPos.X = m_Tri.XPosition + Offset_X
        IDS.Data.Hardware.Needle.Left.CalibratorPos.Y = m_Tri.YPosition - Offset_Y

        distance(0) = Offset_X
        distance(1) = -Offset_Y

        m_Tri.MoveRelative_XY(distance)

        IDS.Data.SaveData()
        NSensorXPos.Text = IDS.Data.Hardware.Needle.Right.CalibratorPos.X - IDS.Data.Hardware.Camera.ReferencePos.X
        NSensorYPos.Text = IDS.Data.Hardware.Needle.Right.CalibratorPos.Y - IDS.Data.Hardware.Camera.ReferencePos.Y
        MessageBox.Show("Touch sensor captured successfully.")
    End Sub

    Public Sub RevertData()
        Para.RecordID = "FactoryDefault"
        Para.GroupID = IDS.Data.Admin.User.Group.ID
        IDS.Data.OpenData()

        LNeedleXPos.Text = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.X
        LNeedleYPos.Text = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Y
        LNeedleZPos.Text = IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z

        RNeedleXPos.Text = IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.X
        RNeedleYPos.Text = IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.Y
        RNeedleZPos.Text = IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.Z
    End Sub

End Class
