Imports System.IO
Imports DLL_Export_Service
Imports DLL_DataManager 'yy

Public Class NeedleCalibrationSettings
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
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents ButtonCalibrate As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ZOffset As System.Windows.Forms.Label
    Friend WithEvents YOffset As System.Windows.Forms.Label
    Friend WithEvents XOffset As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents BtMoveToCalibratorPost As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbLeftHead As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Public WithEvents rbRightHead As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSettings))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.BtMoveToCalibratorPost = New System.Windows.Forms.Button
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ZOffset = New System.Windows.Forms.Label
        Me.YOffset = New System.Windows.Forms.Label
        Me.XOffset = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rbRightHead = New System.Windows.Forms.RadioButton
        Me.rbLeftHead = New System.Windows.Forms.RadioButton
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox3)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.TextBox1)
        Me.PanelToBeAdded.Controls.Add(Me.Label21)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(24, 16)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 911)
        Me.PanelToBeAdded.TabIndex = 68
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.BtMoveToCalibratorPost)
        Me.GroupBox3.Controls.Add(Me.ButtonCalibrate)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Location = New System.Drawing.Point(100, 552)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(312, 320)
        Me.GroupBox3.TabIndex = 107
        Me.GroupBox3.TabStop = False
        '
        'BtMoveToCalibratorPost
        '
        Me.BtMoveToCalibratorPost.Location = New System.Drawing.Point(40, 120)
        Me.BtMoveToCalibratorPost.Name = "BtMoveToCalibratorPost"
        Me.BtMoveToCalibratorPost.Size = New System.Drawing.Size(232, 80)
        Me.BtMoveToCalibratorPost.TabIndex = 97
        Me.BtMoveToCalibratorPost.Text = "Move to Reference Point"
        Me.BtMoveToCalibratorPost.Visible = False
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.Location = New System.Drawing.Point(40, 224)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(232, 80)
        Me.ButtonCalibrate.TabIndex = 61
        Me.ButtonCalibrate.Text = "Save Calibration Data"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(40, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(232, 80)
        Me.Button1.TabIndex = 61
        Me.Button1.Text = "Move to Calibrated Position"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.ZOffset)
        Me.GroupBox2.Controls.Add(Me.YOffset)
        Me.GroupBox2.Controls.Add(Me.XOffset)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.Label6)
        Me.GroupBox2.Location = New System.Drawing.Point(100, 424)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(312, 128)
        Me.GroupBox2.TabIndex = 106
        Me.GroupBox2.TabStop = False
        '
        'Label4
        '
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(64, 19)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 25)
        Me.Label4.TabIndex = 91
        Me.Label4.Text = "X offset :"
        '
        'ZOffset
        '
        Me.ZOffset.Location = New System.Drawing.Point(160, 83)
        Me.ZOffset.Name = "ZOffset"
        Me.ZOffset.Size = New System.Drawing.Size(88, 24)
        Me.ZOffset.TabIndex = 96
        Me.ZOffset.Text = "0.000"
        '
        'YOffset
        '
        Me.YOffset.Location = New System.Drawing.Point(160, 51)
        Me.YOffset.Name = "YOffset"
        Me.YOffset.Size = New System.Drawing.Size(88, 24)
        Me.YOffset.TabIndex = 95
        Me.YOffset.Text = "0.000"
        '
        'XOffset
        '
        Me.XOffset.Location = New System.Drawing.Point(160, 19)
        Me.XOffset.Name = "XOffset"
        Me.XOffset.Size = New System.Drawing.Size(88, 24)
        Me.XOffset.TabIndex = 94
        Me.XOffset.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(64, 83)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 26)
        Me.Label7.TabIndex = 93
        Me.Label7.Text = "Z offset :"
        '
        'Label6
        '
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(64, 51)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 30)
        Me.Label6.TabIndex = 92
        Me.Label6.Text = "Y offset :"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbRightHead)
        Me.GroupBox1.Controls.Add(Me.rbLeftHead)
        Me.GroupBox1.Location = New System.Drawing.Point(100, 336)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 88)
        Me.GroupBox1.TabIndex = 105
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Current Head :"
        '
        'rbRightHead
        '
        Me.rbRightHead.Location = New System.Drawing.Point(174, 40)
        Me.rbRightHead.Name = "rbRightHead"
        Me.rbRightHead.Size = New System.Drawing.Size(68, 24)
        Me.rbRightHead.TabIndex = 1
        Me.rbRightHead.Text = "Right"
        '
        'rbLeftHead
        '
        Me.rbLeftHead.Checked = True
        Me.rbLeftHead.Location = New System.Drawing.Point(70, 40)
        Me.rbLeftHead.Name = "rbLeftHead"
        Me.rbLeftHead.Size = New System.Drawing.Size(60, 24)
        Me.rbLeftHead.TabIndex = 0
        Me.rbLeftHead.TabStop = True
        Me.rbLeftHead.Text = "Left"
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(30, 64)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(452, 264)
        Me.TextBox1.TabIndex = 104
        Me.TextBox1.Text = "Info:                                                                 First, make" & _
        " sure the syringe is not tighten. Press and hold Control Key on Keyboard and use" & _
        " the mouse's scroll ball to move the syringe's needle to the hardware reference " & _
        "point. After that, move down the z stage before touching the reference point. Ti" & _
        "ghten the syringe and continue with the fine adjustment to make sure the needle " & _
        "tip is aligned to reference point. Do not forget to save the calibration data. "
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.Black
        Me.Label21.Location = New System.Drawing.Point(8, 8)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(296, 32)
        Me.Label21.TabIndex = 90
        Me.Label21.Text = "Needle Calibration Settings"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 73
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonExit.Visible = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "txt"
        Me.OpenFileDialog1.Filter = "(*.txt)*|"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "txt"
        Me.SaveFileDialog1.Filter = "(*.txt)|*"
        '
        'NeedleCalibrationSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(560, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSettings"
        Me.Text = "NeedleCalibration1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        OffLaser()
        RemovePanel(CurrentControl)
        m_Tri2.Move_Z(0)
    End Sub

    Private Sub ButtonCalibratePosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click
        Dim fm As InfoForm = New InfoForm
        fm.SetMessage("Are you sure you want to save the calibration data?")
        fm.ShowDialog()
        If fm.DialogResult = DialogResult.Cancel Then
            Return
        End If
        If rbLeftHead.Checked Then
            With IDS.Data.Hardware.Needle.Left
                .NeedleCalibrationPosition.X = .CalibratorPos.X - m_Tri2.XPosition()
                .NeedleCalibrationPosition.Y = .CalibratorPos.Y - m_Tri2.YPosition()
                .NeedleCalibrationPosition.Z = -(.CalibratorPos.Z - m_Tri2.ZPosition())
                '.NeedleCalibrationPosition.Z = .CalibratorPos.Z - m_Tri.ZPosition()  'yy
                XOffset.Text = .NeedleCalibrationPosition.X 'yy 'calibration
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
            Console.WriteLine("#Left Needle calibaration Z offset" & IDS.Data.Hardware.Needle.Left.NeedleCalibrationPosition.Z.ToString("F3"))
        Else
            With IDS.Data.Hardware.Needle.Right
                .NeedleCalibrationPosition.X = .CalibratorPos.X - m_Tri2.XPosition()
                .NeedleCalibrationPosition.Y = .CalibratorPos.Y - m_Tri2.YPosition()
                .NeedleCalibrationPosition.Z = -(.CalibratorPos.Z - m_Tri2.ZPosition())
                '.NeedleCalibrationPosition.Z = .CalibratorPos.Z - m_Tri.ZPosition()  'yy
                XOffset.Text = .NeedleCalibrationPosition.X 'yy 'calibration
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
            Console.WriteLine("#Right Needle calibaration Z offset" & IDS.Data.Hardware.Needle.Right.NeedleCalibrationPosition.Z.ToString("F3"))
        End If

        'IDS.Data.SaveLocalData()
        IDS.Data.SaveGlobalData()
        ' get default value from the default pat file
    End Sub

    Public Sub Revert()
        'IDS.Data.OpenData() 'yy
        OpenPathFileName("C:\IDS\Pattern_Dir\factorydefault.pat")
        If rbLeftHead.Checked Then
            With IDS.Data.Hardware.Needle.Left
                XOffset.Text = .NeedleCalibrationPosition.X
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
        Else
            With IDS.Data.Hardware.Needle.Right
                XOffset.Text = .NeedleCalibrationPosition.X
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim fm As InfoForm = New InfoForm
        fm.Size = New System.Drawing.Size(584, 450)
        fm.SetMessage("Warning:")
        fm.SetMessage("Please make sure syringe is not tighthen!!")
        fm.AddNewLine(" ")
        fm.AddNewLine("There is a chance that the syringe holder can crash with the calibration stage if of the following:")
        fm.AddNewLine(" ")
        fm.AddNewLine("#1: If existing syringe holder was replaced with other syringe holder with different size or type.")
        fm.AddNewLine("#2: If existing syringe holder was reinstalled at different position.")
        fm.AddNewLine(" ")
        fm.AddNewLine("Click OK to contine, otherwise click Cancel to abort the process.")
        If fm.ShowDialog() = DialogResult.Cancel Then
            Return
        End If
        Dim pos(1) As Double
        SetServiceSpeed()
        If rbLeftHead.Checked Then
            With IDS.Data.Hardware.Needle.Left
                pos(0) = .CalibratorPos.X - .NeedleCalibrationPosition.X
                pos(1) = .CalibratorPos.Y - .NeedleCalibrationPosition.Y
                If Not m_Tri2.Move_Z(0) Then Exit Sub
                If Not m_Tri2.Move_XY(pos) Then Exit Sub
                'm_Tri2.Set_Z_Speed(5)
                If Not m_Tri2.Move_Z(.CalibratorPos.Z + .NeedleCalibrationPosition.Z) Then Exit Sub
            End With
        Else
            With IDS.Data.Hardware.Needle.Right
                pos(0) = .CalibratorPos.X - .NeedleCalibrationPosition.X
                pos(1) = .CalibratorPos.Y - .NeedleCalibrationPosition.Y
                If Not m_Tri2.Move_Z(0) Then Exit Sub
                If Not m_Tri2.Move_XY(pos) Then Exit Sub
                'm_Tri2.Set_Z_Speed(5)
                If Not m_Tri2.Move_Z(.CalibratorPos.Z + .NeedleCalibrationPosition.Z) Then Exit Sub
            End With
        End If

    End Sub

    Private Sub BtMoveToCalibratorPost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtMoveToCalibratorPost.Click
        Dim pos(1) As Double
        SetServiceSpeed()
        If rbLeftHead.Checked Then
            With IDS.Data.Hardware.Needle.Left
                pos(0) = .CalibratorPos.X
                pos(1) = .CalibratorPos.Y
                If Not m_Tri2.Move_Z(0) Then Exit Sub
                If Not m_Tri2.Move_XY(pos) Then Exit Sub
                'm_Tri2.Set_Z_Speed(5)
                If Not m_Tri2.Move_Z(.CalibratorPos.Z) Then Exit Sub
            End With
        Else
            With IDS.Data.Hardware.Needle.Right
                pos(0) = .CalibratorPos.X
                pos(1) = .CalibratorPos.Y
                If Not m_Tri2.Move_Z(0) Then Exit Sub
                If Not m_Tri2.Move_XY(pos) Then Exit Sub
                'm_Tri2.Set_Z_Speed(5)
                If Not m_Tri2.Move_Z(.CalibratorPos.Z) Then Exit Sub
            End With
        End If
    End Sub
    Private Sub SetServiceSpeed()
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ServiceXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ServiceZSpeed)
    End Sub

    Private Sub rbLeftHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLeftHead.CheckedChanged
        If rbLeftHead.Checked Then
            With IDS.Data.Hardware.Needle.Left
                XOffset.Text = .NeedleCalibrationPosition.X
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
        Else
            With IDS.Data.Hardware.Needle.Right
                XOffset.Text = .NeedleCalibrationPosition.X
                YOffset.Text = .NeedleCalibrationPosition.Y
                ZOffset.Text = .NeedleCalibrationPosition.Z
            End With
        End If
    End Sub

    Private Sub BtMoveToCalibratorPost_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtMoveToCalibratorPost.DoubleClick

    End Sub

    Private Sub NeedleCalibrationSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
