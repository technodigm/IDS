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
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonSetCalibratorPosition As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ZPosition As System.Windows.Forms.Label
    Friend WithEvents YPosition As System.Windows.Forms.Label
    Friend WithEvents XPosition As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbLeftHead As System.Windows.Forms.RadioButton
    Public WithEvents rbRightHead As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(NeedleCalibrationSetup1))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rbRightHead = New System.Windows.Forms.RadioButton
        Me.rbLeftHead = New System.Windows.Forms.RadioButton
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ZPosition = New System.Windows.Forms.Label
        Me.YPosition = New System.Windows.Forms.Label
        Me.XPosition = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonSetCalibratorPosition = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.TextBox1)
        Me.PanelToBeAdded.Controls.Add(Me.Label4)
        Me.PanelToBeAdded.Controls.Add(Me.ZPosition)
        Me.PanelToBeAdded.Controls.Add(Me.YPosition)
        Me.PanelToBeAdded.Controls.Add(Me.XPosition)
        Me.PanelToBeAdded.Controls.Add(Me.Label7)
        Me.PanelToBeAdded.Controls.Add(Me.Label6)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSetCalibratorPosition)
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.Button1)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(32, 24)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 39
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rbRightHead)
        Me.GroupBox1.Controls.Add(Me.rbLeftHead)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(108, 464)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(296, 64)
        Me.GroupBox1.TabIndex = 104
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Current Head :"
        '
        'rbRightHead
        '
        Me.rbRightHead.Location = New System.Drawing.Point(184, 24)
        Me.rbRightHead.Name = "rbRightHead"
        Me.rbRightHead.Size = New System.Drawing.Size(72, 24)
        Me.rbRightHead.TabIndex = 1
        Me.rbRightHead.Text = "Right"
        '
        'rbLeftHead
        '
        Me.rbLeftHead.Checked = True
        Me.rbLeftHead.Location = New System.Drawing.Point(40, 24)
        Me.rbLeftHead.Name = "rbLeftHead"
        Me.rbLeftHead.Size = New System.Drawing.Size(64, 24)
        Me.rbLeftHead.TabIndex = 0
        Me.rbLeftHead.TabStop = True
        Me.rbLeftHead.Text = "Left"
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(36, 88)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(440, 360)
        Me.TextBox1.TabIndex = 103
        Me.TextBox1.Text = "Info:                                                                 First, make" & _
        " sure the syringe is not tighten up. Press and hold Control Key on Keyboard and " & _
        "use the mouse's scroll ball to move the syringe's needle to the hardware referen" & _
        "ce point. Alternatively, click the Move To Current Reference Point button to mov" & _
        "e the needle to the previously saved reference point position. After that, move " & _
        "down the z stage before touching the reference point. Tighten the syringe and co" & _
        "ntinue with the fine adjustment to make sure the needle tip is aligned to refere" & _
        "nce point. Do not forget to save the reference point position. "
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label4.Location = New System.Drawing.Point(140, 544)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 24)
        Me.Label4.TabIndex = 97
        Me.Label4.Text = "Reference X : "
        '
        'ZPosition
        '
        Me.ZPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ZPosition.Location = New System.Drawing.Point(284, 608)
        Me.ZPosition.Name = "ZPosition"
        Me.ZPosition.Size = New System.Drawing.Size(88, 24)
        Me.ZPosition.TabIndex = 102
        Me.ZPosition.Text = "0.000"
        '
        'YPosition
        '
        Me.YPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.YPosition.Location = New System.Drawing.Point(284, 576)
        Me.YPosition.Name = "YPosition"
        Me.YPosition.Size = New System.Drawing.Size(88, 24)
        Me.YPosition.TabIndex = 101
        Me.YPosition.Text = "0.000"
        '
        'XPosition
        '
        Me.XPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.XPosition.Location = New System.Drawing.Point(284, 544)
        Me.XPosition.Name = "XPosition"
        Me.XPosition.Size = New System.Drawing.Size(88, 24)
        Me.XPosition.TabIndex = 100
        Me.XPosition.Text = "0.000"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(140, 608)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 24)
        Me.Label7.TabIndex = 99
        Me.Label7.Text = "Reference Z :"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(140, 576)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 24)
        Me.Label6.TabIndex = 98
        Me.Label6.Text = "Reference Y :"
        '
        'ButtonSetCalibratorPosition
        '
        Me.ButtonSetCalibratorPosition.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSetCalibratorPosition.Location = New System.Drawing.Point(284, 680)
        Me.ButtonSetCalibratorPosition.Name = "ButtonSetCalibratorPosition"
        Me.ButtonSetCalibratorPosition.Size = New System.Drawing.Size(192, 56)
        Me.ButtonSetCalibratorPosition.TabIndex = 62
        Me.ButtonSetCalibratorPosition.Text = "Save Reference Position"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(8, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(336, 32)
        Me.Label9.TabIndex = 46
        Me.Label9.Text = "Needle Reference Point Setup"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(416, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 45
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonExit.Visible = False
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(36, 680)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(192, 57)
        Me.Button1.TabIndex = 62
        Me.Button1.Text = "Move To Current Reference Point"
        '
        'NeedleCalibrationSetup1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "NeedleCalibrationSetup1"
        Me.Text = "NeedleCalibrationSetup1"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ButtonSetCalibratorPosition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSetCalibratorPosition.Click
        XPosition.Text = m_Tri.XPosition()
        YPosition.Text = m_Tri.YPosition()
        ZPosition.Text = m_Tri.ZPosition()
        If rbLeftHead.Checked Then
            IDS.Data.Hardware.Needle.Left.CalibratorPos.X = m_Tri.XPosition()
            IDS.Data.Hardware.Needle.Left.CalibratorPos.Y = m_Tri.YPosition()
            IDS.Data.Hardware.Needle.Left.CalibratorPos.Z = m_Tri.ZPosition()
        Else
            IDS.Data.Hardware.Needle.Right.CalibratorPos.X = m_Tri.XPosition()
            IDS.Data.Hardware.Needle.Right.CalibratorPos.Y = m_Tri.YPosition()
            IDS.Data.Hardware.Needle.Right.CalibratorPos.Z = m_Tri.ZPosition()
        End If
        IDS.Data.SaveGlobalData()
        m_Tri.Move_Z(0)
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        If Not m_Tri.Move_Z(0) Then Exit Sub
        RemovePanel(CurrentControl)
    End Sub

    Public Sub Revert()
        IDS.Data.OpenData()
        XPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.X
        YPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
        ZPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Z
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
            pos(0) = IDS.Data.Hardware.Needle.Left.CalibratorPos.X
            pos(1) = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
        Else
            pos(0) = IDS.Data.Hardware.Needle.Right.CalibratorPos.X
            pos(1) = IDS.Data.Hardware.Needle.Right.CalibratorPos.Y
        End If

        If Not m_Tri.Move_Z(0) Then Exit Sub
        If Not m_Tri.Move_XY(pos) Then Exit Sub
        If Not m_Tri.Move_Z(IDS.Data.Hardware.Needle.Left.CalibratorPos.Z) Then Exit Sub
    End Sub

    Private Sub NeedleCalibrationSetup1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub rbLeftHead_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLeftHead.CheckedChanged
        If rbLeftHead.Checked Then
            XPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.X
            YPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Y
            ZPosition.Text = IDS.Data.Hardware.Needle.Left.CalibratorPos.Z
        Else
            XPosition.Text = IDS.Data.Hardware.Needle.Right.CalibratorPos.X
            YPosition.Text = IDS.Data.Hardware.Needle.Right.CalibratorPos.Y
            ZPosition.Text = IDS.Data.Hardware.Needle.Right.CalibratorPos.Z
        End If
    End Sub
End Class
