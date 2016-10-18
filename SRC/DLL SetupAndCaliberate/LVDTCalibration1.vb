Imports DLL_Export_Service


Public Class LVDTCalibration1
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
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonNext As System.Windows.Forms.Button
    Friend WithEvents ButtonSetZ As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents LVDTOffsetXPos As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents LVDTOffsetYPos As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents LVDTReferenceZPos As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents NewLVDTZReference As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButtonLVDTOK As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents LVDT_Down As System.Windows.Forms.Button
    Friend WithEvents LVDT_Up As System.Windows.Forms.Button
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(LVDTCalibration1))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.LVDTOffsetXPos = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.LVDTOffsetYPos = New System.Windows.Forms.Label
        Me.LVDTReferenceZPos = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LVDT_Up = New System.Windows.Forms.Button
        Me.LVDT_Down = New System.Windows.Forms.Button
        Me.ButtonLVDTOK = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.NewLVDTZReference = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonNext = New System.Windows.Forms.Button
        Me.ButtonSetZ = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.Controls.Add(Me.Label5)
        Me.PanelToBeAdded.Controls.Add(Me.Panel2)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(64, 0)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 34
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label5.Location = New System.Drawing.Point(0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(168, 32)
        Me.Label5.TabIndex = 41
        Me.Label5.Text = "LVDT Setup"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.Label10)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.LVDTOffsetXPos)
        Me.Panel2.Controls.Add(Me.Label14)
        Me.Panel2.Controls.Add(Me.Label13)
        Me.Panel2.Controls.Add(Me.LVDTOffsetYPos)
        Me.Panel2.Controls.Add(Me.LVDTReferenceZPos)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Location = New System.Drawing.Point(0, 80)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(512, 856)
        Me.Panel2.TabIndex = 40
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(272, 32)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(16, 23)
        Me.Label10.TabIndex = 85
        Me.Label10.Text = ","
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(368, 32)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 23)
        Me.Label4.TabIndex = 88
        Me.Label4.Text = ","
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(464, 32)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(16, 23)
        Me.Label8.TabIndex = 87
        Me.Label8.Text = ")"
        '
        'LVDTOffsetXPos
        '
        Me.LVDTOffsetXPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LVDTOffsetXPos.Location = New System.Drawing.Point(200, 32)
        Me.LVDTOffsetXPos.Name = "LVDTOffsetXPos"
        Me.LVDTOffsetXPos.Size = New System.Drawing.Size(72, 23)
        Me.LVDTOffsetXPos.TabIndex = 81
        Me.LVDTOffsetXPos.Text = "100.000"
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(184, 32)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(16, 23)
        Me.Label14.TabIndex = 82
        Me.Label14.Text = "("
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(256, 32)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(16, 23)
        Me.Label13.TabIndex = 83
        Me.Label13.Text = ","
        '
        'LVDTOffsetYPos
        '
        Me.LVDTOffsetYPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LVDTOffsetYPos.Location = New System.Drawing.Point(296, 32)
        Me.LVDTOffsetYPos.Name = "LVDTOffsetYPos"
        Me.LVDTOffsetYPos.Size = New System.Drawing.Size(72, 23)
        Me.LVDTOffsetYPos.TabIndex = 84
        Me.LVDTOffsetYPos.Text = "100.000"
        '
        'LVDTReferenceZPos
        '
        Me.LVDTReferenceZPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LVDTReferenceZPos.Location = New System.Drawing.Point(392, 32)
        Me.LVDTReferenceZPos.Name = "LVDTReferenceZPos"
        Me.LVDTReferenceZPos.Size = New System.Drawing.Size(72, 23)
        Me.LVDTReferenceZPos.TabIndex = 86
        Me.LVDTReferenceZPos.Text = "100.000"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LVDT_Up)
        Me.GroupBox1.Controls.Add(Me.LVDT_Down)
        Me.GroupBox1.Controls.Add(Me.ButtonLVDTOK)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.NewLVDTZReference)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.ButtonRevert)
        Me.GroupBox1.Controls.Add(Me.ButtonNext)
        Me.GroupBox1.Controls.Add(Me.ButtonSetZ)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 80)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 760)
        Me.GroupBox1.TabIndex = 37
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Z Reference Level"
        '
        'LVDT_Up
        '
        Me.LVDT_Up.Location = New System.Drawing.Point(264, 584)
        Me.LVDT_Up.Name = "LVDT_Up"
        Me.LVDT_Up.Size = New System.Drawing.Size(100, 50)
        Me.LVDT_Up.TabIndex = 22
        Me.LVDT_Up.Text = "LVDT    Up"
        '
        'LVDT_Down
        '
        Me.LVDT_Down.Location = New System.Drawing.Point(136, 584)
        Me.LVDT_Down.Name = "LVDT_Down"
        Me.LVDT_Down.Size = New System.Drawing.Size(100, 50)
        Me.LVDT_Down.TabIndex = 21
        Me.LVDT_Down.Text = "LVDT Down"
        '
        'ButtonLVDTOK
        '
        Me.ButtonLVDTOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonLVDTOK.Location = New System.Drawing.Point(200, 440)
        Me.ButtonLVDTOK.Name = "ButtonLVDTOK"
        Me.ButtonLVDTOK.Size = New System.Drawing.Size(100, 50)
        Me.ButtonLVDTOK.TabIndex = 20
        Me.ButtonLVDTOK.Text = "Set X,Y"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label6.Location = New System.Drawing.Point(32, 360)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(416, 48)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "3. Move LVDT sensing point to align to the center of Offset Calibrator "
        '
        'Label7
        '
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label7.Location = New System.Drawing.Point(96, 208)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(168, 24)
        Me.Label7.TabIndex = 17
        Me.Label7.Text = "Current Z Reference: "
        '
        'NewLVDTZReference
        '
        Me.NewLVDTZReference.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.NewLVDTZReference.Location = New System.Drawing.Point(280, 208)
        Me.NewLVDTZReference.Name = "NewLVDTZReference"
        Me.NewLVDTZReference.Size = New System.Drawing.Size(96, 24)
        Me.NewLVDTZReference.TabIndex = 18
        Me.NewLVDTZReference.Text = "0.000"
        '
        'Label2
        '
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label2.Location = New System.Drawing.Point(32, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(456, 48)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "2. Move down LVDT to touch the Height Reference Block. Click Set Z Button to set " & _
        "the Z Reference Level"
        '
        'Label3
        '
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label3.Location = New System.Drawing.Point(32, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(424, 40)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "1. Jog LVDT to the approximate center of Height Reference Block"
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(32, 680)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(100, 50)
        Me.ButtonRevert.TabIndex = 7
        Me.ButtonRevert.Text = "Cancel"
        '
        'ButtonNext
        '
        Me.ButtonNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNext.Location = New System.Drawing.Point(368, 680)
        Me.ButtonNext.Name = "ButtonNext"
        Me.ButtonNext.Size = New System.Drawing.Size(100, 50)
        Me.ButtonNext.TabIndex = 6
        Me.ButtonNext.Text = "Next"
        '
        'ButtonSetZ
        '
        Me.ButtonSetZ.Location = New System.Drawing.Point(200, 264)
        Me.ButtonSetZ.Name = "ButtonSetZ"
        Me.ButtonSetZ.Size = New System.Drawing.Size(100, 50)
        Me.ButtonSetZ.TabIndex = 9
        Me.ButtonSetZ.Text = "Set Z"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 23)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "LVDT Position Offset :"
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(400, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 39
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'LVDTCalibration1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 878)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "LVDTCalibration1"
        Me.Text = "LVDT"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub ButtonNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNext.Click
        GroupBox1.Controls.Add(MyLVDTSetup1.GroupBox1)
        MyLVDTSetup1.GroupBox1.Location = New Point(0, 0)
        MyLVDTSetup1.GroupBox1.BringToFront()
        MyLVDTSetup1.GroupBox1.Show()
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        GroupBox1.Controls.Remove(MyLVDTSetup1.GroupBox1)
        RemovePanel(CurrentControl)
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        NewLVDTZReference.Text = IDS.Data.Hardware.HeightSensor.TP.HeightReference
        ButtonSetZ.Enabled = True
    End Sub
    'xu long 260707 /*
    Private Sub ButtonSetZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSetZ.Click
        MessageBox.Show("Setting Z Position ... ")
        'Display Current Z reference level - get readings from Shen Jian's LVDT module
        'LVDTZReference.Text = Shen Jian's return value

        'Save reading to IDS.Data.Hardware
        Dim TestValue As Integer
        Dim rtn As Integer = m_Tri.GetLVDTRange(32, TestValue)
        'While Not m_tri.m_trictrl.Ain(5, TestValue)
        'End While

        If rtn = 0 Then
            IDS.Data.Hardware.HeightSensor.TP.HeightReference = TestValue
            MyLVDTSetup1.NewLVDTZReferencePos.Text = IDS.Data.Hardware.HeightSensor.TP.HeightReference
        Else
            MessageBox.Show("LVDT out of range.")
        End If
        'ButtonSetZ.Enabled = False
        While Not m_Tri.m_TriCtrl.Op(32, 0)   'to move LVDT up
        End While
        MessageBox.Show("OK.")
    End Sub

    Private Sub LVDT_Down_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LVDT_Down.Click
        While Not m_Tri.m_TriCtrl.Op(32, 1)   'to move LVDT down
        End While
    End Sub

    Private Sub LVDT_Up_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LVDT_Up.Click
        While Not m_Tri.m_TriCtrl.Op(32, 0)   'to move LVDT up
        End While
    End Sub

    Private Sub ButtonLVDTOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonLVDTOK.Click

        m_Tri.GetIDSState()
        IDS.Data.Hardware.HeightSensor.TP.CurrentPos.X = m_Tri.XPosition
        IDS.Data.Hardware.HeightSensor.TP.CurrentPos.Y = m_Tri.YPosition
    End Sub
    'xu long 260707 */
End Class
