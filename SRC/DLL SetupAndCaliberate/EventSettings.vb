Imports DLL_Export_Service

Public Class EventSettings
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
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Lbl_PromptHeight As System.Windows.Forms.Label
    Friend WithEvents Lbl_SkipHeight As System.Windows.Forms.Label
    Friend WithEvents Lbl_PromptFiducial As System.Windows.Forms.Label
    Friend WithEvents Lbl_SkipFiducial As System.Windows.Forms.Label
    Friend WithEvents Lbl_PromptChip As System.Windows.Forms.Label
    Friend WithEvents Lbl_SkipChip As System.Windows.Forms.Label
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents RB_HeightPrompt As System.Windows.Forms.RadioButton
    Friend WithEvents RB_HeightSkip As System.Windows.Forms.RadioButton
    Friend WithEvents Rb_FidPrompt As System.Windows.Forms.RadioButton
    Friend WithEvents RB_FidSkip As System.Windows.Forms.RadioButton
    Friend WithEvents RB_ChipPrompt As System.Windows.Forms.RadioButton
    Friend WithEvents RB_ChipSkip As System.Windows.Forms.RadioButton
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents LBL_SkipQC As System.Windows.Forms.Label
    Friend WithEvents Lbl_PromptQC As System.Windows.Forms.Label
    Friend WithEvents RB_QCPrompt As System.Windows.Forms.RadioButton
    Friend WithEvents RB_QCSkip As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(EventSettings))
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Panel7 = New System.Windows.Forms.Panel
        Me.Lbl_SkipHeight = New System.Windows.Forms.Label
        Me.Lbl_PromptHeight = New System.Windows.Forms.Label
        Me.RB_HeightPrompt = New System.Windows.Forms.RadioButton
        Me.RB_HeightSkip = New System.Windows.Forms.RadioButton
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Rb_FidPrompt = New System.Windows.Forms.RadioButton
        Me.RB_FidSkip = New System.Windows.Forms.RadioButton
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Lbl_SkipFiducial = New System.Windows.Forms.Label
        Me.Lbl_PromptFiducial = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Panel8 = New System.Windows.Forms.Panel
        Me.Lbl_SkipChip = New System.Windows.Forms.Label
        Me.Lbl_PromptChip = New System.Windows.Forms.Label
        Me.RB_ChipPrompt = New System.Windows.Forms.RadioButton
        Me.RB_ChipSkip = New System.Windows.Forms.RadioButton
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.Panel9 = New System.Windows.Forms.Panel
        Me.LBL_SkipQC = New System.Windows.Forms.Label
        Me.Lbl_PromptQC = New System.Windows.Forms.Label
        Me.RB_QCPrompt = New System.Windows.Forms.RadioButton
        Me.RB_QCSkip = New System.Windows.Forms.RadioButton
        Me.GroupBox2.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Panel7)
        Me.GroupBox2.Controls.Add(Me.RB_HeightPrompt)
        Me.GroupBox2.Controls.Add(Me.RB_HeightSkip)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(48, 256)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(408, 152)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Height Sense Failed Action Handling"
        '
        'Panel7
        '
        Me.Panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel7.Controls.Add(Me.Lbl_SkipHeight)
        Me.Panel7.Controls.Add(Me.Lbl_PromptHeight)
        Me.Panel7.Location = New System.Drawing.Point(16, 32)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(376, 66)
        Me.Panel7.TabIndex = 4
        '
        'Lbl_SkipHeight
        '
        Me.Lbl_SkipHeight.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_SkipHeight.Name = "Lbl_SkipHeight"
        Me.Lbl_SkipHeight.Size = New System.Drawing.Size(358, 46)
        Me.Lbl_SkipHeight.TabIndex = 0
        Me.Lbl_SkipHeight.Text = "Skip only the part associated with the height point and continue to the next part" & _
        ""
        '
        'Lbl_PromptHeight
        '
        Me.Lbl_PromptHeight.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_PromptHeight.Name = "Lbl_PromptHeight"
        Me.Lbl_PromptHeight.Size = New System.Drawing.Size(358, 46)
        Me.Lbl_PromptHeight.TabIndex = 1
        Me.Lbl_PromptHeight.Text = "Pop up message box for operator to choose Skip or Stop the production process"
        '
        'RB_HeightPrompt
        '
        Me.RB_HeightPrompt.Location = New System.Drawing.Point(184, 112)
        Me.RB_HeightPrompt.Name = "RB_HeightPrompt"
        Me.RB_HeightPrompt.Size = New System.Drawing.Size(168, 24)
        Me.RB_HeightPrompt.TabIndex = 1
        Me.RB_HeightPrompt.Text = "Prompt Message"
        '
        'RB_HeightSkip
        '
        Me.RB_HeightSkip.Checked = True
        Me.RB_HeightSkip.Location = New System.Drawing.Point(16, 112)
        Me.RB_HeightSkip.Name = "RB_HeightSkip"
        Me.RB_HeightSkip.TabIndex = 0
        Me.RB_HeightSkip.TabStop = True
        Me.RB_HeightSkip.Text = "Skip"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Rb_FidPrompt)
        Me.GroupBox1.Controls.Add(Me.RB_FidSkip)
        Me.GroupBox1.Controls.Add(Me.Panel6)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(48, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(407, 152)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Fiducial Failed Action Handling"
        '
        'Rb_FidPrompt
        '
        Me.Rb_FidPrompt.Location = New System.Drawing.Point(184, 112)
        Me.Rb_FidPrompt.Name = "Rb_FidPrompt"
        Me.Rb_FidPrompt.Size = New System.Drawing.Size(184, 24)
        Me.Rb_FidPrompt.TabIndex = 1
        Me.Rb_FidPrompt.Text = "Prompt Message"
        '
        'RB_FidSkip
        '
        Me.RB_FidSkip.Checked = True
        Me.RB_FidSkip.Location = New System.Drawing.Point(16, 112)
        Me.RB_FidSkip.Name = "RB_FidSkip"
        Me.RB_FidSkip.TabIndex = 0
        Me.RB_FidSkip.TabStop = True
        Me.RB_FidSkip.Text = "Skip"
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel6.Controls.Add(Me.Lbl_SkipFiducial)
        Me.Panel6.Controls.Add(Me.Lbl_PromptFiducial)
        Me.Panel6.Location = New System.Drawing.Point(16, 32)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(376, 66)
        Me.Panel6.TabIndex = 3
        '
        'Lbl_SkipFiducial
        '
        Me.Lbl_SkipFiducial.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_SkipFiducial.Name = "Lbl_SkipFiducial"
        Me.Lbl_SkipFiducial.Size = New System.Drawing.Size(358, 46)
        Me.Lbl_SkipFiducial.TabIndex = 0
        Me.Lbl_SkipFiducial.Text = "Skip only the part associated with the fiducial point and continue to the next pa" & _
        "rt"
        '
        'Lbl_PromptFiducial
        '
        Me.Lbl_PromptFiducial.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_PromptFiducial.Name = "Lbl_PromptFiducial"
        Me.Lbl_PromptFiducial.Size = New System.Drawing.Size(350, 46)
        Me.Lbl_PromptFiducial.TabIndex = 1
        Me.Lbl_PromptFiducial.Text = "Pop up message box for operator to choose Skip or Stop the production process"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Panel8)
        Me.GroupBox3.Controls.Add(Me.RB_ChipPrompt)
        Me.GroupBox3.Controls.Add(Me.RB_ChipSkip)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(48, 448)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(408, 152)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Chip Recognition Failed Action Handling"
        '
        'Panel8
        '
        Me.Panel8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel8.Controls.Add(Me.Lbl_SkipChip)
        Me.Panel8.Controls.Add(Me.Lbl_PromptChip)
        Me.Panel8.Location = New System.Drawing.Point(16, 32)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(376, 66)
        Me.Panel8.TabIndex = 4
        '
        'Lbl_SkipChip
        '
        Me.Lbl_SkipChip.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_SkipChip.Name = "Lbl_SkipChip"
        Me.Lbl_SkipChip.Size = New System.Drawing.Size(358, 46)
        Me.Lbl_SkipChip.TabIndex = 0
        Me.Lbl_SkipChip.Text = "Skip only the chip checked and continue to the next chip"
        '
        'Lbl_PromptChip
        '
        Me.Lbl_PromptChip.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_PromptChip.Name = "Lbl_PromptChip"
        Me.Lbl_PromptChip.Size = New System.Drawing.Size(358, 46)
        Me.Lbl_PromptChip.TabIndex = 1
        Me.Lbl_PromptChip.Text = "Pop up message box for operator to choose Skip chip or Skip Part"
        '
        'RB_ChipPrompt
        '
        Me.RB_ChipPrompt.Location = New System.Drawing.Point(184, 112)
        Me.RB_ChipPrompt.Name = "RB_ChipPrompt"
        Me.RB_ChipPrompt.Size = New System.Drawing.Size(184, 24)
        Me.RB_ChipPrompt.TabIndex = 1
        Me.RB_ChipPrompt.Text = "Prompt Message"
        '
        'RB_ChipSkip
        '
        Me.RB_ChipSkip.Checked = True
        Me.RB_ChipSkip.Location = New System.Drawing.Point(16, 112)
        Me.RB_ChipSkip.Name = "RB_ChipSkip"
        Me.RB_ChipSkip.TabIndex = 0
        Me.RB_ChipSkip.TabStop = True
        Me.RB_ChipSkip.Text = "Skip"
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox6)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox1)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox3)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(144, -48)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 13
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 50
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(392, 880)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(88, 48)
        Me.ButtonRevert.TabIndex = 36
        Me.ButtonRevert.Text = "Revert"
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(280, 880)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(88, 48)
        Me.ButtonSave.TabIndex = 35
        Me.ButtonSave.Text = "Save"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(232, 32)
        Me.Label17.TabIndex = 34
        Me.Label17.Text = "Event Failed Actions"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Panel9)
        Me.GroupBox6.Controls.Add(Me.RB_QCPrompt)
        Me.GroupBox6.Controls.Add(Me.RB_QCSkip)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(52, 648)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(408, 152)
        Me.GroupBox6.TabIndex = 51
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Dot QC Failed Action Handling"
        '
        'Panel9
        '
        Me.Panel9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel9.Controls.Add(Me.LBL_SkipQC)
        Me.Panel9.Controls.Add(Me.Lbl_PromptQC)
        Me.Panel9.Location = New System.Drawing.Point(16, 32)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(376, 66)
        Me.Panel9.TabIndex = 5
        '
        'LBL_SkipQC
        '
        Me.LBL_SkipQC.Location = New System.Drawing.Point(10, 10)
        Me.LBL_SkipQC.Name = "LBL_SkipQC"
        Me.LBL_SkipQC.Size = New System.Drawing.Size(358, 54)
        Me.LBL_SkipQC.TabIndex = 0
        Me.LBL_SkipQC.Text = "Skip only the part associated with the dot QC and continue to the next part"
        '
        'Lbl_PromptQC
        '
        Me.Lbl_PromptQC.Location = New System.Drawing.Point(10, 10)
        Me.Lbl_PromptQC.Name = "Lbl_PromptQC"
        Me.Lbl_PromptQC.Size = New System.Drawing.Size(366, 46)
        Me.Lbl_PromptQC.TabIndex = 1
        Me.Lbl_PromptQC.Text = "Pop up message box for operator to choose Skip or Stop the production process"
        '
        'RB_QCPrompt
        '
        Me.RB_QCPrompt.Location = New System.Drawing.Point(184, 112)
        Me.RB_QCPrompt.Name = "RB_QCPrompt"
        Me.RB_QCPrompt.Size = New System.Drawing.Size(160, 24)
        Me.RB_QCPrompt.TabIndex = 1
        Me.RB_QCPrompt.Text = "Prompt Message"
        '
        'RB_QCSkip
        '
        Me.RB_QCSkip.Checked = True
        Me.RB_QCSkip.Location = New System.Drawing.Point(16, 112)
        Me.RB_QCSkip.Name = "RB_QCSkip"
        Me.RB_QCSkip.TabIndex = 0
        Me.RB_QCSkip.TabStop = True
        Me.RB_QCSkip.Text = "Skip"
        '
        'EventSettings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(768, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "EventSettings"
        Me.Text = "Event_Failed_Actions"
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.Panel9.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private Sub RB_FidSkip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB_FidSkip.CheckedChanged
        If RB_FidSkip.Checked Then
            Lbl_SkipFiducial.BringToFront()
            'FIducialStatus = 1/0 for example - send to shen jian
        Else
            Lbl_PromptFiducial.BringToFront()
        End If
    End Sub

    Private Sub RB_HeightSkip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB_HeightSkip.CheckedChanged
        If RB_HeightSkip.Checked Then
            Lbl_SkipHeight.BringToFront()
        Else
            Lbl_PromptHeight.BringToFront()
        End If
    End Sub

    Private Sub RB_ChipSkip_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB_ChipSkip.CheckedChanged
        If RB_ChipSkip.Checked Then
            Lbl_SkipChip.BringToFront()
        Else
            Lbl_PromptChip.BringToFront()
        End If
    End Sub


    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        IDS.Data.Hardware.SPC.ChipFailedAction = RB_ChipPrompt.Checked    'true = prompt (manual), false = autoskip
        IDS.Data.Hardware.SPC.HeightFailedAction = RB_HeightPrompt.Checked
        IDS.Data.Hardware.SPC.FiducialFailedAction = Rb_FidPrompt.Checked
        IDS.Data.Hardware.SPC.QCFailedAction = RB_QCPrompt.Checked
        IDS.Data.SaveLocalData()
    End Sub

    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        RevertData()
    End Sub

    Public Sub RevertData()
        IDS.Data.OpenData()
        RB_ChipPrompt.Checked = IDS.Data.Hardware.SPC.ChipFailedAction
        RB_HeightPrompt.Checked = IDS.Data.Hardware.SPC.HeightFailedAction
        Rb_FidPrompt.Checked = IDS.Data.Hardware.SPC.FiducialFailedAction
        RB_QCPrompt.Checked = IDS.Data.Hardware.SPC.QCFailedAction
        RB_ChipSkip.Checked = Not IDS.Data.Hardware.SPC.ChipFailedAction
        RB_HeightSkip.Checked = Not IDS.Data.Hardware.SPC.HeightFailedAction
        RB_FidSkip.Checked = Not IDS.Data.Hardware.SPC.FiducialFailedAction
        RB_QCSkip.Checked = Not IDS.Data.Hardware.SPC.QCFailedAction
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub
End Class
