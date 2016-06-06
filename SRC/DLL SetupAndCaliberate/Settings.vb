Imports DLL_Export_Service

Public Class Settings
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
    Public WithEvents PanelLeft As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ButtonStationPositions As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonEventHandling As System.Windows.Forms.Button
    Friend WithEvents ButtonSPCLogging As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenserSettings As System.Windows.Forms.Button
    Public WithEvents ButtonVolumeCalibSettings As System.Windows.Forms.Button
    Public WithEvents ButtonThermalSettings As System.Windows.Forms.Button
    Public WithEvents panelRight As System.Windows.Forms.Panel
    Friend WithEvents gbXYZ As System.Windows.Forms.GroupBox
    Friend WithEvents gbProgramSettings As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.panelRight = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.ButtonStationPositions = New System.Windows.Forms.Button
        Me.ButtonVolumeCalibSettings = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSettings = New System.Windows.Forms.Button
        Me.ButtonEventHandling = New System.Windows.Forms.Button
        Me.ButtonThermalSettings = New System.Windows.Forms.Button
        Me.ButtonSPCLogging = New System.Windows.Forms.Button
        Me.ButtonDispenserSettings = New System.Windows.Forms.Button
        Me.PanelLeft = New System.Windows.Forms.Panel
        Me.gbXYZ = New System.Windows.Forms.GroupBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.gbProgramSettings = New System.Windows.Forms.GroupBox
        Me.Panel3.SuspendLayout()
        Me.PanelLeft.SuspendLayout()
        Me.gbProgramSettings.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelRight
        '
        Me.panelRight.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.panelRight.Location = New System.Drawing.Point(768, 24)
        Me.panelRight.Name = "panelRight"
        Me.panelRight.Size = New System.Drawing.Size(512, 928)
        Me.panelRight.TabIndex = 32
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.gbProgramSettings)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel3.Location = New System.Drawing.Point(16, 8)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(744, 328)
        Me.Panel3.TabIndex = 65
        '
        'ButtonStationPositions
        '
        Me.ButtonStationPositions.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStationPositions.Location = New System.Drawing.Point(376, 40)
        Me.ButtonStationPositions.Name = "ButtonStationPositions"
        Me.ButtonStationPositions.Size = New System.Drawing.Size(232, 56)
        Me.ButtonStationPositions.TabIndex = 43
        Me.ButtonStationPositions.Text = "Gantry Settings"
        '
        'ButtonVolumeCalibSettings
        '
        Me.ButtonVolumeCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolumeCalibSettings.Location = New System.Drawing.Point(440, 280)
        Me.ButtonVolumeCalibSettings.Name = "ButtonVolumeCalibSettings"
        Me.ButtonVolumeCalibSettings.Size = New System.Drawing.Size(168, 24)
        Me.ButtonVolumeCalibSettings.TabIndex = 46
        Me.ButtonVolumeCalibSettings.Text = "Volume Calibration"
        Me.ButtonVolumeCalibSettings.Visible = False
        '
        'ButtonNeedleCalibSettings
        '
        Me.ButtonNeedleCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSettings.Location = New System.Drawing.Point(112, 128)
        Me.ButtonNeedleCalibSettings.Name = "ButtonNeedleCalibSettings"
        Me.ButtonNeedleCalibSettings.Size = New System.Drawing.Size(232, 56)
        Me.ButtonNeedleCalibSettings.TabIndex = 44
        Me.ButtonNeedleCalibSettings.Text = "Needle Calibration"
        '
        'ButtonEventHandling
        '
        Me.ButtonEventHandling.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEventHandling.Location = New System.Drawing.Point(376, 128)
        Me.ButtonEventHandling.Name = "ButtonEventHandling"
        Me.ButtonEventHandling.Size = New System.Drawing.Size(232, 56)
        Me.ButtonEventHandling.TabIndex = 47
        Me.ButtonEventHandling.Text = "Event Handling"
        '
        'ButtonThermalSettings
        '
        Me.ButtonThermalSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonThermalSettings.Location = New System.Drawing.Point(632, 272)
        Me.ButtonThermalSettings.Name = "ButtonThermalSettings"
        Me.ButtonThermalSettings.Size = New System.Drawing.Size(80, 30)
        Me.ButtonThermalSettings.TabIndex = 46
        Me.ButtonThermalSettings.Text = "Heater"
        Me.ButtonThermalSettings.Visible = False
        '
        'ButtonSPCLogging
        '
        Me.ButtonSPCLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCLogging.Location = New System.Drawing.Point(112, 216)
        Me.ButtonSPCLogging.Name = "ButtonSPCLogging"
        Me.ButtonSPCLogging.Size = New System.Drawing.Size(232, 56)
        Me.ButtonSPCLogging.TabIndex = 51
        Me.ButtonSPCLogging.Text = "SPC Logging"
        '
        'ButtonDispenserSettings
        '
        Me.ButtonDispenserSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenserSettings.Location = New System.Drawing.Point(112, 40)
        Me.ButtonDispenserSettings.Name = "ButtonDispenserSettings"
        Me.ButtonDispenserSettings.Size = New System.Drawing.Size(232, 56)
        Me.ButtonDispenserSettings.TabIndex = 42
        Me.ButtonDispenserSettings.Text = "Dispenser"
        '
        'PanelLeft
        '
        Me.PanelLeft.BackColor = System.Drawing.SystemColors.Control
        Me.PanelLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelLeft.Controls.Add(Me.gbXYZ)
        Me.PanelLeft.Controls.Add(Me.RichTextBox1)
        Me.PanelLeft.Controls.Add(Me.Panel3)
        Me.PanelLeft.Location = New System.Drawing.Point(0, 24)
        Me.PanelLeft.Name = "PanelLeft"
        Me.PanelLeft.Size = New System.Drawing.Size(768, 928)
        Me.PanelLeft.TabIndex = 33
        '
        'gbXYZ
        '
        Me.gbXYZ.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbXYZ.Location = New System.Drawing.Point(127, 408)
        Me.gbXYZ.Name = "gbXYZ"
        Me.gbXYZ.Size = New System.Drawing.Size(512, 376)
        Me.gbXYZ.TabIndex = 66
        Me.gbXYZ.TabStop = False
        Me.gbXYZ.Text = "XYZ Stages:"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.DarkBlue
        Me.RichTextBox1.Location = New System.Drawing.Point(16, 440)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(72, 40)
        Me.RichTextBox1.TabIndex = 19
        Me.RichTextBox1.Text = "It is highly recommended to go through the entire process setup for a new pattern" & _
        "." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "For an existing pattern, the process setup can also be modified."
        Me.RichTextBox1.Visible = False
        '
        'gbProgramSettings
        '
        Me.gbProgramSettings.Controls.Add(Me.ButtonVolumeCalibSettings)
        Me.gbProgramSettings.Controls.Add(Me.ButtonNeedleCalibSettings)
        Me.gbProgramSettings.Controls.Add(Me.ButtonEventHandling)
        Me.gbProgramSettings.Controls.Add(Me.ButtonThermalSettings)
        Me.gbProgramSettings.Controls.Add(Me.ButtonSPCLogging)
        Me.gbProgramSettings.Controls.Add(Me.ButtonDispenserSettings)
        Me.gbProgramSettings.Controls.Add(Me.ButtonStationPositions)
        Me.gbProgramSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProgramSettings.Location = New System.Drawing.Point(16, 8)
        Me.gbProgramSettings.Name = "gbProgramSettings"
        Me.gbProgramSettings.Size = New System.Drawing.Size(720, 312)
        Me.gbProgramSettings.TabIndex = 52
        Me.gbProgramSettings.TabStop = False
        Me.gbProgramSettings.Text = "Program Settings:"
        '
        'Settings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1288, 958)
        Me.Controls.Add(Me.panelRight)
        Me.Controls.Add(Me.PanelLeft)
        Me.Name = "Settings"
        Me.Text = "Basic Setup"
        Me.Panel3.ResumeLayout(False)
        Me.PanelLeft.ResumeLayout(False)
        Me.gbProgramSettings.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CurrentPanel As System.Windows.Forms.Panel

    Public Sub RevertData()

        With IDS.Data.Hardware.Needle
            offset_x = .Left.NeedleCalibrationPosition.X - .Left.CalibratorPos.X
            offset_y = .Left.NeedleCalibrationPosition.Y - .Left.CalibratorPos.Y
            offset_z = .Left.NeedleCalibrationPosition.Z
        End With

    End Sub

    'for programming mode to call when toggling in between program editor/basic setup when a different file has been opened/created
    Public Sub RemoveCurrentPanel()
        RemovePanel(CurrentPanel)
    End Sub

    Private Sub ButtonNeedleCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNeedleCalibSettings.Click
        OnLaser()
        MyNeedleCalibrationSettings.Revert()
        CurrentPanel = MyNeedleCalibrationSettings.PanelToBeAdded
        AddPanel(PanelRight, MyNeedleCalibrationSettings.PanelToBeAdded)
    End Sub

    Private Sub ButtonStationPositions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStationPositions.Click
        CurrentPanel = MyGantrySettings.PanelToBeAdded
        AddPanel(PanelRight, MyGantrySettings.PanelToBeAdded)
        'initialize
        MyGantrySettings.StationPosition.SelectedIndex = 0
        MyGantrySettings.LeftHead.Checked = True
        MyGantrySettings.RevertData()
    End Sub

    Private Sub ButtonDispenserSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDispenserSettings.Click
        CurrentPanel = MyDispenserSettings.PanelToBeAdded
        AddPanel(PanelRight, MyDispenserSettings.PanelToBeAdded)
        MyDispenserSettings.HeadType.SelectedIndex = 0
        MyDispenserSettings.RevertData()
    End Sub

    Private Sub ButtonSPCLogging_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPCLogging.Click
        CurrentPanel = MySPCLogging.PanelToBeAdded
        AddPanel(PanelRight, MySPCLogging.PanelToBeAdded)
        MySPCLogging.RevertData()
    End Sub

    Private Sub ButtonEventHandling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEventHandling.Click
        CurrentPanel = MyEventSettings.PanelToBeAdded
        AddPanel(PanelRight, MyEventSettings.PanelToBeAdded)
        MyEventSettings.RevertData()
    End Sub

    Private Sub ButtonThermalSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonThermalSettings.Click
        CurrentPanel = MyHeaterSettings.PanelToBeAdded
        AddPanel(PanelRight, MyHeaterSettings.PanelToBeAdded)
        MyHeaterSettings.RevertData()
    End Sub

    Private Sub ButtonVolumeCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolumeCalibSettings.Click
        CurrentPanel = MyVolumeCalibrationSettings.PanelToBeAdded
        AddPanel(PanelRight, MyVolumeCalibrationSettings.PanelToBeAdded)
        MyVolumeCalibrationSettings.RevertData()
    End Sub

End Class