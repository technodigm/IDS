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
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Public WithEvents PanelLeft As System.Windows.Forms.Panel
    Public WithEvents PanelRight As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ButtonStationPositions As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonEventHandling As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonSPCLogging As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenserSettings As System.Windows.Forms.Button
    Public WithEvents ButtonVolumeCalibSettings As System.Windows.Forms.Button
    Public WithEvents ButtonThermalSettings As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PanelRight = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label1 = New System.Windows.Forms.Label
        Me.ButtonStationPositions = New System.Windows.Forms.Button
        Me.ButtonVolumeCalibSettings = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSettings = New System.Windows.Forms.Button
        Me.ButtonEventHandling = New System.Windows.Forms.Button
        Me.ButtonConveyorSettings = New System.Windows.Forms.Button
        Me.ButtonThermalSettings = New System.Windows.Forms.Button
        Me.ButtonSPCLogging = New System.Windows.Forms.Button
        Me.ButtonDispenserSettings = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.PanelLeft = New System.Windows.Forms.Panel
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.PanelRight.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.PanelLeft.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelRight
        '
        Me.PanelRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelRight.Controls.Add(Me.Panel3)
        Me.PanelRight.Controls.Add(Me.Label17)
        Me.PanelRight.Location = New System.Drawing.Point(768, 24)
        Me.PanelRight.Name = "PanelRight"
        Me.PanelRight.Size = New System.Drawing.Size(512, 911)
        Me.PanelRight.TabIndex = 32
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.ButtonStationPositions)
        Me.Panel3.Controls.Add(Me.ButtonVolumeCalibSettings)
        Me.Panel3.Controls.Add(Me.ButtonNeedleCalibSettings)
        Me.Panel3.Controls.Add(Me.ButtonEventHandling)
        Me.Panel3.Controls.Add(Me.ButtonConveyorSettings)
        Me.Panel3.Controls.Add(Me.ButtonThermalSettings)
        Me.Panel3.Controls.Add(Me.ButtonSPCLogging)
        Me.Panel3.Controls.Add(Me.ButtonDispenserSettings)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel3.Location = New System.Drawing.Point(0, 440)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(512, 328)
        Me.Panel3.TabIndex = 65
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(0, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(192, 32)
        Me.Label1.TabIndex = 17
        Me.Label1.Text = "Program Settings"
        '
        'ButtonStationPositions
        '
        Me.ButtonStationPositions.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStationPositions.Location = New System.Drawing.Point(264, 48)
        Me.ButtonStationPositions.Name = "ButtonStationPositions"
        Me.ButtonStationPositions.Size = New System.Drawing.Size(224, 30)
        Me.ButtonStationPositions.TabIndex = 43
        Me.ButtonStationPositions.Text = "Gantry Settings"
        '
        'ButtonVolumeCalibSettings
        '
        Me.ButtonVolumeCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolumeCalibSettings.Location = New System.Drawing.Point(264, 104)
        Me.ButtonVolumeCalibSettings.Name = "ButtonVolumeCalibSettings"
        Me.ButtonVolumeCalibSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonVolumeCalibSettings.TabIndex = 46
        Me.ButtonVolumeCalibSettings.Text = "Volume Calibration"
        '
        'ButtonNeedleCalibSettings
        '
        Me.ButtonNeedleCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSettings.Location = New System.Drawing.Point(24, 104)
        Me.ButtonNeedleCalibSettings.Name = "ButtonNeedleCalibSettings"
        Me.ButtonNeedleCalibSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonNeedleCalibSettings.TabIndex = 44
        Me.ButtonNeedleCalibSettings.Text = "Needle Calibration"
        '
        'ButtonEventHandling
        '
        Me.ButtonEventHandling.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEventHandling.Location = New System.Drawing.Point(264, 160)
        Me.ButtonEventHandling.Name = "ButtonEventHandling"
        Me.ButtonEventHandling.Size = New System.Drawing.Size(224, 30)
        Me.ButtonEventHandling.TabIndex = 47
        Me.ButtonEventHandling.Text = "Event Handling"
        '
        'ButtonConveyorSettings
        '
        Me.ButtonConveyorSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonConveyorSettings.Location = New System.Drawing.Point(24, 216)
        Me.ButtonConveyorSettings.Name = "ButtonConveyorSettings"
        Me.ButtonConveyorSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonConveyorSettings.TabIndex = 45
        Me.ButtonConveyorSettings.Text = "Conveyor"
        '
        'ButtonThermalSettings
        '
        Me.ButtonThermalSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonThermalSettings.Location = New System.Drawing.Point(264, 216)
        Me.ButtonThermalSettings.Name = "ButtonThermalSettings"
        Me.ButtonThermalSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonThermalSettings.TabIndex = 46
        Me.ButtonThermalSettings.Text = "Heater"
        '
        'ButtonSPCLogging
        '
        Me.ButtonSPCLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCLogging.Location = New System.Drawing.Point(24, 160)
        Me.ButtonSPCLogging.Name = "ButtonSPCLogging"
        Me.ButtonSPCLogging.Size = New System.Drawing.Size(224, 30)
        Me.ButtonSPCLogging.TabIndex = 51
        Me.ButtonSPCLogging.Text = "SPC Logging"
        '
        'ButtonDispenserSettings
        '
        Me.ButtonDispenserSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenserSettings.Location = New System.Drawing.Point(24, 48)
        Me.ButtonDispenserSettings.Name = "ButtonDispenserSettings"
        Me.ButtonDispenserSettings.Size = New System.Drawing.Size(224, 30)
        Me.ButtonDispenserSettings.TabIndex = 42
        Me.ButtonDispenserSettings.Text = "Dispenser"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(136, 32)
        Me.Label17.TabIndex = 32
        Me.Label17.Text = "Basic Setup"
        '
        'PanelLeft
        '
        Me.PanelLeft.BackColor = System.Drawing.Color.LightSteelBlue
        Me.PanelLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelLeft.Controls.Add(Me.RichTextBox1)
        Me.PanelLeft.Location = New System.Drawing.Point(0, 56)
        Me.PanelLeft.Name = "PanelLeft"
        Me.PanelLeft.Size = New System.Drawing.Size(768, 311)
        Me.PanelLeft.TabIndex = 33
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.DarkBlue
        Me.RichTextBox1.Location = New System.Drawing.Point(47, 16)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(288, 272)
        Me.RichTextBox1.TabIndex = 19
        Me.RichTextBox1.Text = "It is highly recommended to go through the entire process setup for a new pattern" & _
        "." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "For an existing pattern, the process setup can also be modified."
        '
        'Settings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1284, 878)
        Me.Controls.Add(Me.PanelRight)
        Me.Controls.Add(Me.PanelLeft)
        Me.Name = "Settings"
        Me.Text = "Basic Setup"
        Me.PanelRight.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.PanelLeft.ResumeLayout(False)
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
        CurrentPanel = MyNeedleCalibrationSettings.PanelToBeAdded
        AddPanel(PanelRight, MyNeedleCalibrationSettings.PanelToBeAdded)
        If IDS.Data.Hardware.Dispenser.Left.HeadType = "Jetting Valve" Then
            MyNeedleCalibrationSettings.BoxStep1.Visible = False
            MyNeedleCalibrationSettings.BoxStep1Jetting.Visible = True
        Else
            MyNeedleCalibrationSettings.BoxStep1.Visible = True
            MyNeedleCalibrationSettings.BoxStep1Jetting.Visible = False
        End If
        MyNeedleCalibrationSettings.RevertData()
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

    Private Sub ButtonConveyorSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorSettings.Click
        CurrentPanel = MyConveyorSettings.PanelToBeAdded
        AddPanel(PanelRight, MyConveyorSettings.PanelToBeAdded)
        Conveyor.OpenPort()
        Conveyor.PositionTimer.Start()
        MyConveyorSettings.PositionTimer.Start()
        MyConveyorSettings.RevertData()
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