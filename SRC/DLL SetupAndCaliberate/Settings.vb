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
    Public WithEvents PanelRight As System.Windows.Forms.Panel
    Public WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents ButtonStationPositions As System.Windows.Forms.Button
    Friend WithEvents ButtonNeedleCalibSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonEventHandling As System.Windows.Forms.Button
    Friend WithEvents ButtonConveyorSettings As System.Windows.Forms.Button
    Friend WithEvents ButtonSPCLogging As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenserSettings As System.Windows.Forms.Button
    Public WithEvents ButtonVolumeCalibSettings As System.Windows.Forms.Button
    Public WithEvents ButtonThermalSettings As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Public WithEvents gbRelMove As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PanelRight = New System.Windows.Forms.Panel
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ButtonEventHandling = New System.Windows.Forms.Button
        Me.ButtonConveyorSettings = New System.Windows.Forms.Button
        Me.ButtonThermalSettings = New System.Windows.Forms.Button
        Me.ButtonSPCLogging = New System.Windows.Forms.Button
        Me.ButtonDispenserSettings = New System.Windows.Forms.Button
        Me.ButtonStationPositions = New System.Windows.Forms.Button
        Me.ButtonVolumeCalibSettings = New System.Windows.Forms.Button
        Me.ButtonNeedleCalibSettings = New System.Windows.Forms.Button
        Me.PanelLeft = New System.Windows.Forms.Panel
        Me.gbRelMove = New System.Windows.Forms.GroupBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Panel3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.PanelLeft.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelRight
        '
        Me.PanelRight.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelRight.Location = New System.Drawing.Point(768, 8)
        Me.PanelRight.Name = "PanelRight"
        Me.PanelRight.Size = New System.Drawing.Size(512, 911)
        Me.PanelRight.TabIndex = 32
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.GroupBox1)
        Me.Panel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel3.Location = New System.Drawing.Point(408, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(352, 304)
        Me.Panel3.TabIndex = 65
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ButtonEventHandling)
        Me.GroupBox1.Controls.Add(Me.ButtonConveyorSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonThermalSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonSPCLogging)
        Me.GroupBox1.Controls.Add(Me.ButtonDispenserSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonStationPositions)
        Me.GroupBox1.Controls.Add(Me.ButtonVolumeCalibSettings)
        Me.GroupBox1.Controls.Add(Me.ButtonNeedleCalibSettings)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(336, 304)
        Me.GroupBox1.TabIndex = 66
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Program Settings :"
        '
        'ButtonEventHandling
        '
        Me.ButtonEventHandling.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonEventHandling.Location = New System.Drawing.Point(176, 172)
        Me.ButtonEventHandling.Name = "ButtonEventHandling"
        Me.ButtonEventHandling.Size = New System.Drawing.Size(112, 48)
        Me.ButtonEventHandling.TabIndex = 47
        Me.ButtonEventHandling.Text = "Event Handling"
        '
        'ButtonConveyorSettings
        '
        Me.ButtonConveyorSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonConveyorSettings.Location = New System.Drawing.Point(48, 238)
        Me.ButtonConveyorSettings.Name = "ButtonConveyorSettings"
        Me.ButtonConveyorSettings.Size = New System.Drawing.Size(112, 48)
        Me.ButtonConveyorSettings.TabIndex = 45
        Me.ButtonConveyorSettings.Text = "Conveyor"
        '
        'ButtonThermalSettings
        '
        Me.ButtonThermalSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonThermalSettings.Location = New System.Drawing.Point(176, 238)
        Me.ButtonThermalSettings.Name = "ButtonThermalSettings"
        Me.ButtonThermalSettings.Size = New System.Drawing.Size(112, 48)
        Me.ButtonThermalSettings.TabIndex = 46
        Me.ButtonThermalSettings.Text = "Heater"
        Me.ButtonThermalSettings.Visible = False
        '
        'ButtonSPCLogging
        '
        Me.ButtonSPCLogging.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSPCLogging.Location = New System.Drawing.Point(48, 172)
        Me.ButtonSPCLogging.Name = "ButtonSPCLogging"
        Me.ButtonSPCLogging.Size = New System.Drawing.Size(112, 48)
        Me.ButtonSPCLogging.TabIndex = 51
        Me.ButtonSPCLogging.Text = "SPC Logging"
        '
        'ButtonDispenserSettings
        '
        Me.ButtonDispenserSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonDispenserSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenserSettings.Location = New System.Drawing.Point(48, 40)
        Me.ButtonDispenserSettings.Name = "ButtonDispenserSettings"
        Me.ButtonDispenserSettings.Size = New System.Drawing.Size(112, 48)
        Me.ButtonDispenserSettings.TabIndex = 42
        Me.ButtonDispenserSettings.Text = "Dispenser"
        '
        'ButtonStationPositions
        '
        Me.ButtonStationPositions.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonStationPositions.Location = New System.Drawing.Point(176, 40)
        Me.ButtonStationPositions.Name = "ButtonStationPositions"
        Me.ButtonStationPositions.Size = New System.Drawing.Size(112, 48)
        Me.ButtonStationPositions.TabIndex = 43
        Me.ButtonStationPositions.Text = "Gantry"
        '
        'ButtonVolumeCalibSettings
        '
        Me.ButtonVolumeCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonVolumeCalibSettings.Location = New System.Drawing.Point(176, 106)
        Me.ButtonVolumeCalibSettings.Name = "ButtonVolumeCalibSettings"
        Me.ButtonVolumeCalibSettings.Size = New System.Drawing.Size(112, 48)
        Me.ButtonVolumeCalibSettings.TabIndex = 46
        Me.ButtonVolumeCalibSettings.Text = "Volume Calibration"
        '
        'ButtonNeedleCalibSettings
        '
        Me.ButtonNeedleCalibSettings.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonNeedleCalibSettings.Location = New System.Drawing.Point(48, 106)
        Me.ButtonNeedleCalibSettings.Name = "ButtonNeedleCalibSettings"
        Me.ButtonNeedleCalibSettings.Size = New System.Drawing.Size(112, 48)
        Me.ButtonNeedleCalibSettings.TabIndex = 44
        Me.ButtonNeedleCalibSettings.Text = "Needle Calibration"
        '
        'PanelLeft
        '
        Me.PanelLeft.BackColor = System.Drawing.SystemColors.Control
        Me.PanelLeft.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelLeft.Controls.Add(Me.gbRelMove)
        Me.PanelLeft.Controls.Add(Me.Panel3)
        Me.PanelLeft.Location = New System.Drawing.Point(0, 56)
        Me.PanelLeft.Name = "PanelLeft"
        Me.PanelLeft.Size = New System.Drawing.Size(768, 311)
        Me.PanelLeft.TabIndex = 33
        '
        'gbRelMove
        '
        Me.gbRelMove.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRelMove.Location = New System.Drawing.Point(24, 0)
        Me.gbRelMove.Name = "gbRelMove"
        Me.gbRelMove.Size = New System.Drawing.Size(370, 304)
        Me.gbRelMove.TabIndex = 66
        Me.gbRelMove.TabStop = False
        Me.gbRelMove.Text = "XYZ Relative Movement :"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.Info
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBox1.ForeColor = System.Drawing.Color.DarkBlue
        Me.RichTextBox1.Location = New System.Drawing.Point(728, 0)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(49, 32)
        Me.RichTextBox1.TabIndex = 19
        Me.RichTextBox1.Text = "It is highly recommended to go through the entire process setup for a new pattern" & _
        "." & Microsoft.VisualBasic.ChrW(10) & Microsoft.VisualBasic.ChrW(10) & "For an existing pattern, the process setup can also be modified."
        Me.RichTextBox1.Visible = False
        '
        'Settings
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1284, 878)
        Me.Controls.Add(Me.PanelRight)
        Me.Controls.Add(Me.PanelLeft)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Name = "Settings"
        Me.Text = "Basic Setup"
        Me.Panel3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.PanelLeft.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private CurrentPanel As System.Windows.Forms.Panel
    Private selectedButton As System.Windows.Forms.Button

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
        ButtonSelected(sender)
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
        ButtonSelected(sender)
        AddPanel(PanelRight, MyGantrySettings.PanelToBeAdded)
        'initialize
        MyGantrySettings.StationPosition.SelectedIndex = 0
        MyGantrySettings.LeftHead.Checked = True
        MyGantrySettings.RevertData()
    End Sub

    Private Sub ButtonDispenserSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDispenserSettings.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MyDispenserSettings.PanelToBeAdded)
        MyDispenserSettings.HeadType.SelectedIndex = 0
        MyDispenserSettings.RevertData()
    End Sub

    Private Sub ButtonSPCLogging_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSPCLogging.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MySPCLogging.PanelToBeAdded)
        MySPCLogging.RevertData()
    End Sub

    Private Sub ButtonEventHandling_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEventHandling.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MyEventSettings.PanelToBeAdded)
        MyEventSettings.RevertData()
    End Sub

    Private Sub ButtonConveyorSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonConveyorSettings.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MyConveyorSettings.PanelToBeAdded)
        Conveyor.OpenPort()
        Conveyor.PositionTimer.Start()
        MyConveyorSettings.PositionTimer.Start()
        MyConveyorSettings.RevertData()
    End Sub

    Private Sub ButtonThermalSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonThermalSettings.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MyHeaterSettings.PanelToBeAdded)
        MyHeaterSettings.RevertData()
    End Sub

    Private Sub ButtonVolumeCalibSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonVolumeCalibSettings.Click
        ButtonSelected(sender)
        AddPanel(PanelRight, MyVolumeCalibrationSettings.PanelToBeAdded)
        MyVolumeCalibrationSettings.RevertData()
    End Sub

    Private Sub ButtonSelected(ByVal sender As System.Object)
        If Not (selectedButton Is Nothing) Then
            selectedButton.FlatStyle = FlatStyle.Standard
            selectedButton.BackColor = System.Drawing.SystemColors.Control
        End If
        selectedButton = DirectCast(sender, System.Windows.Forms.Button)
        selectedButton.FlatStyle = FlatStyle.Flat
        selectedButton.BackColor = System.Drawing.SystemColors.ControlDarkDark
    End Sub

    Public Function AddDefaultView()
        If CurrentPanel Is Nothing Then
            If Not (selectedButton Is Nothing) Then
                selectedButton.FlatStyle = FlatStyle.Standard
                selectedButton.BackColor = System.Drawing.SystemColors.Control
            End If
            selectedButton = ButtonDispenserSettings
            selectedButton.FlatStyle = FlatStyle.Flat
            selectedButton.BackColor = System.Drawing.SystemColors.ControlDarkDark
            AddPanel(PanelRight, MyDispenserSettings.PanelToBeAdded)
            MyDispenserSettings.HeadType.SelectedIndex = 0
            MyDispenserSettings.RevertData()
        End If
    End Function
End Class