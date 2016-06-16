Imports DLL_Export_Service

Public Class HardwareCommunicationSetup
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
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents ButtonPostLifterOn As System.Windows.Forms.Button
    Friend WithEvents BT_DownSignal As System.Windows.Forms.Button
    Friend WithEvents BT_UpSignal As System.Windows.Forms.Button
    Friend WithEvents BT_Retrieve As System.Windows.Forms.Button
    Friend WithEvents ButtonDispenseLifterON As System.Windows.Forms.Button
    Friend WithEvents ResetPLCLogic As System.Windows.Forms.Button
    Friend WithEvents BT_Release As System.Windows.Forms.Button
    Friend WithEvents ButtonPreLifterOn As System.Windows.Forms.Button
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents HeaterStatus As System.Windows.Forms.Label
    Friend WithEvents LifterStatus As System.Windows.Forms.Label
    Friend WithEvents ConveyorStatus As System.Windows.Forms.Label
    Friend WithEvents MotionControllerStatus As System.Windows.Forms.Label
    Friend WithEvents WeighingScaleStatus As System.Windows.Forms.Label
    Friend WithEvents ConveyorBox As System.Windows.Forms.GroupBox
    Friend WithEvents LifterBox As System.Windows.Forms.GroupBox
    Friend WithEvents HeaterButton As System.Windows.Forms.Button
    Friend WithEvents WeighingScaleButton As System.Windows.Forms.Button
    Friend WithEvents ConveyorButton As System.Windows.Forms.Button
    Friend WithEvents LaserButton As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(HardwareCommunicationSetup))
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.LifterBox = New System.Windows.Forms.GroupBox
        Me.ButtonPostLifterOn = New System.Windows.Forms.Button
        Me.ButtonDispenseLifterON = New System.Windows.Forms.Button
        Me.ButtonPreLifterOn = New System.Windows.Forms.Button
        Me.Label17 = New System.Windows.Forms.Label
        Me.ConveyorBox = New System.Windows.Forms.GroupBox
        Me.BT_UpSignal = New System.Windows.Forms.Button
        Me.BT_Retrieve = New System.Windows.Forms.Button
        Me.BT_Release = New System.Windows.Forms.Button
        Me.BT_DownSignal = New System.Windows.Forms.Button
        Me.ResetPLCLogic = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.HeaterStatus = New System.Windows.Forms.Label
        Me.LifterStatus = New System.Windows.Forms.Label
        Me.WeighingScaleStatus = New System.Windows.Forms.Label
        Me.MotionControllerStatus = New System.Windows.Forms.Label
        Me.ConveyorStatus = New System.Windows.Forms.Label
        Me.WeighingScaleButton = New System.Windows.Forms.Button
        Me.ConveyorButton = New System.Windows.Forms.Button
        Me.HeaterButton = New System.Windows.Forms.Button
        Me.LaserButton = New System.Windows.Forms.Button
        Me.PanelToBeAdded.SuspendLayout()
        Me.LifterBox.SuspendLayout()
        Me.ConveyorBox.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Controls.Add(Me.LifterBox)
        Me.PanelToBeAdded.Controls.Add(Me.Label17)
        Me.PanelToBeAdded.Controls.Add(Me.ConveyorBox)
        Me.PanelToBeAdded.Controls.Add(Me.GroupBox2)
        Me.PanelToBeAdded.Controls.Add(Me.WeighingScaleButton)
        Me.PanelToBeAdded.Controls.Add(Me.ConveyorButton)
        Me.PanelToBeAdded.Controls.Add(Me.HeaterButton)
        Me.PanelToBeAdded.Controls.Add(Me.LaserButton)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(36, 1)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(512, 944)
        Me.PanelToBeAdded.TabIndex = 6
        '
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(432, 8)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 48
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonExit.Visible = False
        '
        'LifterBox
        '
        Me.LifterBox.Controls.Add(Me.ButtonPostLifterOn)
        Me.LifterBox.Controls.Add(Me.ButtonDispenseLifterON)
        Me.LifterBox.Controls.Add(Me.ButtonPreLifterOn)
        Me.LifterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LifterBox.Location = New System.Drawing.Point(264, 352)
        Me.LifterBox.Name = "LifterBox"
        Me.LifterBox.Size = New System.Drawing.Size(232, 270)
        Me.LifterBox.TabIndex = 3
        Me.LifterBox.TabStop = False
        Me.LifterBox.Text = "Lifter Test"
        Me.LifterBox.Visible = False
        '
        'ButtonPostLifterOn
        '
        Me.ButtonPostLifterOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonPostLifterOn.Location = New System.Drawing.Point(16, 120)
        Me.ButtonPostLifterOn.Name = "ButtonPostLifterOn"
        Me.ButtonPostLifterOn.Size = New System.Drawing.Size(88, 48)
        Me.ButtonPostLifterOn.TabIndex = 28
        Me.ButtonPostLifterOn.Text = "Lifter 3 On"
        '
        'ButtonDispenseLifterON
        '
        Me.ButtonDispenseLifterON.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonDispenseLifterON.Location = New System.Drawing.Point(128, 40)
        Me.ButtonDispenseLifterON.Name = "ButtonDispenseLifterON"
        Me.ButtonDispenseLifterON.Size = New System.Drawing.Size(88, 48)
        Me.ButtonDispenseLifterON.TabIndex = 27
        Me.ButtonDispenseLifterON.Text = "Lifter 2 On"
        '
        'ButtonPreLifterOn
        '
        Me.ButtonPreLifterOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonPreLifterOn.Location = New System.Drawing.Point(16, 40)
        Me.ButtonPreLifterOn.Name = "ButtonPreLifterOn"
        Me.ButtonPreLifterOn.Size = New System.Drawing.Size(88, 48)
        Me.ButtonPreLifterOn.TabIndex = 26
        Me.ButtonPreLifterOn.Text = "Lifter 1 On"
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(176, 32)
        Me.Label17.TabIndex = 33
        Me.Label17.Text = "Hardware Setup"
        '
        'ConveyorBox
        '
        Me.ConveyorBox.Controls.Add(Me.BT_UpSignal)
        Me.ConveyorBox.Controls.Add(Me.BT_Retrieve)
        Me.ConveyorBox.Controls.Add(Me.BT_Release)
        Me.ConveyorBox.Controls.Add(Me.BT_DownSignal)
        Me.ConveyorBox.Controls.Add(Me.ResetPLCLogic)
        Me.ConveyorBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ConveyorBox.Location = New System.Drawing.Point(264, 64)
        Me.ConveyorBox.Name = "ConveyorBox"
        Me.ConveyorBox.Size = New System.Drawing.Size(232, 270)
        Me.ConveyorBox.TabIndex = 3
        Me.ConveyorBox.TabStop = False
        Me.ConveyorBox.Text = "Conveyor Test"
        Me.ConveyorBox.Visible = False
        '
        'BT_UpSignal
        '
        Me.BT_UpSignal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_UpSignal.Location = New System.Drawing.Point(16, 48)
        Me.BT_UpSignal.Name = "BT_UpSignal"
        Me.BT_UpSignal.Size = New System.Drawing.Size(88, 48)
        Me.BT_UpSignal.TabIndex = 21
        Me.BT_UpSignal.Text = "Up-Flow"
        '
        'BT_Retrieve
        '
        Me.BT_Retrieve.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Retrieve.Location = New System.Drawing.Point(16, 120)
        Me.BT_Retrieve.Name = "BT_Retrieve"
        Me.BT_Retrieve.Size = New System.Drawing.Size(88, 48)
        Me.BT_Retrieve.TabIndex = 19
        Me.BT_Retrieve.Text = "Retrieve"
        '
        'BT_Release
        '
        Me.BT_Release.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_Release.Location = New System.Drawing.Point(128, 120)
        Me.BT_Release.Name = "BT_Release"
        Me.BT_Release.Size = New System.Drawing.Size(88, 48)
        Me.BT_Release.TabIndex = 20
        Me.BT_Release.Text = "Release"
        '
        'BT_DownSignal
        '
        Me.BT_DownSignal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BT_DownSignal.Location = New System.Drawing.Point(128, 48)
        Me.BT_DownSignal.Name = "BT_DownSignal"
        Me.BT_DownSignal.Size = New System.Drawing.Size(88, 48)
        Me.BT_DownSignal.TabIndex = 22
        Me.BT_DownSignal.Text = "Down-Flow"
        '
        'ResetPLCLogic
        '
        Me.ResetPLCLogic.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ResetPLCLogic.Location = New System.Drawing.Point(16, 192)
        Me.ResetPLCLogic.Name = "ResetPLCLogic"
        Me.ResetPLCLogic.Size = New System.Drawing.Size(200, 48)
        Me.ResetPLCLogic.TabIndex = 31
        Me.ResetPLCLogic.Text = "Reset PLC Logic"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.HeaterStatus)
        Me.GroupBox2.Controls.Add(Me.LifterStatus)
        Me.GroupBox2.Controls.Add(Me.WeighingScaleStatus)
        Me.GroupBox2.Controls.Add(Me.MotionControllerStatus)
        Me.GroupBox2.Controls.Add(Me.ConveyorStatus)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(16, 64)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(232, 216)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Communication Interfaces"
        Me.GroupBox2.Visible = False
        '
        'HeaterStatus
        '
        Me.HeaterStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HeaterStatus.Location = New System.Drawing.Point(16, 24)
        Me.HeaterStatus.Name = "HeaterStatus"
        Me.HeaterStatus.Size = New System.Drawing.Size(64, 24)
        Me.HeaterStatus.TabIndex = 0
        Me.HeaterStatus.Text = "Heater"
        '
        'LifterStatus
        '
        Me.LifterStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LifterStatus.Location = New System.Drawing.Point(16, 62)
        Me.LifterStatus.Name = "LifterStatus"
        Me.LifterStatus.Size = New System.Drawing.Size(48, 24)
        Me.LifterStatus.TabIndex = 0
        Me.LifterStatus.Text = "Lifter"
        '
        'WeighingScaleStatus
        '
        Me.WeighingScaleStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.WeighingScaleStatus.Location = New System.Drawing.Point(16, 176)
        Me.WeighingScaleStatus.Name = "WeighingScaleStatus"
        Me.WeighingScaleStatus.Size = New System.Drawing.Size(136, 24)
        Me.WeighingScaleStatus.TabIndex = 0
        Me.WeighingScaleStatus.Text = "Weighing Scale"
        '
        'MotionControllerStatus
        '
        Me.MotionControllerStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MotionControllerStatus.Location = New System.Drawing.Point(16, 100)
        Me.MotionControllerStatus.Name = "MotionControllerStatus"
        Me.MotionControllerStatus.Size = New System.Drawing.Size(144, 24)
        Me.MotionControllerStatus.TabIndex = 0
        Me.MotionControllerStatus.Text = "Motion Controller"
        '
        'ConveyorStatus
        '
        Me.ConveyorStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ConveyorStatus.Location = New System.Drawing.Point(16, 138)
        Me.ConveyorStatus.Name = "ConveyorStatus"
        Me.ConveyorStatus.Size = New System.Drawing.Size(88, 24)
        Me.ConveyorStatus.TabIndex = 0
        Me.ConveyorStatus.Text = "Conveyor"
        '
        'WeighingScaleButton
        '
        Me.WeighingScaleButton.Location = New System.Drawing.Point(16, 476)
        Me.WeighingScaleButton.Name = "WeighingScaleButton"
        Me.WeighingScaleButton.Size = New System.Drawing.Size(232, 40)
        Me.WeighingScaleButton.TabIndex = 49
        Me.WeighingScaleButton.Text = "Weighing Scale Interface"
        Me.WeighingScaleButton.Visible = False
        '
        'ConveyorButton
        '
        Me.ConveyorButton.Location = New System.Drawing.Point(16, 534)
        Me.ConveyorButton.Name = "ConveyorButton"
        Me.ConveyorButton.Size = New System.Drawing.Size(232, 40)
        Me.ConveyorButton.TabIndex = 49
        Me.ConveyorButton.Text = "Conveyor Interface"
        Me.ConveyorButton.Visible = False
        '
        'HeaterButton
        '
        Me.HeaterButton.Location = New System.Drawing.Point(16, 418)
        Me.HeaterButton.Name = "HeaterButton"
        Me.HeaterButton.Size = New System.Drawing.Size(232, 40)
        Me.HeaterButton.TabIndex = 49
        Me.HeaterButton.Text = "Heater Interface"
        '
        'LaserButton
        '
        Me.LaserButton.Location = New System.Drawing.Point(16, 360)
        Me.LaserButton.Name = "LaserButton"
        Me.LaserButton.Size = New System.Drawing.Size(232, 40)
        Me.LaserButton.TabIndex = 49
        Me.LaserButton.Text = "Laser Interface"
        '
        'HardwareCommunicationSetup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "HardwareCommunicationSetup"
        Me.Text = "HardwareCommunicationSetup"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.LifterBox.ResumeLayout(False)
        Me.ConveyorBox.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub BT_Retrieve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Retrieve.Click
        Conveyor.Command("Retrieve")
    End Sub

    Private Sub BT_Release_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Release.Click
        Conveyor.Command("Release")
    End Sub

    Private Sub ResetPLCLogic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ResetPLCLogic.Click
        Conveyor.Command("Reset Conveyor Status")
    End Sub

    Private Sub BT_UpSignal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_UpSignal.Click
        DIO_Service.TriggerUpstream()
    End Sub

    Private Sub BT_DownSignal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DownSignal.Click
        Static Dim checkClick As Boolean = True
        If checkClick = True Then
            IDS.Devices.DIO.DIO_SetOutputBit(1, 3, True)
            mysleep(40)
            checkClick = False
        Else
            IDS.Devices.DIO.DIO_SetOutputBit(1, 3, False)
            mysleep(40)
            checkClick = True
        End If
    End Sub

    Private Sub ButtonPreLifterOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPreLifterOn.Click
        Conveyor.Command("Left Lifter")
        If (ButtonPreLifterOn.Text = "Lifter 1 Off") Then
            ButtonPreLifterOn.Text = "Lifter 1 On"
        Else
            ButtonPreLifterOn.Text = "Lifter 1 Off"
        End If
    End Sub

    Private Sub ButtonDispenseLifterON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDispenseLifterON.Click
        Conveyor.Command("Center Lifter")
        If (ButtonDispenseLifterON.Text = "Lifter 2 Off") Then
            ButtonDispenseLifterON.Text = "Lifter 2 On"
        Else
            ButtonDispenseLifterON.Text = "Lifter 2 Off"
        End If
    End Sub

    Private Sub ButtonPostLifterOn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPostLifterOn.Click
        Conveyor.Command("Right Lifter")
        If (ButtonPostLifterOn.Text = "Lifter 3 Off") Then
            ButtonPostLifterOn.Text = "Lifter 3 On"
        Else
            ButtonPostLifterOn.Text = "Lifter 3 Off"
        End If
    End Sub

    Public Sub RefreshDisplay()
        'WeighingScaleButton.Visible = MySetup.CheckBoxVolume.Checked 'yy
        'HeaterButton.Visible = MySetup.CheckBoxHeater.Checked
        'LifterBox.Visible = MySetup.CheckBoxLifter.Checked
    End Sub

    Public Sub UpdateStatus()
        'WeighingScaleStatus.Visible = MySetup.CheckBoxVolume.Checked
        'HeaterStatus.Visible = MySetup.CheckBoxHeater.Checked
        'LifterStatus.Visible = MySetup.CheckBoxLifter.Checked
    End Sub

    Public Sub ChangeStatus()

    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        RemovePanel(CurrentControl)
    End Sub

    Private Sub HeaterButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HeaterButton.Click
        'Heater.Show()
    End Sub

    Private Sub WeighingScaleButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WeighingScaleButton.Click
        Dim temp As CIDSService.CIDSServiceDevices.CIDSServiceWeighting = Weighting_Scale.Instance
        temp.Show()
    End Sub

    Private Sub ConveyorButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConveyorButton.Click
        Dim temp As CIDSService.CIDSServiceDevices.CIDSServiceConveyor = Conveyor.Instance
        temp.Show()
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LaserButton.Click
        OnLaser()
        Dim temp As CIDSService.CIDSServiceDevices.CIDSServiceLaser = Laser.Instance
        temp.Show()
    End Sub
End Class
