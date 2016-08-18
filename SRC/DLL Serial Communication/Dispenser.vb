Public Class Dispenser

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
    Friend WithEvents Download As System.Windows.Forms.Button
    Friend WithEvents RPMLabel As System.Windows.Forms.Label
    Friend WithEvents HeadType As System.Windows.Forms.ComboBox
    Friend WithEvents RetractDelayLabel As System.Windows.Forms.Label
    Friend WithEvents RetractTimeLabel As System.Windows.Forms.Label
    Friend WithEvents RetractDelayLabel2 As System.Windows.Forms.Label
    Friend WithEvents RetractTimeLabel2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CurrentHeads As System.Windows.Forms.ComboBox
    Friend WithEvents AxMSComm1 As AxMSCommLib.AxMSComm
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents SliderOrJetting As System.Windows.Forms.GroupBox
    Friend WithEvents Auger As System.Windows.Forms.GroupBox
    Friend WithEvents Pulse As System.Windows.Forms.NumericUpDown
    Friend WithEvents pause As System.Windows.Forms.NumericUpDown
    Friend WithEvents Count As System.Windows.Forms.NumericUpDown
    Friend WithEvents SliderOrJettingTemperature As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractDelay As System.Windows.Forms.NumericUpDown
    Friend WithEvents RPM As System.Windows.Forms.NumericUpDown
    Friend WithEvents RetractTime As System.Windows.Forms.NumericUpDown
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Dispenser))
        Me.Download = New System.Windows.Forms.Button
        Me.RPMLabel = New System.Windows.Forms.Label
        Me.HeadType = New System.Windows.Forms.ComboBox
        Me.RetractDelayLabel = New System.Windows.Forms.Label
        Me.RetractTimeLabel = New System.Windows.Forms.Label
        Me.RetractDelayLabel2 = New System.Windows.Forms.Label
        Me.RetractTimeLabel2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.SliderOrJetting = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Pulse = New System.Windows.Forms.NumericUpDown
        Me.pause = New System.Windows.Forms.NumericUpDown
        Me.Count = New System.Windows.Forms.NumericUpDown
        Me.SliderOrJettingTemperature = New System.Windows.Forms.NumericUpDown
        Me.CurrentHeads = New System.Windows.Forms.ComboBox
        Me.AxMSComm1 = New AxMSCommLib.AxMSComm
        Me.Auger = New System.Windows.Forms.GroupBox
        Me.RetractDelay = New System.Windows.Forms.NumericUpDown
        Me.RPM = New System.Windows.Forms.NumericUpDown
        Me.RetractTime = New System.Windows.Forms.NumericUpDown
        Me.GroupBox1.SuspendLayout()
        Me.SliderOrJetting.SuspendLayout()
        CType(Me.Pulse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pause, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Count, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SliderOrJettingTemperature, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Auger.SuspendLayout()
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RPM, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RetractTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Download
        '
        Me.Download.Location = New System.Drawing.Point(204, 88)
        Me.Download.Name = "Download"
        Me.Download.Size = New System.Drawing.Size(184, 32)
        Me.Download.TabIndex = 53
        Me.Download.Text = "Download"
        '
        'RPMLabel
        '
        Me.RPMLabel.Location = New System.Drawing.Point(8, 40)
        Me.RPMLabel.Name = "RPMLabel"
        Me.RPMLabel.Size = New System.Drawing.Size(96, 24)
        Me.RPMLabel.TabIndex = 57
        Me.RPMLabel.Text = "Valve RPM"
        '
        'HeadType
        '
        Me.HeadType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.HeadType.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.HeadType.ItemHeight = 20
        Me.HeadType.Items.AddRange(New Object() {"Jetting Valve", "Auger Valve", "Slider Valve", "Time Pressure Valve", "Time Pressure Syringe"})
        Me.HeadType.Location = New System.Drawing.Point(204, 40)
        Me.HeadType.Name = "HeadType"
        Me.HeadType.Size = New System.Drawing.Size(184, 28)
        Me.HeadType.TabIndex = 46
        '
        'RetractDelayLabel
        '
        Me.RetractDelayLabel.Location = New System.Drawing.Point(8, 120)
        Me.RetractDelayLabel.Name = "RetractDelayLabel"
        Me.RetractDelayLabel.Size = New System.Drawing.Size(168, 24)
        Me.RetractDelayLabel.TabIndex = 59
        Me.RetractDelayLabel.Text = "Valve Retract Delay"
        '
        'RetractTimeLabel
        '
        Me.RetractTimeLabel.Location = New System.Drawing.Point(8, 80)
        Me.RetractTimeLabel.Name = "RetractTimeLabel"
        Me.RetractTimeLabel.Size = New System.Drawing.Size(160, 24)
        Me.RetractTimeLabel.TabIndex = 56
        Me.RetractTimeLabel.Text = "Valve Retract Time"
        '
        'RetractDelayLabel2
        '
        Me.RetractDelayLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.RetractDelayLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RetractDelayLabel2.Location = New System.Drawing.Point(248, 128)
        Me.RetractDelayLabel2.Name = "RetractDelayLabel2"
        Me.RetractDelayLabel2.Size = New System.Drawing.Size(32, 16)
        Me.RetractDelayLabel2.TabIndex = 50
        Me.RetractDelayLabel2.Text = "ms"
        Me.RetractDelayLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RetractTimeLabel2
        '
        Me.RetractTimeLabel2.BackColor = System.Drawing.SystemColors.Control
        Me.RetractTimeLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RetractTimeLabel2.Location = New System.Drawing.Point(248, 88)
        Me.RetractTimeLabel2.Name = "RetractTimeLabel2"
        Me.RetractTimeLabel2.Size = New System.Drawing.Size(32, 16)
        Me.RetractTimeLabel2.TabIndex = 49
        Me.RetractTimeLabel2.Text = "ms"
        Me.RetractTimeLabel2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Download)
        Me.GroupBox1.Controls.Add(Me.HeadType)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(24, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(584, 136)
        Me.GroupBox1.TabIndex = 64
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Dispenser Information"
        '
        'SliderOrJetting
        '
        Me.SliderOrJetting.Controls.Add(Me.Label1)
        Me.SliderOrJetting.Controls.Add(Me.Label3)
        Me.SliderOrJetting.Controls.Add(Me.Label4)
        Me.SliderOrJetting.Controls.Add(Me.Label6)
        Me.SliderOrJetting.Controls.Add(Me.Label8)
        Me.SliderOrJetting.Controls.Add(Me.Label9)
        Me.SliderOrJetting.Controls.Add(Me.Label5)
        Me.SliderOrJetting.Controls.Add(Me.Pulse)
        Me.SliderOrJetting.Controls.Add(Me.pause)
        Me.SliderOrJetting.Controls.Add(Me.Count)
        Me.SliderOrJetting.Controls.Add(Me.SliderOrJettingTemperature)
        Me.SliderOrJetting.Location = New System.Drawing.Point(328, 224)
        Me.SliderOrJetting.Name = "SliderOrJetting"
        Me.SliderOrJetting.Size = New System.Drawing.Size(280, 200)
        Me.SliderOrJetting.TabIndex = 56
        Me.SliderOrJetting.TabStop = False
        Me.SliderOrJetting.Text = "Slider/Jetting Valve Controller"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 24)
        Me.Label1.TabIndex = 67
        Me.Label1.Text = "Pulse Time"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 24)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "Count"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(104, 24)
        Me.Label4.TabIndex = 66
        Me.Label4.Text = "Pause Time"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(240, 88)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 16)
        Me.Label6.TabIndex = 64
        Me.Label6.Text = "ms"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(8, 160)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(112, 24)
        Me.Label8.TabIndex = 68
        Me.Label8.Text = "Temperature"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(240, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 16)
        Me.Label9.TabIndex = 65
        Me.Label9.Text = "ms"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(240, 168)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(16, 16)
        Me.Label5.TabIndex = 64
        Me.Label5.Text = "C"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Pulse
        '
        Me.Pulse.Location = New System.Drawing.Point(160, 40)
        Me.Pulse.Name = "Pulse"
        Me.Pulse.Size = New System.Drawing.Size(80, 27)
        Me.Pulse.TabIndex = 65
        Me.Pulse.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'pause
        '
        Me.pause.Location = New System.Drawing.Point(160, 80)
        Me.pause.Name = "pause"
        Me.pause.Size = New System.Drawing.Size(80, 27)
        Me.pause.TabIndex = 65
        Me.pause.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'Count
        '
        Me.Count.Location = New System.Drawing.Point(160, 120)
        Me.Count.Name = "Count"
        Me.Count.Size = New System.Drawing.Size(80, 27)
        Me.Count.TabIndex = 65
        Me.Count.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'SliderOrJettingTemperature
        '
        Me.SliderOrJettingTemperature.Location = New System.Drawing.Point(160, 160)
        Me.SliderOrJettingTemperature.Name = "SliderOrJettingTemperature"
        Me.SliderOrJettingTemperature.Size = New System.Drawing.Size(80, 27)
        Me.SliderOrJettingTemperature.TabIndex = 65
        Me.SliderOrJettingTemperature.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'CurrentHeads
        '
        Me.CurrentHeads.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CurrentHeads.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CurrentHeads.ItemHeight = 20
        Me.CurrentHeads.Items.AddRange(New Object() {"Left Head", "Right Head"})
        Me.CurrentHeads.Location = New System.Drawing.Point(224, 24)
        Me.CurrentHeads.Name = "CurrentHeads"
        Me.CurrentHeads.Size = New System.Drawing.Size(200, 28)
        Me.CurrentHeads.TabIndex = 46
        '
        'AxMSComm1
        '
        Me.AxMSComm1.Enabled = True
        Me.AxMSComm1.Location = New System.Drawing.Point(592, 0)
        Me.AxMSComm1.Name = "AxMSComm1"
        Me.AxMSComm1.OcxState = CType(resources.GetObject("AxMSComm1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxMSComm1.Size = New System.Drawing.Size(38, 38)
        Me.AxMSComm1.TabIndex = 48
        '
        'Auger
        '
        Me.Auger.Controls.Add(Me.RPMLabel)
        Me.Auger.Controls.Add(Me.RetractDelayLabel)
        Me.Auger.Controls.Add(Me.RetractTimeLabel)
        Me.Auger.Controls.Add(Me.RetractDelay)
        Me.Auger.Controls.Add(Me.RPM)
        Me.Auger.Controls.Add(Me.RetractTime)
        Me.Auger.Controls.Add(Me.RetractDelayLabel2)
        Me.Auger.Controls.Add(Me.RetractTimeLabel2)
        Me.Auger.Location = New System.Drawing.Point(24, 224)
        Me.Auger.Name = "Auger"
        Me.Auger.Size = New System.Drawing.Size(288, 200)
        Me.Auger.TabIndex = 56
        Me.Auger.TabStop = False
        Me.Auger.Text = "Auger Valve Controller"
        '
        'RetractDelay
        '
        Me.RetractDelay.DecimalPlaces = 2
        Me.RetractDelay.Location = New System.Drawing.Point(176, 120)
        Me.RetractDelay.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RetractDelay.Name = "RetractDelay"
        Me.RetractDelay.Size = New System.Drawing.Size(72, 27)
        Me.RetractDelay.TabIndex = 65
        Me.RetractDelay.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'RPM
        '
        Me.RPM.DecimalPlaces = 1
        Me.RPM.Location = New System.Drawing.Point(176, 40)
        Me.RPM.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RPM.Name = "RPM"
        Me.RPM.Size = New System.Drawing.Size(72, 27)
        Me.RPM.TabIndex = 65
        Me.RPM.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'RetractTime
        '
        Me.RetractTime.DecimalPlaces = 2
        Me.RetractTime.Location = New System.Drawing.Point(176, 80)
        Me.RetractTime.Maximum = New Decimal(New Integer() {999, 0, 0, 0})
        Me.RetractTime.Name = "RetractTime"
        Me.RetractTime.Size = New System.Drawing.Size(72, 27)
        Me.RetractTime.TabIndex = 65
        Me.RetractTime.Value = New Decimal(New Integer() {33, 0, 0, 0})
        '
        'Dispenser
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 20)
        Me.ClientSize = New System.Drawing.Size(632, 542)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.SliderOrJetting)
        Me.Controls.Add(Me.CurrentHeads)
        Me.Controls.Add(Me.AxMSComm1)
        Me.Controls.Add(Me.Auger)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Dispenser"
        Me.Text = "Form1"
        Me.GroupBox1.ResumeLayout(False)
        Me.SliderOrJetting.ResumeLayout(False)
        CType(Me.Pulse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pause, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Count, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SliderOrJettingTemperature, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxMSComm1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Auger.ResumeLayout(False)
        CType(Me.RetractDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RPM, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RetractTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub HeadType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HeadType.SelectedIndexChanged
        If HeadType.Text = "Jetting Valve" Then
            Auger.Hide()
            SliderOrJetting.Show()
        ElseIf HeadType.Text = "Auger Valve" Then
            Auger.Show()
            SliderOrJetting.Hide()
        ElseIf HeadType.Text = "Slider Valve" Then
            Auger.Hide()
            SliderOrJetting.Show()
        ElseIf HeadType.Text = "Time Pressure Valve" Then
            Auger.Hide()
            SliderOrJetting.Hide()
        ElseIf HeadType.Text = "Time Pressure Syringe" Then
            Auger.Hide()
            SliderOrJetting.Hide()
        End If
    End Sub

    Private Sub Download_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Download.Click
        Dim SendData, SendVal As String

        'for psi: IDS.Devices.Regulator.FrmRegulator.Left_SetPressure(1, pressure, suckback)
        'we use bar as the default value. see above if you want to change.

        If HeadType.SelectedItem.ToString = "Auger Valve" Then
            DownloadAugerRPM(RPM.Value, RetractTime.Value, RetractDelay.Value)
        ElseIf HeadType.SelectedItem.ToString = "Slider Valve" Or HeadType.SelectedItem.ToString = "Jetting Valve" Then
            DownloadJettingParameters(Pulse.Value, pause.Value, Count.Value, SliderOrJettingTemperature.Value)
        Else

        End If

    End Sub

    Public Function DownloadJettingParameters(ByVal Pulse As Double, ByVal Pause As Double, ByVal Count As Double, ByVal Temperature As Double) As Boolean
        Try
            With AxMSComm1
                .InBufferCount = 0
                .OutBufferCount = 0
                .SThreshold = 1
                .RThreshold = 25
            End With

            Dim SendData, SendVal As String
            'Pulse Time
            Call JettingData(Pulse.ToString, SendVal)
            SendData = SendData & Trim(SendVal)
            'Pause Time
            Call JettingData(Pause.ToString, SendVal)
            SendData = SendData & Trim(SendVal)
            'Count
            If Len(Count.ToString) = 1 Then
                SendData = SendData & "000" & Trim(Count.ToString)
            ElseIf Len(Count.ToString) = 2 Then
                SendData = SendData & "00" & Trim(Count.ToString)
            ElseIf Len(Count.ToString) = 3 Then
                SendData = SendData & "0" & Trim(Count.ToString)
            Else
                SendData = SendData & Trim(Count.ToString)
            End If
            'Temperature
            If Len(Temperature.ToString) = 1 Then
                SendData = SendData & "00" & Trim(Temperature.ToString)
            ElseIf Len(Temperature.ToString) = 2 Then
                SendData = SendData & "0" & Trim(Temperature.ToString)
            Else
                SendData = SendData & Trim(Temperature.ToString)
            End If

            SendData = SendData & "O"
            AxMSComm1.Output = SendData
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub DownloadAugerRPM(ByVal RPM As Double, ByVal RetractTime As Double, ByVal RetractDelay As Double)
        With AxMSComm1
            .InBufferCount = 0
            .OutBufferCount = 0
            .SThreshold = 21
            .RThreshold = 21
        End With
        'for RPM and Reverse RPM
        Dim SendData, SendVal As String
        SendData = "S"
        'Dispense Speed/RPM
        Call Format4Digits(RPM.ToString, SendVal)
        SendData = SendData & Trim(SendVal)
        'Dispense Time - not used in robot. It is only necessary when using AVC in standalone dispenser
        SendData = SendData & "00000"
        'Retract Time
        Call Format5Digits(RetractTime.ToString, SendVal)
        SendData = SendData & Trim(SendVal)
        'Retract Delay
        Call Format5Digits(RetractDelay.ToString, SendVal)
        SendData = SendData & Trim(SendVal)

        SendData = SendData & "O"
        AxMSComm1.Output = SendData
    End Sub

    Public Sub Format5Digits(ByVal toadd As String, ByRef tosend As String)
        If toadd.Length = 5 Then
            tosend = toadd
        ElseIf toadd.Length = 4 Then
            tosend = "0" & toadd
        ElseIf toadd.Length = 3 Then
            tosend = "00" & toadd
        ElseIf toadd.Length = 2 Then
            tosend = "000" & toadd
        ElseIf toadd.Length = 1 Then
            tosend = "0000" & toadd
        Else
            MsgBox("wrong format")
        End If
    End Sub

    Public Sub Format4Digits(ByVal toadd As String, ByRef tosend As String)
        If toadd.Length = 4 Then
            tosend = toadd
        ElseIf toadd.Length = 3 Then
            tosend = "0" & toadd
        ElseIf toadd.Length = 2 Then
            tosend = "00" & toadd
        ElseIf toadd.Length = 1 Then
            tosend = "000" & toadd
        Else
            MsgBox("wrong format")
        End If
    End Sub

    Public Sub JettingData(ByVal actualString As String, ByRef sendString As String)
        Dim tempData As String
        With AxMSComm1
            .InBufferCount = 0
            .OutBufferCount = 0
            .SThreshold = 1
            .RThreshold = 21
        End With
        tempData = ""

        If actualString.Length = 1 Then
            tempData = "00000" & actualString
        ElseIf actualString.Length = 2 Then
            tempData = "0000" & actualString
        ElseIf actualString.Length = 3 Then
            tempData = "000" & actualString
        ElseIf actualString.Length = 4 Then
            tempData = "00" & actualString
        ElseIf actualString.Length = 5 Then
            tempData = "0" & actualString
        Else
            tempData = actualString
        End If

        sendString = tempData
    End Sub

    Public Sub OpenPort()
        If AxMSComm1.PortOpen = True Then
            Exit Sub
        End If
        With AxMSComm1
            .Settings = "38400,N,8,1"       'Set baud rate, parity, data bits and stop bits as string
            .OutBufferSize = 1              'Transmit buffer size - byte   
            .InBufferSize = 128             'Receive buffer size
            .InputLen = 0                   'Set number of characters to read from receive buffer, 0 to read entire buffer
            .RThreshold = 1                 'Set 2 characters to receive before generate OnComm event
            .InputMode = MSCommLib.InputModeConstants.comInputModeText  'Set to return text data
            .InBufferCount = 0                                          'Initialize receive buffer
            .CommPort = DispenserPort
            Try
                If Not .PortOpen Then
                    .PortOpen = True
                End If
            Catch ex As System.Runtime.InteropServices.COMException
                Exit Sub
            End Try
        End With
    End Sub

    Public Sub ClosePort()
        If AxMSComm1.PortOpen = True Then
            AxMSComm1.PortOpen = False
        End If
    End Sub

    Private Sub Dispenser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        OpenPort()
    End Sub

    Private Sub AxMSComm1_OnComm(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxMSComm1.OnComm
        With AxMSComm1
            Select Case .CommEvent
                Case Else
            End Select

        End With
    End Sub
End Class
