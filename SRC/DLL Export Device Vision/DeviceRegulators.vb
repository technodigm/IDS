Public Class DeviceRegulators
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
    Friend WithEvents Btn_BothStop As System.Windows.Forms.Button
    Friend WithEvents Btn_BothStart As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents RightSuckBack As System.Windows.Forms.NumericUpDown
    Friend WithEvents RightPosPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Btn_RightStop As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Right_Bar As System.Windows.Forms.RadioButton
    Friend WithEvents Right_PSI As System.Windows.Forms.RadioButton
    Friend WithEvents Btn_RightStart As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents LeftSuckBack As System.Windows.Forms.NumericUpDown
    Friend WithEvents LeftPosPressure As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Btn_LStop As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Left_Bar As System.Windows.Forms.RadioButton
    Friend WithEvents Left_PSI As System.Windows.Forms.RadioButton
    Friend WithEvents Btn_LStart As System.Windows.Forms.Button
    Friend WithEvents AxDAQAO1 As AxDAQAOLib.AxDAQAO
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DeviceRegulators))
        Me.Btn_BothStop = New System.Windows.Forms.Button
        Me.Btn_BothStart = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RightSuckBack = New System.Windows.Forms.NumericUpDown
        Me.RightPosPressure = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.Btn_RightStop = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.Right_Bar = New System.Windows.Forms.RadioButton
        Me.Right_PSI = New System.Windows.Forms.RadioButton
        Me.Btn_RightStart = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.AxDAQAO1 = New AxDAQAOLib.AxDAQAO
        Me.LeftSuckBack = New System.Windows.Forms.NumericUpDown
        Me.LeftPosPressure = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.Btn_LStop = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Left_Bar = New System.Windows.Forms.RadioButton
        Me.Left_PSI = New System.Windows.Forms.RadioButton
        Me.Btn_LStart = New System.Windows.Forms.Button
        Me.GroupBox2.SuspendLayout()
        CType(Me.RightSuckBack, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.RightPosPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        CType(Me.AxDAQAO1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LeftSuckBack, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LeftPosPressure, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_BothStop
        '
        Me.Btn_BothStop.Location = New System.Drawing.Point(233, 409)
        Me.Btn_BothStop.Name = "Btn_BothStop"
        Me.Btn_BothStop.Size = New System.Drawing.Size(63, 49)
        Me.Btn_BothStop.TabIndex = 16
        Me.Btn_BothStop.Text = "BothStop"
        '
        'Btn_BothStart
        '
        Me.Btn_BothStart.Location = New System.Drawing.Point(100, 409)
        Me.Btn_BothStart.Name = "Btn_BothStart"
        Me.Btn_BothStart.Size = New System.Drawing.Size(67, 49)
        Me.Btn_BothStart.TabIndex = 15
        Me.Btn_BothStart.Text = "BothStart"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RightSuckBack)
        Me.GroupBox2.Controls.Add(Me.RightPosPressure)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Btn_RightStop)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Right_Bar)
        Me.GroupBox2.Controls.Add(Me.Right_PSI)
        Me.GroupBox2.Controls.Add(Me.Btn_RightStart)
        Me.GroupBox2.Location = New System.Drawing.Point(13, 215)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(434, 180)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "RightHead"
        '
        'RightSuckBack
        '
        Me.RightSuckBack.DecimalPlaces = 2
        Me.RightSuckBack.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.RightSuckBack.Location = New System.Drawing.Point(200, 90)
        Me.RightSuckBack.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.RightSuckBack.Minimum = New Decimal(New Integer() {1, 0, 0, -2147483648})
        Me.RightSuckBack.Name = "RightSuckBack"
        Me.RightSuckBack.Size = New System.Drawing.Size(100, 20)
        Me.RightSuckBack.TabIndex = 11
        Me.RightSuckBack.Value = New Decimal(New Integer() {1, 0, 0, -2147483648})
        '
        'RightPosPressure
        '
        Me.RightPosPressure.DecimalPlaces = 2
        Me.RightPosPressure.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.RightPosPressure.Location = New System.Drawing.Point(200, 55)
        Me.RightPosPressure.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.RightPosPressure.Name = "RightPosPressure"
        Me.RightPosPressure.Size = New System.Drawing.Size(100, 20)
        Me.RightPosPressure.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(67, 55)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 20)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "RightPositivePressure"
        '
        'Btn_RightStop
        '
        Me.Btn_RightStop.Location = New System.Drawing.Point(220, 125)
        Me.Btn_RightStop.Name = "Btn_RightStop"
        Me.Btn_RightStop.Size = New System.Drawing.Size(62, 48)
        Me.Btn_RightStop.TabIndex = 4
        Me.Btn_RightStop.Text = "RightStop"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(67, 90)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(126, 20)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "RightSuckBackPressure"
        '
        'Right_Bar
        '
        Me.Right_Bar.Checked = True
        Me.Right_Bar.Location = New System.Drawing.Point(73, 21)
        Me.Right_Bar.Name = "Right_Bar"
        Me.Right_Bar.Size = New System.Drawing.Size(87, 21)
        Me.Right_Bar.TabIndex = 7
        Me.Right_Bar.TabStop = True
        Me.Right_Bar.Text = "Bar"
        '
        'Right_PSI
        '
        Me.Right_PSI.Location = New System.Drawing.Point(213, 21)
        Me.Right_PSI.Name = "Right_PSI"
        Me.Right_PSI.Size = New System.Drawing.Size(87, 21)
        Me.Right_PSI.TabIndex = 8
        Me.Right_PSI.Text = "PSI"
        '
        'Btn_RightStart
        '
        Me.Btn_RightStart.Location = New System.Drawing.Point(87, 125)
        Me.Btn_RightStart.Name = "Btn_RightStart"
        Me.Btn_RightStart.Size = New System.Drawing.Size(66, 48)
        Me.Btn_RightStart.TabIndex = 1
        Me.Btn_RightStart.Text = "RightStart"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AxDAQAO1)
        Me.GroupBox1.Controls.Add(Me.LeftSuckBack)
        Me.GroupBox1.Controls.Add(Me.LeftPosPressure)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Btn_LStop)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Left_Bar)
        Me.GroupBox1.Controls.Add(Me.Left_PSI)
        Me.GroupBox1.Controls.Add(Me.Btn_LStart)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 28)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(434, 180)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "LeftHead"
        '
        'AxDAQAO1
        '
        Me.AxDAQAO1.ContainingControl = Me
        Me.AxDAQAO1.Enabled = True
        Me.AxDAQAO1.Location = New System.Drawing.Point(376, 104)
        Me.AxDAQAO1.Name = "AxDAQAO1"
        Me.AxDAQAO1.OcxState = CType(resources.GetObject("AxDAQAO1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDAQAO1.Size = New System.Drawing.Size(33, 33)
        Me.AxDAQAO1.TabIndex = 11
        '
        'LeftSuckBack
        '
        Me.LeftSuckBack.DecimalPlaces = 2
        Me.LeftSuckBack.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.LeftSuckBack.Location = New System.Drawing.Point(200, 90)
        Me.LeftSuckBack.Maximum = New Decimal(New Integer() {0, 0, 0, 0})
        Me.LeftSuckBack.Minimum = New Decimal(New Integer() {100, 0, 0, -2147352576})
        Me.LeftSuckBack.Name = "LeftSuckBack"
        Me.LeftSuckBack.Size = New System.Drawing.Size(100, 20)
        Me.LeftSuckBack.TabIndex = 10
        '
        'LeftPosPressure
        '
        Me.LeftPosPressure.DecimalPlaces = 2
        Me.LeftPosPressure.Increment = New Decimal(New Integer() {1, 0, 0, 131072})
        Me.LeftPosPressure.Location = New System.Drawing.Point(200, 55)
        Me.LeftPosPressure.Maximum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.LeftPosPressure.Name = "LeftPosPressure"
        Me.LeftPosPressure.Size = New System.Drawing.Size(100, 20)
        Me.LeftPosPressure.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(67, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "LeftPositivePressure"
        '
        'Btn_LStop
        '
        Me.Btn_LStop.Location = New System.Drawing.Point(220, 125)
        Me.Btn_LStop.Name = "Btn_LStop"
        Me.Btn_LStop.Size = New System.Drawing.Size(62, 48)
        Me.Btn_LStop.TabIndex = 4
        Me.Btn_LStop.Text = "LeftStop"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(67, 90)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "LeftSuckBackPressure"
        '
        'Left_Bar
        '
        Me.Left_Bar.Checked = True
        Me.Left_Bar.Location = New System.Drawing.Point(73, 21)
        Me.Left_Bar.Name = "Left_Bar"
        Me.Left_Bar.Size = New System.Drawing.Size(87, 21)
        Me.Left_Bar.TabIndex = 7
        Me.Left_Bar.TabStop = True
        Me.Left_Bar.Text = "Bar"
        '
        'Left_PSI
        '
        Me.Left_PSI.Location = New System.Drawing.Point(213, 21)
        Me.Left_PSI.Name = "Left_PSI"
        Me.Left_PSI.Size = New System.Drawing.Size(87, 21)
        Me.Left_PSI.TabIndex = 8
        Me.Left_PSI.Text = "PSI"
        '
        'Btn_LStart
        '
        Me.Btn_LStart.Location = New System.Drawing.Point(87, 125)
        Me.Btn_LStart.Name = "Btn_LStart"
        Me.Btn_LStart.Size = New System.Drawing.Size(66, 48)
        Me.Btn_LStart.TabIndex = 1
        Me.Btn_LStart.Text = "LeftStart"
        '
        'DeviceRegulators
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(459, 485)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Btn_BothStop)
        Me.Controls.Add(Me.Btn_BothStart)
        Me.Controls.Add(Me.GroupBox2)
        Me.Name = "DeviceRegulators"
        Me.Text = "DeviceRegulators"
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.RightSuckBack, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.RightPosPressure, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.AxDAQAO1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LeftSuckBack, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LeftPosPressure, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Set Pressure"
    Const DeviceNumber As Integer = 2       'Define DeviceNumber =0
    Dim PressureValue As Single
    Dim SuckBackValue As Single
    Dim LeftBarOrPSI As Integer = 0
    Dim RightBarOrPSI As Integer = 0
    Dim PSItoBarConversionFactor As Single = 0.06895   'Note: 1 PSI = 0.06895 Bar
    Dim readvoltage() As Single
    Dim ErrCode As String = ""

    Enum Channel
        Left
        Right
    End Enum 'Channel

    'Public Sub RegulatorErr(ByRef ID As Object)
    '    If ErrCode = "Open failed" Then
    '        ID += 11  'Failed to open Regulator Driver
    '    ElseIf ErrCode = "Left pressure setting failed" Then
    '        ID += 12  'Failed to setup left pressure regulator
    '    ElseIf ErrCode = "Left suction setting failed" Then
    '        ID += 13  'Failed to setup left suction regulator
    '    ElseIf ErrCode = "Right pressure setting failed" Then
    '        ID += 14  'Failed to setup right pressure regulator
    '    ElseIf ErrCode = "Right suction setting failed" Then
    '        ID += 15  'Failed to setup right suction regulator
    '    ElseIf ErrCode = "Close failed" Then
    '        ID += 16  'Failed to open Regulator Driver
    '    End If
    '    ErrCode = ""
    'End Sub

    Sub SetPressure(ByVal PressureUnit As Integer, ByVal PositivePressure As Single, ByVal SuckBackPressure As Single, ByVal WhichChannel As Channel)
        'PressureConversion
        If PressureUnit = 0 Then 'Bar
            PressureValue = PositivePressure
            SuckBackValue = 5 * (SuckBackPressure) + 5
        ElseIf PressureUnit = 1 Then    'PSI 
            PressureValue = PositivePressure * PSItoBarConversionFactor
            SuckBackValue = (5 * (SuckBackPressure) + 5) * PSItoBarConversionFactor
        End If

        'Setup DA Card 
        AxDAQAO1.DeviceNumber = DeviceNumber
        If AxDAQAO1.OpenDevice And ErrCode = "" Then
            ErrCode = "Open failed."
        End If

        AxDAQAO1.DataType = DAQAOLib.DATA_TYPE.adReal
        AxDAQAO1.EventEnabled = True
        AxDAQAO1.OutputType = DAQAOLib.OUTPUT_TYPE.adVoltage

        'Setup Left/Right Channel
        If WhichChannel = Channel.Left Then
            'Output Left Positive Pressure
            AxDAQAO1.Channel = 0
            If AxDAQAO1.RealOutput(PressureValue) Then
                AxDAQAO1.CloseDevice()
                If ErrCode = "" Then
                    ErrCode = "Left pressure setting failed."
                End If
            End If

            'Output Left SuckBack Pressure
            AxDAQAO1.Channel = 1
            If AxDAQAO1.RealOutput(SuckBackValue) Then
                If AxDAQAO1.CloseDevice() And ErrCode = "" Then
                    ErrCode = "Close failed."
                End If
                If ErrCode = "" Then
                    ErrCode = "Left suction setting failed."
                End If
            End If

        ElseIf WhichChannel = Channel.Right Then
            'Output Right Positive Pressure
            AxDAQAO1.Channel = 2
            If AxDAQAO1.RealOutput(PressureValue) Then
                If AxDAQAO1.CloseDevice() And ErrCode = "" Then
                    ErrCode = "Close failed."
                End If
                If ErrCode = "" Then
                    ErrCode = "Right pressure setting failed."
                End If
            End If

            'Output Rightt SuckBack Pressure
            AxDAQAO1.Channel = 3
            If AxDAQAO1.RealOutput(SuckBackValue) Then
                If AxDAQAO1.CloseDevice() And ErrCode = "" Then
                    ErrCode = "Close failed."
                End If
                If ErrCode = "" Then
                    ErrCode = "Right suction setting failed."
                End If
            End If

        End If

    End Sub

    Function Left_SetPressure(ByVal PressureUnit As Integer, ByVal PositivePressure As Single, ByVal SuckBackPressure As Single)
        SetPressure(PressureUnit, PositivePressure, SuckBackPressure, Channel.Left)
    End Function

    Function Right_SetPressure(ByVal PressureUnit As Integer, ByVal PositivePressure As Single, ByVal SuckBackPressure As Single)
        SetPressure(PressureUnit, PositivePressure, SuckBackPressure, Channel.Right)
    End Function
#End Region

#Region "Left Head"
    Private Sub Left_Bar_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Left_Bar.CheckedChanged
        If Left_Bar.Checked Then
            LeftBarOrPSI = 0    ' Bar is selected

            LeftPosPressure.DecimalPlaces = 2
            LeftPosPressure.Maximum = 10.0
            LeftPosPressure.Minimum = 0.0
            LeftPosPressure.Increment = 0.01
            LeftPosPressure.Value = 0.0

            LeftSuckBack.DecimalPlaces = 2
            LeftSuckBack.Maximum = 0.0
            LeftSuckBack.Minimum = -1.0
            LeftSuckBack.Increment = 0.01
            LeftSuckBack.Value = -1.0
        Else
            LeftBarOrPSI = 1    ' PSI is selected

            LeftPosPressure.DecimalPlaces = 1
            LeftPosPressure.Maximum = 145.0326
            LeftPosPressure.Minimum = 0.0
            LeftPosPressure.Increment = 0.1
            LeftPosPressure.Value = 0.0

            LeftSuckBack.DecimalPlaces = 1
            LeftSuckBack.Maximum = 0.0
            LeftSuckBack.Minimum = -14.5033
            LeftSuckBack.Increment = 0.1
            LeftSuckBack.Value = -14.5

        End If
    End Sub

    Private Sub Btn_LStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_LStart.Click
        Left_SetPressure(LeftBarOrPSI, LeftPosPressure.Text, LeftSuckBack.Text)
    End Sub

    Private Sub Btn_Stop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_LStop.Click
        Left_SetPressure(LeftBarOrPSI, 0, -1)
    End Sub
#End Region

#Region "Right Head"
    Private Sub Right_Bar_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Right_Bar.CheckedChanged
        If Right_Bar.Checked Then
            RightBarOrPSI = 0    ' Bar is selected

            RightPosPressure.DecimalPlaces = 2
            RightPosPressure.Maximum = 10.0
            RightPosPressure.Minimum = 0.0
            RightPosPressure.Increment = 0.01
            RightPosPressure.Value = 0.0

            RightSuckBack.DecimalPlaces = 2
            RightSuckBack.Maximum = 0.0
            RightSuckBack.Minimum = -1.0
            RightSuckBack.Increment = 0.01
            RightSuckBack.Value = -1.0
        Else
            RightBarOrPSI = 1    ' PSI is selected

            RightPosPressure.DecimalPlaces = 1
            RightPosPressure.Maximum = 145.0326
            RightPosPressure.Minimum = 0.0
            RightPosPressure.Increment = 0.1
            RightPosPressure.Value = 0.0

            RightSuckBack.DecimalPlaces = 1
            RightSuckBack.Maximum = 0.0
            RightSuckBack.Minimum = -14.5033
            RightSuckBack.Increment = 0.1
            RightSuckBack.Value = -14.5
        End If
    End Sub

    Private Sub Btn_RightStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_RightStart.Click
        Right_SetPressure(RightBarOrPSI, RightPosPressure.Text, RightSuckBack.Text)
    End Sub

    Private Sub Btn_RightStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_RightStop.Click
        Right_SetPressure(RightBarOrPSI, 0, -1)
    End Sub
#End Region

#Region "Both Heads"
    Private Sub Btn_BothStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_BothStart.Click
        Left_SetPressure(LeftBarOrPSI, LeftPosPressure.Text, LeftSuckBack.Text)
        Right_SetPressure(RightBarOrPSI, RightPosPressure.Text, RightSuckBack.Text)

    End Sub

    Private Sub Btn_BothStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_BothStop.Click
        Left_SetPressure(LeftBarOrPSI, 0, -1)           'set output voltage = 0V
        Right_SetPressure(RightBarOrPSI, 0, -1)
    End Sub
#End Region


End Class
