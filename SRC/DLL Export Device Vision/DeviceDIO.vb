Public Class DeviceDIO
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
    Friend WithEvents GB_Inputs As System.Windows.Forms.GroupBox
    Friend WithEvents CB_Input0 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input1 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input2 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input3 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input4 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input5 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input6 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input7 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input8 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input9 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input10 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input11 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input12 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input13 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input14 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Input15 As System.Windows.Forms.CheckBox
    Friend WithEvents GB_Outputs As System.Windows.Forms.GroupBox
    Friend WithEvents CB_Output15 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output14 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output13 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output12 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output11 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output10 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output9 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output8 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output7 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output6 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output5 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output4 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output3 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output2 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output1 As System.Windows.Forms.CheckBox
    Friend WithEvents CB_Output0 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents AxDAQDI1 As AxDAQDILib.AxDAQDI
    Friend WithEvents AxDAQDO1 As AxDAQDOLib.AxDAQDO
    Friend WithEvents R_ON As System.Windows.Forms.Button
    Friend WithEvents Y_ON As System.Windows.Forms.Button
    Friend WithEvents G_ON As System.Windows.Forms.Button
    Friend WithEvents S_ON As System.Windows.Forms.Button
    Friend WithEvents R_OFF As System.Windows.Forms.Button
    Friend WithEvents Y_OFF As System.Windows.Forms.Button
    Friend WithEvents G_OFF As System.Windows.Forms.Button
    Friend WithEvents S_OFF As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents TB_DoorSensor As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DeviceDIO))
        Me.GB_Inputs = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.CB_Input15 = New System.Windows.Forms.CheckBox
        Me.CB_Input14 = New System.Windows.Forms.CheckBox
        Me.CB_Input13 = New System.Windows.Forms.CheckBox
        Me.CB_Input12 = New System.Windows.Forms.CheckBox
        Me.CB_Input11 = New System.Windows.Forms.CheckBox
        Me.CB_Input10 = New System.Windows.Forms.CheckBox
        Me.CB_Input9 = New System.Windows.Forms.CheckBox
        Me.CB_Input8 = New System.Windows.Forms.CheckBox
        Me.CB_Input7 = New System.Windows.Forms.CheckBox
        Me.CB_Input6 = New System.Windows.Forms.CheckBox
        Me.CB_Input5 = New System.Windows.Forms.CheckBox
        Me.CB_Input4 = New System.Windows.Forms.CheckBox
        Me.CB_Input3 = New System.Windows.Forms.CheckBox
        Me.CB_Input2 = New System.Windows.Forms.CheckBox
        Me.CB_Input1 = New System.Windows.Forms.CheckBox
        Me.CB_Input0 = New System.Windows.Forms.CheckBox
        Me.GB_Outputs = New System.Windows.Forms.GroupBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.CB_Output15 = New System.Windows.Forms.CheckBox
        Me.CB_Output14 = New System.Windows.Forms.CheckBox
        Me.CB_Output13 = New System.Windows.Forms.CheckBox
        Me.CB_Output12 = New System.Windows.Forms.CheckBox
        Me.CB_Output11 = New System.Windows.Forms.CheckBox
        Me.CB_Output10 = New System.Windows.Forms.CheckBox
        Me.CB_Output9 = New System.Windows.Forms.CheckBox
        Me.CB_Output8 = New System.Windows.Forms.CheckBox
        Me.CB_Output7 = New System.Windows.Forms.CheckBox
        Me.CB_Output6 = New System.Windows.Forms.CheckBox
        Me.CB_Output5 = New System.Windows.Forms.CheckBox
        Me.CB_Output4 = New System.Windows.Forms.CheckBox
        Me.CB_Output3 = New System.Windows.Forms.CheckBox
        Me.CB_Output2 = New System.Windows.Forms.CheckBox
        Me.CB_Output1 = New System.Windows.Forms.CheckBox
        Me.CB_Output0 = New System.Windows.Forms.CheckBox
        Me.AxDAQDO1 = New AxDAQDOLib.AxDAQDO
        Me.AxDAQDI1 = New AxDAQDILib.AxDAQDI
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.R_ON = New System.Windows.Forms.Button
        Me.Y_ON = New System.Windows.Forms.Button
        Me.G_ON = New System.Windows.Forms.Button
        Me.S_ON = New System.Windows.Forms.Button
        Me.R_OFF = New System.Windows.Forms.Button
        Me.Y_OFF = New System.Windows.Forms.Button
        Me.G_OFF = New System.Windows.Forms.Button
        Me.S_OFF = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.TB_DoorSensor = New System.Windows.Forms.TextBox
        Me.GB_Inputs.SuspendLayout()
        Me.GB_Outputs.SuspendLayout()
        CType(Me.AxDAQDO1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AxDAQDI1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GB_Inputs
        '
        Me.GB_Inputs.Controls.Add(Me.Button2)
        Me.GB_Inputs.Controls.Add(Me.Button1)
        Me.GB_Inputs.Controls.Add(Me.TextBox6)
        Me.GB_Inputs.Controls.Add(Me.TextBox4)
        Me.GB_Inputs.Controls.Add(Me.CB_Input15)
        Me.GB_Inputs.Controls.Add(Me.CB_Input14)
        Me.GB_Inputs.Controls.Add(Me.CB_Input13)
        Me.GB_Inputs.Controls.Add(Me.CB_Input12)
        Me.GB_Inputs.Controls.Add(Me.CB_Input11)
        Me.GB_Inputs.Controls.Add(Me.CB_Input10)
        Me.GB_Inputs.Controls.Add(Me.CB_Input9)
        Me.GB_Inputs.Controls.Add(Me.CB_Input8)
        Me.GB_Inputs.Controls.Add(Me.CB_Input7)
        Me.GB_Inputs.Controls.Add(Me.CB_Input6)
        Me.GB_Inputs.Controls.Add(Me.CB_Input5)
        Me.GB_Inputs.Controls.Add(Me.CB_Input4)
        Me.GB_Inputs.Controls.Add(Me.CB_Input3)
        Me.GB_Inputs.Controls.Add(Me.CB_Input2)
        Me.GB_Inputs.Controls.Add(Me.CB_Input1)
        Me.GB_Inputs.Controls.Add(Me.CB_Input0)
        Me.GB_Inputs.Location = New System.Drawing.Point(12, 14)
        Me.GB_Inputs.Name = "GB_Inputs"
        Me.GB_Inputs.Size = New System.Drawing.Size(274, 457)
        Me.GB_Inputs.TabIndex = 2
        Me.GB_Inputs.TabStop = False
        Me.GB_Inputs.Text = "Inputs"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(140, 187)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 20)
        Me.Button2.TabIndex = 23
        Me.Button2.Text = "Port1"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(140, 159)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 20)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = "Port0"
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(140, 83)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(83, 20)
        Me.TextBox6.TabIndex = 20
        Me.TextBox6.Text = "TextBox6"
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(140, 55)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(83, 20)
        Me.TextBox4.TabIndex = 19
        Me.TextBox4.Text = "TextBox4"
        '
        'CB_Input15
        '
        Me.CB_Input15.Location = New System.Drawing.Point(13, 430)
        Me.CB_Input15.Name = "CB_Input15"
        Me.CB_Input15.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input15.TabIndex = 15
        Me.CB_Input15.Text = "CB_Input15"
        '
        'CB_Input14
        '
        Me.CB_Input14.Location = New System.Drawing.Point(13, 402)
        Me.CB_Input14.Name = "CB_Input14"
        Me.CB_Input14.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input14.TabIndex = 14
        Me.CB_Input14.Text = "CB_Input14"
        '
        'CB_Input13
        '
        Me.CB_Input13.Location = New System.Drawing.Point(13, 374)
        Me.CB_Input13.Name = "CB_Input13"
        Me.CB_Input13.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input13.TabIndex = 13
        Me.CB_Input13.Text = "CB_Input13"
        '
        'CB_Input12
        '
        Me.CB_Input12.Location = New System.Drawing.Point(13, 347)
        Me.CB_Input12.Name = "CB_Input12"
        Me.CB_Input12.Size = New System.Drawing.Size(87, 20)
        Me.CB_Input12.TabIndex = 12
        Me.CB_Input12.Text = "CB_Input12"
        '
        'CB_Input11
        '
        Me.CB_Input11.Location = New System.Drawing.Point(13, 319)
        Me.CB_Input11.Name = "CB_Input11"
        Me.CB_Input11.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input11.TabIndex = 11
        Me.CB_Input11.Text = "CB_Input11"
        '
        'CB_Input10
        '
        Me.CB_Input10.Location = New System.Drawing.Point(13, 291)
        Me.CB_Input10.Name = "CB_Input10"
        Me.CB_Input10.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input10.TabIndex = 10
        Me.CB_Input10.Text = "CB_Input10"
        '
        'CB_Input9
        '
        Me.CB_Input9.Location = New System.Drawing.Point(13, 263)
        Me.CB_Input9.Name = "CB_Input9"
        Me.CB_Input9.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input9.TabIndex = 9
        Me.CB_Input9.Text = "CB_Input9"
        '
        'CB_Input8
        '
        Me.CB_Input8.Location = New System.Drawing.Point(13, 236)
        Me.CB_Input8.Name = "CB_Input8"
        Me.CB_Input8.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input8.TabIndex = 8
        Me.CB_Input8.Text = "CB_Input8"
        '
        'CB_Input7
        '
        Me.CB_Input7.Location = New System.Drawing.Point(13, 208)
        Me.CB_Input7.Name = "CB_Input7"
        Me.CB_Input7.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input7.TabIndex = 7
        Me.CB_Input7.Text = "CB_Input7"
        '
        'CB_Input6
        '
        Me.CB_Input6.Location = New System.Drawing.Point(13, 180)
        Me.CB_Input6.Name = "CB_Input6"
        Me.CB_Input6.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input6.TabIndex = 6
        Me.CB_Input6.Text = "CB_Input6"
        '
        'CB_Input5
        '
        Me.CB_Input5.Location = New System.Drawing.Point(13, 153)
        Me.CB_Input5.Name = "CB_Input5"
        Me.CB_Input5.Size = New System.Drawing.Size(87, 20)
        Me.CB_Input5.TabIndex = 5
        Me.CB_Input5.Text = "CB_Input5"
        '
        'CB_Input4
        '
        Me.CB_Input4.Location = New System.Drawing.Point(13, 125)
        Me.CB_Input4.Name = "CB_Input4"
        Me.CB_Input4.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input4.TabIndex = 4
        Me.CB_Input4.Text = "CB_Input4"
        '
        'CB_Input3
        '
        Me.CB_Input3.Location = New System.Drawing.Point(13, 97)
        Me.CB_Input3.Name = "CB_Input3"
        Me.CB_Input3.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input3.TabIndex = 3
        Me.CB_Input3.Text = "CB_Input3"
        '
        'CB_Input2
        '
        Me.CB_Input2.Location = New System.Drawing.Point(13, 69)
        Me.CB_Input2.Name = "CB_Input2"
        Me.CB_Input2.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input2.TabIndex = 2
        Me.CB_Input2.Text = "CB_Input2"
        '
        'CB_Input1
        '
        Me.CB_Input1.Location = New System.Drawing.Point(13, 42)
        Me.CB_Input1.Name = "CB_Input1"
        Me.CB_Input1.Size = New System.Drawing.Size(87, 20)
        Me.CB_Input1.TabIndex = 1
        Me.CB_Input1.Text = "CB_Input1"
        '
        'CB_Input0
        '
        Me.CB_Input0.Location = New System.Drawing.Point(13, 14)
        Me.CB_Input0.Name = "CB_Input0"
        Me.CB_Input0.Size = New System.Drawing.Size(87, 21)
        Me.CB_Input0.TabIndex = 0
        Me.CB_Input0.Text = "CB_Input0"
        '
        'GB_Outputs
        '
        Me.GB_Outputs.Controls.Add(Me.TextBox3)
        Me.GB_Outputs.Controls.Add(Me.TextBox2)
        Me.GB_Outputs.Controls.Add(Me.TextBox1)
        Me.GB_Outputs.Controls.Add(Me.CB_Output15)
        Me.GB_Outputs.Controls.Add(Me.CB_Output14)
        Me.GB_Outputs.Controls.Add(Me.CB_Output13)
        Me.GB_Outputs.Controls.Add(Me.CB_Output12)
        Me.GB_Outputs.Controls.Add(Me.CB_Output11)
        Me.GB_Outputs.Controls.Add(Me.CB_Output10)
        Me.GB_Outputs.Controls.Add(Me.CB_Output9)
        Me.GB_Outputs.Controls.Add(Me.CB_Output8)
        Me.GB_Outputs.Controls.Add(Me.CB_Output7)
        Me.GB_Outputs.Controls.Add(Me.CB_Output6)
        Me.GB_Outputs.Controls.Add(Me.CB_Output5)
        Me.GB_Outputs.Controls.Add(Me.CB_Output4)
        Me.GB_Outputs.Controls.Add(Me.CB_Output3)
        Me.GB_Outputs.Controls.Add(Me.CB_Output2)
        Me.GB_Outputs.Controls.Add(Me.CB_Output1)
        Me.GB_Outputs.Controls.Add(Me.CB_Output0)
        Me.GB_Outputs.Location = New System.Drawing.Point(299, 14)
        Me.GB_Outputs.Name = "GB_Outputs"
        Me.GB_Outputs.Size = New System.Drawing.Size(273, 457)
        Me.GB_Outputs.TabIndex = 3
        Me.GB_Outputs.TabStop = False
        Me.GB_Outputs.Text = "Outputs"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(167, 118)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(83, 20)
        Me.TextBox3.TabIndex = 18
        Me.TextBox3.Text = "TextBox3"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(167, 90)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(83, 20)
        Me.TextBox2.TabIndex = 17
        Me.TextBox2.Text = "TextBox2"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(167, 62)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(83, 20)
        Me.TextBox1.TabIndex = 16
        Me.TextBox1.Text = "TextBox1"
        '
        'CB_Output15
        '
        Me.CB_Output15.Location = New System.Drawing.Point(13, 430)
        Me.CB_Output15.Name = "CB_Output15"
        Me.CB_Output15.Size = New System.Drawing.Size(114, 21)
        Me.CB_Output15.TabIndex = 15
        Me.CB_Output15.Text = "CB_Output15"
        '
        'CB_Output14
        '
        Me.CB_Output14.Location = New System.Drawing.Point(13, 402)
        Me.CB_Output14.Name = "CB_Output14"
        Me.CB_Output14.Size = New System.Drawing.Size(107, 21)
        Me.CB_Output14.TabIndex = 14
        Me.CB_Output14.Text = "CB_Output14"
        '
        'CB_Output13
        '
        Me.CB_Output13.Location = New System.Drawing.Point(13, 374)
        Me.CB_Output13.Name = "CB_Output13"
        Me.CB_Output13.Size = New System.Drawing.Size(114, 21)
        Me.CB_Output13.TabIndex = 13
        Me.CB_Output13.Text = "CB_Output13"
        '
        'CB_Output12
        '
        Me.CB_Output12.Location = New System.Drawing.Point(13, 347)
        Me.CB_Output12.Name = "CB_Output12"
        Me.CB_Output12.Size = New System.Drawing.Size(114, 20)
        Me.CB_Output12.TabIndex = 12
        Me.CB_Output12.Text = "CB_Output12"
        '
        'CB_Output11
        '
        Me.CB_Output11.Location = New System.Drawing.Point(13, 319)
        Me.CB_Output11.Name = "CB_Output11"
        Me.CB_Output11.Size = New System.Drawing.Size(114, 21)
        Me.CB_Output11.TabIndex = 11
        Me.CB_Output11.Text = "CB_Output11"
        '
        'CB_Output10
        '
        Me.CB_Output10.Location = New System.Drawing.Point(13, 291)
        Me.CB_Output10.Name = "CB_Output10"
        Me.CB_Output10.Size = New System.Drawing.Size(114, 21)
        Me.CB_Output10.TabIndex = 10
        Me.CB_Output10.Text = "CB_Output10"
        '
        'CB_Output9
        '
        Me.CB_Output9.Location = New System.Drawing.Point(13, 263)
        Me.CB_Output9.Name = "CB_Output9"
        Me.CB_Output9.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output9.TabIndex = 9
        Me.CB_Output9.Text = "CB_Output9"
        '
        'CB_Output8
        '
        Me.CB_Output8.Location = New System.Drawing.Point(13, 236)
        Me.CB_Output8.Name = "CB_Output8"
        Me.CB_Output8.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output8.TabIndex = 8
        Me.CB_Output8.Text = "CB_Output8"
        '
        'CB_Output7
        '
        Me.CB_Output7.Location = New System.Drawing.Point(13, 208)
        Me.CB_Output7.Name = "CB_Output7"
        Me.CB_Output7.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output7.TabIndex = 7
        Me.CB_Output7.Text = "CB_Output7"
        '
        'CB_Output6
        '
        Me.CB_Output6.Location = New System.Drawing.Point(13, 180)
        Me.CB_Output6.Name = "CB_Output6"
        Me.CB_Output6.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output6.TabIndex = 6
        Me.CB_Output6.Text = "CB_Output6"
        '
        'CB_Output5
        '
        Me.CB_Output5.Location = New System.Drawing.Point(13, 153)
        Me.CB_Output5.Name = "CB_Output5"
        Me.CB_Output5.Size = New System.Drawing.Size(87, 20)
        Me.CB_Output5.TabIndex = 5
        Me.CB_Output5.Text = "CB_Output5"
        '
        'CB_Output4
        '
        Me.CB_Output4.Location = New System.Drawing.Point(13, 125)
        Me.CB_Output4.Name = "CB_Output4"
        Me.CB_Output4.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output4.TabIndex = 4
        Me.CB_Output4.Text = "CB_Output4"
        '
        'CB_Output3
        '
        Me.CB_Output3.Location = New System.Drawing.Point(13, 97)
        Me.CB_Output3.Name = "CB_Output3"
        Me.CB_Output3.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output3.TabIndex = 3
        Me.CB_Output3.Text = "CB_Output3"
        '
        'CB_Output2
        '
        Me.CB_Output2.Location = New System.Drawing.Point(13, 69)
        Me.CB_Output2.Name = "CB_Output2"
        Me.CB_Output2.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output2.TabIndex = 2
        Me.CB_Output2.Text = "CB_Output2"
        '
        'CB_Output1
        '
        Me.CB_Output1.Location = New System.Drawing.Point(13, 42)
        Me.CB_Output1.Name = "CB_Output1"
        Me.CB_Output1.Size = New System.Drawing.Size(87, 20)
        Me.CB_Output1.TabIndex = 1
        Me.CB_Output1.Text = "CB_Output1"
        '
        'CB_Output0
        '
        Me.CB_Output0.Location = New System.Drawing.Point(13, 14)
        Me.CB_Output0.Name = "CB_Output0"
        Me.CB_Output0.Size = New System.Drawing.Size(87, 21)
        Me.CB_Output0.TabIndex = 0
        Me.CB_Output0.Text = "CB_Output0"
        '
        'AxDAQDO1
        '
        Me.AxDAQDO1.Enabled = True
        Me.AxDAQDO1.Location = New System.Drawing.Point(664, 48)
        Me.AxDAQDO1.Name = "AxDAQDO1"
        Me.AxDAQDO1.OcxState = CType(resources.GetObject("AxDAQDO1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDAQDO1.Size = New System.Drawing.Size(33, 33)
        Me.AxDAQDO1.TabIndex = 20
        '
        'AxDAQDI1
        '
        Me.AxDAQDI1.Enabled = True
        Me.AxDAQDI1.Location = New System.Drawing.Point(608, 48)
        Me.AxDAQDI1.Name = "AxDAQDI1"
        Me.AxDAQDI1.OcxState = CType(resources.GetObject("AxDAQDI1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDAQDI1.Size = New System.Drawing.Size(33, 33)
        Me.AxDAQDI1.TabIndex = 19
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(120, 480)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(440, 20)
        Me.TextBox5.TabIndex = 19
        Me.TextBox5.Text = "TextBox5"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 480)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 24
        Me.Label1.Text = "Error Message"
        '
        'TextBox7
        '
        Me.TextBox7.Location = New System.Drawing.Point(120, 512)
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(440, 20)
        Me.TextBox7.TabIndex = 25
        Me.TextBox7.Text = "TextBox7"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 512)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Error Message"
        '
        'TextBox8
        '
        Me.TextBox8.Location = New System.Drawing.Point(120, 544)
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(440, 20)
        Me.TextBox8.TabIndex = 27
        Me.TextBox8.Text = "TextBox8"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 544)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Error Message"
        '
        'R_ON
        '
        Me.R_ON.Location = New System.Drawing.Point(632, 104)
        Me.R_ON.Name = "R_ON"
        Me.R_ON.TabIndex = 29
        Me.R_ON.Text = "R_ON"
        '
        'Y_ON
        '
        Me.Y_ON.Location = New System.Drawing.Point(632, 144)
        Me.Y_ON.Name = "Y_ON"
        Me.Y_ON.TabIndex = 30
        Me.Y_ON.Text = "Y_ON"
        '
        'G_ON
        '
        Me.G_ON.Location = New System.Drawing.Point(632, 184)
        Me.G_ON.Name = "G_ON"
        Me.G_ON.TabIndex = 31
        Me.G_ON.Text = "G_ON"
        '
        'S_ON
        '
        Me.S_ON.Location = New System.Drawing.Point(632, 224)
        Me.S_ON.Name = "S_ON"
        Me.S_ON.TabIndex = 32
        Me.S_ON.Text = "S_ON"
        '
        'R_OFF
        '
        Me.R_OFF.Location = New System.Drawing.Point(632, 304)
        Me.R_OFF.Name = "R_OFF"
        Me.R_OFF.TabIndex = 29
        Me.R_OFF.Text = "R_OFF"
        '
        'Y_OFF
        '
        Me.Y_OFF.Location = New System.Drawing.Point(632, 344)
        Me.Y_OFF.Name = "Y_OFF"
        Me.Y_OFF.TabIndex = 30
        Me.Y_OFF.Text = "Y_OFF"
        '
        'G_OFF
        '
        Me.G_OFF.Location = New System.Drawing.Point(632, 384)
        Me.G_OFF.Name = "G_OFF"
        Me.G_OFF.TabIndex = 31
        Me.G_OFF.Text = "G_OFF"
        '
        'S_OFF
        '
        Me.S_OFF.Location = New System.Drawing.Point(632, 424)
        Me.S_OFF.Name = "S_OFF"
        Me.S_OFF.TabIndex = 32
        Me.S_OFF.Text = "S_OFF"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(632, 496)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 33
        Me.Button3.Text = "DoorSensor"
        '
        'TB_DoorSensor
        '
        Me.TB_DoorSensor.Location = New System.Drawing.Point(632, 536)
        Me.TB_DoorSensor.Name = "TB_DoorSensor"
        Me.TB_DoorSensor.Size = New System.Drawing.Size(72, 20)
        Me.TB_DoorSensor.TabIndex = 34
        Me.TB_DoorSensor.Text = "TextBox9"
        '
        'DeviceDIO
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(736, 582)
        Me.Controls.Add(Me.TB_DoorSensor)
        Me.Controls.Add(Me.TextBox8)
        Me.Controls.Add(Me.TextBox7)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.S_ON)
        Me.Controls.Add(Me.G_ON)
        Me.Controls.Add(Me.Y_ON)
        Me.Controls.Add(Me.R_ON)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GB_Outputs)
        Me.Controls.Add(Me.GB_Inputs)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.R_OFF)
        Me.Controls.Add(Me.Y_OFF)
        Me.Controls.Add(Me.G_OFF)
        Me.Controls.Add(Me.S_OFF)
        Me.Controls.Add(Me.AxDAQDO1)
        Me.Controls.Add(Me.AxDAQDI1)
        Me.Name = "DeviceDIO"
        Me.Text = "DeviceDIO"
        Me.GB_Inputs.ResumeLayout(False)
        Me.GB_Outputs.ResumeLayout(False)
        CType(Me.AxDAQDO1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AxDAQDI1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)  'delay function

#Region "SetOutputBit"
    Dim ErrCode As String

    Public Function CheckDIO() As Integer
        Dim Err As Integer = 1001200
        Dim ErrID As Integer
        If ErrCode = "Open failed." Then
            Return ErrID = Err + 21  ' failed to open device
        ElseIf ErrCode = "Setting failed." Then
            Return ErrID = Err + 22  ' failed to set i/o
        ElseIf ErrCode = "Close failed." Then
            Return ErrID = Err + 23  ' failed to close device
        ElseIf ErrCode = "Read failed." Then
            Return ErrID = Err + 24  ' failed to read device
        Else
            Return ErrID = Err      ' no error
        End If

    End Function

    ' changed by lsgoh
    Public Function SetOutputBit(ByVal PortNumber As Integer, ByVal Bitnumber As Integer, ByVal SetValue As Boolean) As Boolean
        AxDAQDO1.DeviceNumber = 0
        AxDAQDO1.Port = PortNumber


        If AxDAQDO1.OpenDevice() Then
            ErrCode = "Open failed."
            Return False
        End If

        AxDAQDO1.Bit = Bitnumber
        'AxDAQDO1.Mask = 0

        If AxDAQDO1.BitOutput(SetValue) Then
            ErrCode = "Setting failed."
            Return False
        End If

        TextBox1.Text = AxDAQDO1.Port
        TextBox2.Text = AxDAQDO1.Bit
        TextBox3.Text = Hex(2 ^ Bitnumber)
        TextBox7.Text = Hex(AxDAQDO1.ByteReadBack())
        TextBox8.Text = AxDAQDO1.Mask

        Return True

        ' If AxDAQDO1.CloseDevice() Then
        'ErrCode = "Close failed."
        'End If
    End Function

    Private Sub CB_Output0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output0.CheckedChanged
        If CB_Output0.Checked = True Then
            SetOutputBit(0, 0, True)
        ElseIf CB_Output0.Checked = False Then
            SetOutputBit(0, 0, False)
        End If

    End Sub

    Private Sub CB_Output1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output1.CheckedChanged
        If CB_Output1.Checked Then
            SetOutputBit(0, 1, True)
        Else
            SetOutputBit(0, 1, False)
        End If
    End Sub

    Private Sub CB_Output2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output2.CheckedChanged
        If CB_Output2.Checked Then
            SetOutputBit(0, 2, True)
        Else
            SetOutputBit(0, 2, False)
        End If
    End Sub

    Private Sub CB_Output3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output3.CheckedChanged
        If CB_Output3.Checked Then
            SetOutputBit(0, 3, True)
        Else
            SetOutputBit(0, 3, False)
        End If
    End Sub

    Private Sub CB_Output4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output4.CheckedChanged
        If CB_Output4.Checked Then
            SetOutputBit(0, 4, True)
        Else
            SetOutputBit(0, 4, False)
        End If
    End Sub

    Private Sub CB_Output5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output5.CheckedChanged
        If CB_Output5.Checked Then
            SetOutputBit(0, 5, True)
        Else
            SetOutputBit(0, 5, False)
        End If
    End Sub

    Private Sub CB_Output6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output6.CheckedChanged
        If CB_Output6.Checked Then
            SetOutputBit(0, 6, True)
        Else
            SetOutputBit(0, 6, False)
        End If
    End Sub

    Private Sub CB_Output7_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output7.CheckedChanged
        If CB_Output7.Checked Then
            SetOutputBit(0, 7, True)
        Else
            SetOutputBit(0, 7, False)
        End If
    End Sub

    Private Sub CB_Output8_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output8.CheckedChanged
        If CB_Output8.Checked Then
            SetOutputBit(1, 0, True)
        Else
            SetOutputBit(1, 0, False)
        End If
    End Sub

    Private Sub CB_Output9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output9.CheckedChanged
        If CB_Output9.Checked Then
            SetOutputBit(1, 1, True)
        Else
            SetOutputBit(1, 1, False)
        End If
    End Sub

    Private Sub CB_Output10_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output10.CheckedChanged
        If CB_Output10.Checked Then
            SetOutputBit(1, 2, True)
        Else
            SetOutputBit(1, 2, False)
        End If
    End Sub

    Private Sub CB_Output11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output11.CheckedChanged
        If CB_Output11.Checked Then
            SetOutputBit(1, 3, True)
        Else
            SetOutputBit(1, 3, False)
        End If
    End Sub

    Private Sub CB_Output12_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output12.CheckedChanged
        If CB_Output12.Checked Then
            SetOutputBit(1, 4, True)
        Else
            SetOutputBit(1, 4, False)
        End If
    End Sub

    Private Sub CB_Output13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output13.CheckedChanged
        If CB_Output13.Checked Then
            SetOutputBit(1, 5, True)
        Else
            SetOutputBit(1, 5, False)
        End If
    End Sub

    Private Sub CB_Output14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output14.CheckedChanged
        If CB_Output14.Checked Then
            SetOutputBit(1, 6, True)
        Else
            SetOutputBit(1, 6, False)
        End If
    End Sub

    Private Sub CB_Output15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Output15.CheckedChanged
        If CB_Output15.Checked Then
            SetOutputBit(1, 7, True)
        Else
            SetOutputBit(1, 7, False)
        End If
    End Sub
#End Region

#Region "CheckInputByte"
  

    'changed by lsgoh 
    Public Function CheckInputByte(ByVal PortNumber As Integer, ByRef ReadByte As Byte) As Boolean

        'AxDAQDI1.is_open()?
        'if open, then EXIT here

        AxDAQDI1.DeviceNumber = 0
        AxDAQDI1.Port = PortNumber

   
        If AxDAQDI1.OpenDevice() Then
            ErrCode = "Open failed."
            Return False
        End If

        'TextBox4.Text = AxDAQDI1.Port
        'TextBox6.Text = 255 - AxDAQDI1.ByteInput

        ReadByte = 255 - AxDAQDI1.ByteInput

        'ReadByte = AxDAQDI1.ByteInput
        'ReadByte = (&HFF And ReadByte)

        If AxDAQDI1.ByteScanValue Then
            ErrCode = "Read failed."
            Return False
        End If


        If AxDAQDI1.CloseDevice() Then
            ErrCode = "Close failed."
            Return False
        End If

        Return True
        'Return TextBox6.Text

    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ReadByte As Byte
        CheckInputByte(0, ReadByte)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ReadByte As Byte
        CheckInputByte(1, ReadByte)
    End Sub
#End Region

    Private Sub R_ON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R_ON.Click
        TowerLight_Red(True)
    End Sub
    Private Sub Y_ON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Y_ON.Click
        TowerLight_Yellow(True)
    End Sub
    Private Sub G_ON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles G_ON.Click
        TowerLight_Green(True)
    End Sub
    Private Sub S_ON_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles S_ON.Click
        TowerLight_Siren(True)
    End Sub
    Private Sub R_OFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles R_OFF.Click
        TowerLight_Red(False)
    End Sub
    Private Sub Y_OFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Y_OFF.Click
        TowerLight_Yellow(False)
    End Sub
    Private Sub G_OFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles G_OFF.Click
        TowerLight_Green(False)
    End Sub
    Private Sub S_OFF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles S_OFF.Click
        TowerLight_Siren(False)
    End Sub
    Private Sub TowerLight_Red(ByVal _on As Boolean)
        SetOutputBit(0, 3, _on)
    End Sub
    Private Sub TowerLight_Yellow(ByVal _on As Boolean)
        SetOutputBit(0, 5, _on)
    End Sub
    Private Sub TowerLight_Green(ByVal _on As Boolean)
        SetOutputBit(0, 7, _on)
    End Sub
    Private Sub TowerLight_Siren(ByVal _on As Boolean)
        SetOutputBit(1, 2, _on)
    End Sub

    Public Sub DoorLock(ByVal _on As Boolean)
        SetOutputBit(1, 7, Not _on)
    End Sub
    Public Sub TurnOnTowerLightRed()
        TowerLight_Red(True)
    End Sub
    Public Sub TurnOnTowerLightYellow()
        TowerLight_Yellow(True)
    End Sub
    Public Sub TurnOnTowerLightGreen()
        TowerLight_Green(True)
    End Sub
    Public Sub TurnOnTowerSiren()
        TowerLight_Siren(True)
    End Sub
    Public Sub TurnOffTowerLightRed()
        TowerLight_Red(False)
    End Sub
    Public Sub TurnOffTowerLightYellow()
        TowerLight_Yellow(False)
    End Sub
    Public Sub TurnOffTowerLightGreen()
        TowerLight_Green(False)
    End Sub
    Public Sub TurnOffTowerSiren()
        TowerLight_Siren(False)
    End Sub

    Public Function BoardEnabled() As Boolean
        AxDAQDI1.DeviceNumber = 0

        If AxDAQDI1.OpenDevice = False Then
            ErrCode = "Open failed."
            Return False
        End If

        Return True
    End Function

    Public Function BoardDisabled() As Boolean
        If AxDAQDI1.OpenDevice = True Then
            AxDAQDI1.CloseDevice()
        End If

        Return True
    End Function

    Public Function LevelSensor(ByVal Left As Boolean) As Boolean
        'If Left Then
        '    AxDAQDI1.DeviceNumber = 0
        '    If AxDAQDI1.OpenDevice() Then
        '        ErrCode = "Open failed."
        '        Return False
        '    End If

        '    AxDAQDI1.Port = 0
        '    AxDAQDI1.Bit = 7
        '    If Not AxDAQDI1.BitInput Then
        '        Return True
        '    End If
        'Else
        '    AxDAQDI1.DeviceNumber = 0
        '    If AxDAQDI1.OpenDevice() Then
        '        ErrCode = "Open failed."
        '        Return False
        '    End If

        '    AxDAQDI1.Port = 1
        '    AxDAQDI1.Bit = 1
        '    If Not AxDAQDI1.BitInput Then
        '        Return True
        '    End If
        'End If



        Dim IOBitMap As Integer = 255
        Dim ReadByte0 As Byte
        Dim ReadByte1 As Byte
        'PortNumber = 0
        'If IDS.Devices.DIO.CheckInputByte(PortNumber, ReadByte) Then
        If Left Then
            If CheckInputByte(0, ReadByte0) Then
                IOBitMap = CInt(ReadByte0)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
                ' convert the bit map to integer
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
                'Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 0) 'read input port 0 
                'If (IOBitMap And &H1) Then Return xx
                'If (IOBitMap And &H2) Then
                'Return 3
                'ElseIf (IOBitMap And &H4) Then
                'Return 9
                'ElseIf (IOBitMap And &H8) Then
                'Return 4
                'If (IOBitMap And &H10) Then Return xx
                'ElseIf (IOBitMap And &H20) Then
                'Return 5
                'ElseIf (IOBitMap And &H40) Then
                'Return 6
                If (IOBitMap And &H80) Then
                    Return True
                End If
            End If
        Else
            If CheckInputByte(1, ReadByte1) Then
                IOBitMap = CInt(ReadByte1)
                'If (IOBitMap And &H1) Then
                'Return 8
                If (IOBitMap And &H2) Then
                    Return True
                    'ElseIf (IOBitMap And &H4) Then
                    'Return 7
                    'If (IOBitMap And &H8) Then Return xx
                    'ElseIf (IOBitMap And &H10) Then
                    'Return 1
                    'If (IOBitMap And &H20) Then Return xx
                    'ElseIf (IOBitMap And &H40) Then
                    'Return 2
                    'If (IOBitMap And &H80) Then Return xx
                End If
                'Else
                'If failcount < 4 Then
                'MessageBox.Show("PC-IO Read error")
                'failcount += 1
                'End If
            End If
        End If
        Return False
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If doorclose_flag = True Then
            TB_DoorSensor.Text = "on"
        Else
            TB_DoorSensor.Text = "off"
        End If
    End Sub

    Public doorclose_flag As Boolean
    Public interlockon_flag As Boolean
    Public pressureerror_flag As Boolean = False

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '                                       kr                                       '
    ' checkoperationswitch is a IO check for any machine buttons being pressed       '
    ' this function is called continuously from machinestatus - thethreadmonitor     '
    ' we check door/key status by setting flags or check button pressed/pressure     '
    ' (done by bitwise conjunction by returning values)                              '
    ' thethreadmonitor is activated in operator mode from                            '
    ' FrmProgramming.ShowDialog() -> loadprogrammermaintenance -> formlogin.vb       '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'return: 0=Default; 1=Park/ChangeSyr; 2=Purge; 3=Clean; 4=NeedleCal; 5= VolumeCal
    'return: 6=Run; 7=Stop; 8=Pause; 9= Reset

    Public Function CheckOperationSwitch() As Integer

        Dim ReadByte0 As Byte = 0

        If CheckInputByte(0, ReadByte0) Then 'kr door_interlock and door_close respectively
            If ReadByte0 = 17 Then '0001 0001. 16 + 1
                interlockon_flag = True
                doorclose_flag = True
            ElseIf ReadByte0 = 16 Then '0001 0000. 0 + 1
                interlockon_flag = True
                doorclose_flag = False
            ElseIf ReadByte0 = 1 Then '0000 00001. 1 + 0
                interlockon_flag = False
                doorclose_flag = True
            ElseIf ReadByte0 = 0 Then '0000 0000. 0 + 0
                interlockon_flag = False
                doorclose_flag = False
            End If
        End If

        If CheckInputByte(1, ReadByte0) Then
            If (ReadByte0 And &H2) = 0 Then                'low pressure sensor
                pressureerror_flag = True
            End If
        End If

        ReadByte0 = 0
        If CheckInputByte(0, ReadByte0) Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
            ' convert the bit map to integer
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 
            'Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 0) 'read input port 0 
            'If (IOBitMap And &H1) Then Return xx
            'If (ReadByte0 And &H2) Then
            If (ReadByte0 And &H2) Then 'clean
                Return 3
            ElseIf (ReadByte0 And &H4) Then
                Return 9
            ElseIf (ReadByte0 And &H8) Then ' needlecalib
                Return 4
                'If (IOBitMap And &H10) Then Return xx
            ElseIf (ReadByte0 And &H20) Then
                Return 5
            ElseIf (ReadByte0 And &H40) Then
                Return 6
                'If (IOBitMap And &H80) Then Return xx
            End If
        End If
        If CheckInputByte(1, ReadByte0) Then
            'IOBitMap = CInt(ReadByte1)
            If (ReadByte0 And &H1) Then
                Return 8
                'If (IOBitMap And &H2) Then Return xx
            ElseIf (ReadByte0 And &H4) Then
                Return 7
                'If (IOBitMap And &H8) Then Return xx
            ElseIf (ReadByte0 And &H10) Then
                Return 1
                'If (IOBitMap And &H20) Then Return xx
            ElseIf (ReadByte0 And &H40) Then
                Return 2
                'If (IOBitMap And &H80) Then Return xx
            ElseIf (ReadByte0 And &H2) = 0 Then                ' LOW Pressure sensor  ' xu long
                Return 11
            End If
            'Else
            'If failcount < 4 Then
            'MessageBox.Show("PC-IO Read error")
            'failcount += 1
            'End If
        End If

        Return 0
    End Function

End Class