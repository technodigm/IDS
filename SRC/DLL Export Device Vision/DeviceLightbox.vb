Public Class DeviceLightbox
    Inherits System.Windows.Forms.Form
    Public Declare Sub Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents TextBox1 As System.Windows.Forms.TextBox
    Public WithEvents txtReadBack As System.Windows.Forms.TextBox
    Public WithEvents txtOutByte As System.Windows.Forms.TextBox
    Public WithEvents cmdByteOut As System.Windows.Forms.Button
    Public WithEvents cmdReadBack As System.Windows.Forms.Button
    Public WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TB_ScanTime As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Btn_StopScanInput As System.Windows.Forms.Button
    Friend WithEvents Btn_ScanInput As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TB_DigitalOutput As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents AxDAQDO1 As AxDAQDOLib.AxDAQDO
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(DeviceLightbox))
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.txtReadBack = New System.Windows.Forms.TextBox
        Me.txtOutByte = New System.Windows.Forms.TextBox
        Me.cmdByteOut = New System.Windows.Forms.Button
        Me.cmdReadBack = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TB_ScanTime = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Btn_StopScanInput = New System.Windows.Forms.Button
        Me.Btn_ScanInput = New System.Windows.Forms.Button
        Me.TB_DigitalOutput = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.AxDAQDO1 = New AxDAQDOLib.AxDAQDO
        CType(Me.AxDAQDO1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(193, 139)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(140, 20)
        Me.Label1.TabIndex = 68
        Me.Label1.Text = "Hex conversion of ByteOut"
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsReturn = True
        Me.TextBox1.AutoSize = False
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBox1.Location = New System.Drawing.Point(107, 139)
        Me.TextBox1.MaxLength = 0
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBox1.Size = New System.Drawing.Size(64, 24)
        Me.TextBox1.TabIndex = 67
        Me.TextBox1.Text = "00"
        '
        'txtReadBack
        '
        Me.txtReadBack.AcceptsReturn = True
        Me.txtReadBack.AutoSize = False
        Me.txtReadBack.BackColor = System.Drawing.SystemColors.Window
        Me.txtReadBack.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReadBack.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReadBack.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtReadBack.Location = New System.Drawing.Point(107, 187)
        Me.txtReadBack.MaxLength = 0
        Me.txtReadBack.Name = "txtReadBack"
        Me.txtReadBack.ReadOnly = True
        Me.txtReadBack.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReadBack.Size = New System.Drawing.Size(64, 24)
        Me.txtReadBack.TabIndex = 66
        Me.txtReadBack.Text = "00"
        '
        'txtOutByte
        '
        Me.txtOutByte.AcceptsReturn = True
        Me.txtOutByte.AutoSize = False
        Me.txtOutByte.BackColor = System.Drawing.SystemColors.Window
        Me.txtOutByte.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOutByte.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOutByte.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOutByte.Location = New System.Drawing.Point(107, 90)
        Me.txtOutByte.MaxLength = 0
        Me.txtOutByte.Name = "txtOutByte"
        Me.txtOutByte.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOutByte.Size = New System.Drawing.Size(64, 24)
        Me.txtOutByte.TabIndex = 65
        Me.txtOutByte.Text = "00"
        '
        'cmdByteOut
        '
        Me.cmdByteOut.BackColor = System.Drawing.SystemColors.Control
        Me.cmdByteOut.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdByteOut.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdByteOut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdByteOut.Location = New System.Drawing.Point(200, 90)
        Me.cmdByteOut.Name = "cmdByteOut"
        Me.cmdByteOut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdByteOut.Size = New System.Drawing.Size(80, 25)
        Me.cmdByteOut.TabIndex = 64
        Me.cmdByteOut.Text = "&Byte Out"
        '
        'cmdReadBack
        '
        Me.cmdReadBack.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReadBack.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReadBack.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReadBack.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReadBack.Location = New System.Drawing.Point(200, 187)
        Me.cmdReadBack.Name = "cmdReadBack"
        Me.cmdReadBack.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReadBack.Size = New System.Drawing.Size(80, 25)
        Me.cmdReadBack.TabIndex = 63
        Me.cmdReadBack.Text = "&Read Back"
        '
        'TextBox2
        '
        Me.TextBox2.AcceptsReturn = True
        Me.TextBox2.AutoSize = False
        Me.TextBox2.BackColor = System.Drawing.SystemColors.Window
        Me.TextBox2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.TextBox2.Location = New System.Drawing.Point(107, 229)
        Me.TextBox2.MaxLength = 0
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TextBox2.Size = New System.Drawing.Size(64, 24)
        Me.TextBox2.TabIndex = 70
        Me.TextBox2.Text = "00"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(193, 236)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(140, 20)
        Me.Label2.TabIndex = 71
        Me.Label2.Text = "Inversion of Hex input"
        '
        'TB_ScanTime
        '
        Me.TB_ScanTime.Location = New System.Drawing.Point(187, 284)
        Me.TB_ScanTime.Name = "TB_ScanTime"
        Me.TB_ScanTime.Size = New System.Drawing.Size(113, 20)
        Me.TB_ScanTime.TabIndex = 77
        Me.TB_ScanTime.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(87, 284)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(83, 20)
        Me.Label6.TabIndex = 76
        Me.Label6.Text = "Scan Time"
        '
        'Btn_StopScanInput
        '
        Me.Btn_StopScanInput.Location = New System.Drawing.Point(260, 367)
        Me.Btn_StopScanInput.Name = "Btn_StopScanInput"
        Me.Btn_StopScanInput.Size = New System.Drawing.Size(53, 28)
        Me.Btn_StopScanInput.TabIndex = 75
        Me.Btn_StopScanInput.Text = "Stop"
        '
        'Btn_ScanInput
        '
        Me.Btn_ScanInput.Location = New System.Drawing.Point(193, 367)
        Me.Btn_ScanInput.Name = "Btn_ScanInput"
        Me.Btn_ScanInput.Size = New System.Drawing.Size(54, 28)
        Me.Btn_ScanInput.TabIndex = 74
        Me.Btn_ScanInput.Text = "Scan"
        '
        'TB_DigitalOutput
        '
        Me.TB_DigitalOutput.Location = New System.Drawing.Point(187, 326)
        Me.TB_DigitalOutput.Name = "TB_DigitalOutput"
        Me.TB_DigitalOutput.Size = New System.Drawing.Size(113, 20)
        Me.TB_DigitalOutput.TabIndex = 73
        Me.TB_DigitalOutput.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(80, 326)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 20)
        Me.Label3.TabIndex = 72
        Me.Label3.Text = "Digital Output"
        '
        'AxDAQDO1
        '
        Me.AxDAQDO1.Enabled = True
        Me.AxDAQDO1.Location = New System.Drawing.Point(312, 32)
        Me.AxDAQDO1.Name = "AxDAQDO1"
        Me.AxDAQDO1.OcxState = CType(resources.GetObject("AxDAQDO1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.AxDAQDO1.Size = New System.Drawing.Size(33, 33)
        Me.AxDAQDO1.TabIndex = 78
        '
        'DeviceLightbox
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(386, 422)
        Me.Controls.Add(Me.AxDAQDO1)
        Me.Controls.Add(Me.TB_ScanTime)
        Me.Controls.Add(Me.TB_DigitalOutput)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txtReadBack)
        Me.Controls.Add(Me.txtOutByte)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Btn_StopScanInput)
        Me.Controls.Add(Me.Btn_ScanInput)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmdByteOut)
        Me.Controls.Add(Me.cmdReadBack)
        Me.Name = "DeviceLightbox"
        Me.Text = "DeviceLightbox"
        CType(Me.AxDAQDO1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Const DeviceNumber As Integer = 2     'Define DeviceNumber =0 'xiong: look at advantech device manager to determine the number
    Const OutputPort As Integer = 1         'Define OutputPort = 1
    Dim ErrCode As String



    Private Sub cmdByteOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdByteOut.Click
        TextBox1.Text = Hex(txtOutByte.Text)
        TextBox2.Text = Hex((Int(txtOutByte.Text) Xor Int(255)))    'inversion of hex 
        SetBrightness((Int(txtOutByte.Text) Xor Int(255)))
    End Sub


    Private Sub cmdReadBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReadBack.Click
        AxDAQDO1.OpenDevice()
        AxDAQDO1.Port = 1
        txtReadBack.Text = Hex(AxDAQDO1.ByteReadBack)
        AxDAQDO1.CloseDevice()
    End Sub

    Public Function LightBox_Err() As Integer
        Dim Err As Integer = 1001200
        Dim ErrID As Integer

        If ErrCode = "Open failed." Then
            Return ErrID = Err + 1 'Failed to connect to lightbox driver
        ElseIf ErrCode = "Setting failed." Then
            Return ErrID = Err + 2 'Failed to apply brightness setting
        ElseIf ErrCode = "Close failed." Then
            Return ErrID = Err + 3 'Improper ending of lightbox connection
        Else
            Return ErrID = Err      'No error encounter
        End If

    End Function

    Sub SetBrightness(ByVal Brightness As Byte)
        AxDAQDO1.DeviceNumber = DeviceNumber

        If AxDAQDO1.OpenDevice Then
            ErrCode = "Open failed."
        End If

        AxDAQDO1.Port = OutputPort
        AxDAQDO1.Mask = &HFF          'Getting the 1st 8 bits

        If AxDAQDO1.ByteOutput("&H" & Hex(Brightness Xor Int(255))) Then
            ErrCode = "Setting failed."
        End If

        AxDAQDO1.Port = 0
        AxDAQDO1.Bit = 0

        If AxDAQDO1.BitOutput(1) Then
            ErrCode = "Setting failed."
        End If

        'For i As Integer = 0 To 9000000

        'Next

        If AxDAQDO1.BitOutput(0) Then
            ErrCode = "Setting failed."
        End If

        If AxDAQDO1.CloseDevice() Then
            ErrCode = "Close failed."
        End If

        Sleep(10)
    End Sub

    Public Sub SetLaser(ByVal _on As Boolean)

        AxDAQDO1.DeviceNumber = DeviceNumber

        If AxDAQDO1.OpenDevice Then
            ErrCode = "Open failed."
        End If

        AxDAQDO1.Port = 0
        AxDAQDO1.Bit = 5
        AxDAQDO1.BitOutput(_on)

        If AxDAQDO1.CloseDevice() Then
            ErrCode = "Close failed."
        End If

        Sleep(10)

    End Sub

    Private Sub Btn_ScanInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_ScanInput.Click

    End Sub
End Class
