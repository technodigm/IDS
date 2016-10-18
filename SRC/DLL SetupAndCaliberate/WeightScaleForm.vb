
Public Class WeightScaleForm
    Inherits System.Windows.Forms.Form
    Private ambientMode As String = "Stable"
    Private timeOut As Double = 5000   'ms
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
    Friend WithEvents btGetWeight As System.Windows.Forms.Button
    Friend WithEvents btTare As System.Windows.Forms.Button
    Friend WithEvents rtbInfo As System.Windows.Forms.RichTextBox
    Friend WithEvents btOpenPort As System.Windows.Forms.Button
    Friend WithEvents btRestart As System.Windows.Forms.Button
    Friend WithEvents RestartTimer As System.Windows.Forms.Timer
    Friend WithEvents btZero As System.Windows.Forms.Button
    Friend WithEvents bt20 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.btGetWeight = New System.Windows.Forms.Button
        Me.btTare = New System.Windows.Forms.Button
        Me.rtbInfo = New System.Windows.Forms.RichTextBox
        Me.btOpenPort = New System.Windows.Forms.Button
        Me.btRestart = New System.Windows.Forms.Button
        Me.RestartTimer = New System.Windows.Forms.Timer(Me.components)
        Me.btZero = New System.Windows.Forms.Button
        Me.bt20 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btGetWeight
        '
        Me.btGetWeight.Enabled = False
        Me.btGetWeight.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btGetWeight.Location = New System.Drawing.Point(8, 8)
        Me.btGetWeight.Name = "btGetWeight"
        Me.btGetWeight.Size = New System.Drawing.Size(128, 48)
        Me.btGetWeight.TabIndex = 0
        Me.btGetWeight.Text = "Get Weight"
        '
        'btTare
        '
        Me.btTare.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btTare.Location = New System.Drawing.Point(160, 8)
        Me.btTare.Name = "btTare"
        Me.btTare.Size = New System.Drawing.Size(72, 48)
        Me.btTare.TabIndex = 1
        Me.btTare.Text = "Tare"
        '
        'rtbInfo
        '
        Me.rtbInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbInfo.HideSelection = False
        Me.rtbInfo.Location = New System.Drawing.Point(16, 104)
        Me.rtbInfo.Name = "rtbInfo"
        Me.rtbInfo.ReadOnly = True
        Me.rtbInfo.Size = New System.Drawing.Size(464, 568)
        Me.rtbInfo.TabIndex = 3
        Me.rtbInfo.Text = ""
        '
        'btOpenPort
        '
        Me.btOpenPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btOpenPort.Location = New System.Drawing.Point(352, 8)
        Me.btOpenPort.Name = "btOpenPort"
        Me.btOpenPort.Size = New System.Drawing.Size(128, 48)
        Me.btOpenPort.TabIndex = 4
        Me.btOpenPort.Text = "Open Port"
        Me.btOpenPort.Visible = False
        '
        'btRestart
        '
        Me.btRestart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btRestart.Location = New System.Drawing.Point(256, 8)
        Me.btRestart.Name = "btRestart"
        Me.btRestart.Size = New System.Drawing.Size(72, 48)
        Me.btRestart.TabIndex = 5
        Me.btRestart.Text = "Restart"
        '
        'RestartTimer
        '
        Me.RestartTimer.Interval = 5000
        '
        'btZero
        '
        Me.btZero.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btZero.Location = New System.Drawing.Point(160, 56)
        Me.btZero.Name = "btZero"
        Me.btZero.Size = New System.Drawing.Size(72, 48)
        Me.btZero.TabIndex = 6
        Me.btZero.Text = "Clear Text"
        '
        'bt20
        '
        Me.bt20.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bt20.Location = New System.Drawing.Point(352, 56)
        Me.bt20.Name = "bt20"
        Me.bt20.Size = New System.Drawing.Size(128, 48)
        Me.bt20.TabIndex = 7
        Me.bt20.Text = "Read 20 times"
        '
        'WeightScaleForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 688)
        Me.Controls.Add(Me.bt20)
        Me.Controls.Add(Me.btZero)
        Me.Controls.Add(Me.btRestart)
        Me.Controls.Add(Me.btOpenPort)
        Me.Controls.Add(Me.rtbInfo)
        Me.Controls.Add(Me.btTare)
        Me.Controls.Add(Me.btGetWeight)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "WeightScaleForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Weighting Scale "
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btGetWeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGetWeight.Click
        Weighting_Scale.RequestWeightUpdate = True
        btGetWeight.Enabled = False
        btTare.Enabled = False
        btRestart.Enabled = False
        AddInfo("Reading current weight")
        Weighting_Scale.ReadWeightReturned = New WeightScale.IWeightingScale.ReadWeightDel(AddressOf Me.WeightReturn)
        Weighting_Scale.GetWeight()
    End Sub
    Public Sub WeightReturn(ByVal weight As Double)
        AddInfo("Current weight: " & weight & " mg")
        If loopCount > 0 Then
            loopCount -= 1
        End If
        If loopCount > 0 Then
            AddInfo("Read #" & loopCount)
            Weighting_Scale.GetWeight()
        Else
            btGetWeight.Enabled = True
            btTare.Enabled = True
            btRestart.Enabled = True
        End If
    End Sub
    Private Sub btTare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btTare.Click
        Weighting_Scale.DoTare()
        AddInfo("Tared")
    End Sub

    Private Sub WeightScaleForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not (Weighting_Scale Is Nothing) Then
            If Weighting_Scale.OpenPort() Then
                Weighting_Scale.SetAmbient(ambientMode)
                btGetWeight.Enabled = True
                btTare.Enabled = True
                btOpenPort.Text = "Close Port"
            Else
                AddInfo("Unable to connect to weighting scale")
                btGetWeight.Enabled = False
                btTare.Enabled = False
                Me.bt20.Enabled = False
                Me.btRestart.Enabled = False
            End If
        Else
            btGetWeight.Enabled = False
            btTare.Enabled = False
            Me.bt20.Enabled = False
            Me.btRestart.Enabled = False
            AddInfo("Weighting scale is not available")
        End If
        'If Weighting_Scale.OpenPort() Then
        '    Weighting_Scale.SetAmbient(ambientMode)
        '    btGetWeight.Enabled = True
        '    btTare.Enabled = True
        '    btOpenPort.Text = "Close Port"
        'Else
        '    AddInfo("Unable to connect to weighting scale")
        '    btGetWeight.Enabled = False
        '    btTare.Enabled = False
        'End If
    End Sub
    Private lineCount As Integer = 0
    Private Sub AddInfo(ByVal info As String)
        If lineCount > 200 Then
            rtbInfo.Text = ""
            lineCount = 0
        End If
        rtbInfo.AppendText(info + vbLf)
        rtbInfo.SelectionStart = rtbInfo.Text.Length
        rtbInfo.ScrollToCaret()
        lineCount += 1
    End Sub
    'Private Sub WaitForData(ByVal timeOut As Double)
    '    Dim time As DateTime = DateTime.Now
    '    Dim diff As Long = ((DateTime.Now.Ticks - time.Ticks) / 10000)
    '    While Not (Weighting_Scale.ValueUpdated) And (diff < timeOut) 'ms
    '        Application.DoEvents()
    '        diff = (DateTime.Now.Ticks - time.Ticks) / 10000
    '    End While
    '    If Not (Weighting_Scale.ValueUpdated) Then
    '        AddInfo("Timeout error! Unable to get weighting data")
    '    Else
    '        AddInfo("Current weight: " & Weighting_Scale.returnString + " mg")
    '    End If
    '    btGetWeight.Enabled = True
    '    btTare.Enabled = True
    '    btRestart.Enabled = True
    'End Sub

    Private Sub btOpenPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btOpenPort.Click
        If btOpenPort.Text = "Open Port" Then
            If Weighting_Scale.OpenPort() Then
                Weighting_Scale.SetAmbient(ambientMode)
                btGetWeight.Enabled = True
                btTare.Enabled = True
                btOpenPort.Text = "Close Port"
            Else
                AddInfo("Unable to connect to weighting scale")
                btGetWeight.Enabled = False
                btTare.Enabled = False
            End If
        Else
            btOpenPort.Text = "Open Port"
            Weighting_Scale.ClosePort()
            btGetWeight.Enabled = False
            btTare.Enabled = False
        End If
    End Sub

    Private Sub WeightScaleForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        loopCount = 0
        e.Cancel = True
        RestartTimer.Stop()
        RestartTimer.Enabled = False
        btGetWeight.Enabled = True
        btTare.Enabled = True
        btRestart.Enabled = True
        Hide()
    End Sub

    Private Sub btRestart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btRestart.Click
        Weighting_Scale.Restart()
        Me.Enabled = False
        RestartTimer.Enabled = True
        AddInfo("Restarting Scale")
    End Sub

    Private Sub RestartTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestartTimer.Tick
        RestartTimer.Stop()
        RestartTimer.Enabled = False
        Me.Enabled = True
        AddInfo("Restarting Scale done")
    End Sub

    Private Sub btZero_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btZero.Click
        rtbInfo.Text = ""
        'Weighting_Scale.Zero()
    End Sub
    Private loopCount As Integer = 0
    Private Sub bt100_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt20.Click
        'For i As Integer = 0 To 19
        btGetWeight.Enabled = False
        btTare.Enabled = False
        btRestart.Enabled = False
        loopCount = 20
        AddInfo("Read #" & loopCount)
        Weighting_Scale.RequestWeightUpdate = True
        Weighting_Scale.ReadWeightReturned = New WeightScale.IWeightingScale.ReadWeightDel(AddressOf Me.WeightReturn)
        Weighting_Scale.GetWeight()
        'WaitForData(timeOut)
        'Next
        'AddInfo("20 read done")
        'btGetWeight.Enabled = True
        'btTare.Enabled = True
        'btRestart.Enabled = True
    End Sub
End Class
