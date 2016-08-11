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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btGetWeight = New System.Windows.Forms.Button
        Me.btTare = New System.Windows.Forms.Button
        Me.rtbInfo = New System.Windows.Forms.RichTextBox
        Me.btOpenPort = New System.Windows.Forms.Button
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
        Me.btTare.Location = New System.Drawing.Point(144, 8)
        Me.btTare.Name = "btTare"
        Me.btTare.Size = New System.Drawing.Size(128, 48)
        Me.btTare.TabIndex = 1
        Me.btTare.Text = "Tare"
        '
        'rtbInfo
        '
        Me.rtbInfo.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rtbInfo.Location = New System.Drawing.Point(16, 72)
        Me.rtbInfo.Name = "rtbInfo"
        Me.rtbInfo.ReadOnly = True
        Me.rtbInfo.Size = New System.Drawing.Size(464, 376)
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
        '
        'WeightScaleForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 464)
        Me.Controls.Add(Me.btOpenPort)
        Me.Controls.Add(Me.rtbInfo)
        Me.Controls.Add(Me.btTare)
        Me.Controls.Add(Me.btGetWeight)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "WeightScaleForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Weighting Scale "
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btGetWeight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btGetWeight.Click
        btGetWeight.Enabled = False
        AddInfo("Readig current weight")
        Weighting_Scale.GetWeight()
        WaitForData(timeOut)
    End Sub

    Private Sub btTare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btTare.Click
        Weighting_Scale.DoTare()
        AddInfo("Tared")
    End Sub

    Private Sub WeightScaleForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Weighting_Scale.OpenPort() Then
            Weighting_Scale.CommandFormat1(ambientMode)
            btGetWeight.Enabled = True
            btTare.Enabled = True
            btOpenPort.Text = "Close Port"
        Else
            AddInfo("Unable to connect to weighting scale")
            btGetWeight.Enabled = False
            btTare.Enabled = False
        End If
    End Sub
    Private Sub AddInfo(ByVal info As String)
        rtbInfo.AppendText(info + vbLf)
        rtbInfo.SelectionLength = rtbInfo.Text.Length
        rtbInfo.ScrollToCaret()
    End Sub
    Private Sub WaitForData(ByVal timeOut As Double)
        Dim time As DateTime = DateTime.Now
        Dim diff As Long = ((DateTime.Now.Ticks - time.Ticks) / 10000)
        While Not (Weighting_Scale.ValueUpdated) And (diff < timeOut) 'ms
            Application.DoEvents()
            diff = (DateTime.Now.Ticks - time.Ticks) / 10000
        End While
        If Not (Weighting_Scale.ValueUpdated) Then
            AddInfo("Timeout error! Unable to get weighting data")
        Else
            AddInfo("Current weight: " & Weighting_Scale.returnString + " mg")
        End If
        btGetWeight.Enabled = True
    End Sub

    Private Sub btOpenPort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btOpenPort.Click
        If btOpenPort.Text = "Open Port" Then
            If Weighting_Scale.OpenPort() Then
                Weighting_Scale.CommandFormat1(ambientMode)
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
        e.Cancel = True
        Hide()
    End Sub
End Class
