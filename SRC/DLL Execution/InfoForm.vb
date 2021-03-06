Imports DLL_Export_Service
Public Class InfoForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call

    'End Sub

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
    Friend WithEvents tbInfo As System.Windows.Forms.TextBox
    Friend WithEvents btOK As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents btIgnore As System.Windows.Forms.Button
    Friend WithEvents btStopBuzzer As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbInfo = New System.Windows.Forms.TextBox
        Me.btOK = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
        Me.btIgnore = New System.Windows.Forms.Button
        Me.btStopBuzzer = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'tbInfo
        '
        Me.tbInfo.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.tbInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbInfo.Location = New System.Drawing.Point(9, 8)
        Me.tbInfo.Multiline = True
        Me.tbInfo.Name = "tbInfo"
        Me.tbInfo.ReadOnly = True
        Me.tbInfo.Size = New System.Drawing.Size(560, 144)
        Me.tbInfo.TabIndex = 0
        Me.tbInfo.TabStop = False
        Me.tbInfo.Text = ""
        Me.tbInfo.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btOK
        '
        Me.btOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btOK.Location = New System.Drawing.Point(161, 168)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(96, 40)
        Me.btOK.TabIndex = 1
        Me.btOK.Text = "OK"
        '
        'btCancel
        '
        Me.btCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(321, 168)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(96, 40)
        Me.btCancel.TabIndex = 2
        Me.btCancel.Text = "Cancel"
        '
        'btIgnore
        '
        Me.btIgnore.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btIgnore.DialogResult = System.Windows.Forms.DialogResult.Ignore
        Me.btIgnore.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btIgnore.Location = New System.Drawing.Point(472, 168)
        Me.btIgnore.Name = "btIgnore"
        Me.btIgnore.Size = New System.Drawing.Size(96, 40)
        Me.btIgnore.TabIndex = 3
        Me.btIgnore.Text = "Ignore"
        Me.btIgnore.Visible = False
        '
        'btStopBuzzer
        '
        Me.btStopBuzzer.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btStopBuzzer.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btStopBuzzer.Location = New System.Drawing.Point(8, 160)
        Me.btStopBuzzer.Name = "btStopBuzzer"
        Me.btStopBuzzer.Size = New System.Drawing.Size(72, 48)
        Me.btStopBuzzer.TabIndex = 4
        Me.btStopBuzzer.Text = "Stop Buzzer"
        Me.btStopBuzzer.Visible = False
        '
        'InfoForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(578, 216)
        Me.Controls.Add(Me.btStopBuzzer)
        Me.Controls.Add(Me.btIgnore)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Controls.Add(Me.tbInfo)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "InfoForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Info"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region
    Delegate Sub OKClicked()
    Delegate Sub CancelClicked()
    Delegate Sub StopBuzzerClicked()
    Public okClickedDelegate As OKClicked
    Public cancelClickedDelegate As CancelClicked
    Public isError As Boolean = False
    'Public stopBuzzerClickedDelegate As StopBuzzerClicked
    Public Sub New(Optional ByVal isError As Boolean = False)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.isError = isError
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btOK.Click
        If Not (okClickedDelegate Is Nothing) Then
            okClickedDelegate.Invoke()
            Return
        End If
        Close()
    End Sub

    Private Sub btCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancel.Click
        If Not (cancelClickedDelegate Is Nothing) Then
            cancelClickedDelegate.Invoke()
        End If
        Close()
    End Sub

    Public Function SetMessage(ByVal message As String)
        tbInfo.Text = message
    End Function

    Public Function AddNewLine(ByVal str As String)
        tbInfo.AppendText(Environment.NewLine & str)
    End Function

    Private Sub InfoForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If isError Then
            Me.btStopBuzzer.Visible = True
            ' IDS.Devices.DIO.DIO.TurnOnTowerSiren()
            IDS.Devices.DIO.DIO.TurnOnTowerLightRed()
        End If
    End Sub
    Public Sub SetOKButtonText(ByVal text As String)
        btOK.Text = text
    End Sub
    Public Sub SetCancelButtonText(ByVal text As String)
        btCancel.Text = text
    End Sub
    Public Sub OkOnly()
        btOK.Location = New Point((Me.Size.Width - (btOK.Size.Width)) / 2, btOK.Location.Y)
        btCancel.Hide()
    End Sub
    Public Sub SetTitle(ByVal title As String)
        Me.Text = title
    End Sub
    Public Sub ShowIgnoreButton()
        btCancel.Location = New Point(Me.Size.Width / 2 - btCancel.Size.Width / 2, btCancel.Location.Y)
        btOK.Location = New Point(btCancel.Location.X - btOK.Size.Width - 20, btOK.Location.Y)
        btIgnore.Location = New Point(btCancel.Location.X + btIgnore.Size.Width + 20, btIgnore.Location.Y)
        btIgnore.Visible = True

    End Sub

    Private Sub btIgnore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btIgnore.Click
        Close()
    End Sub

    Private Sub btStopBuzzer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btStopBuzzer.Click
        'If Not (stopBuzzerClickedDelegate Is Nothing) Then
        '    stopBuzzerClickedDelegate.Invoke()
        'End If
        IDS.Devices.DIO.DIO.TurnOffTowerSiren()
        'IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
    End Sub

    Private Sub InfoForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If isError Then
            IDS.Devices.DIO.DIO.TurnOffTowerSiren()
            IDS.Devices.DIO.DIO.TurnOffTowerLightRed()
        End If
    End Sub
End Class
