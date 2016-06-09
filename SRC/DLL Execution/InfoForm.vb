Public Class InfoForm
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
    Friend WithEvents tbInfo As System.Windows.Forms.TextBox
    Friend WithEvents btOK As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbInfo = New System.Windows.Forms.TextBox
        Me.btOK = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
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
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(321, 168)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(96, 40)
        Me.btCancel.TabIndex = 2
        Me.btCancel.Text = "Cancel"
        '
        'InfoForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(578, 216)
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
    Public okClickedDelegate As OKClicked
    Public cancelClickedDelegate As CancelClicked
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
End Class
