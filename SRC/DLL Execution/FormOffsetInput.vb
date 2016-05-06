Public Class FormOffsetInput
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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents ButtonAccept As System.Windows.Forms.Button
    Friend WithEvents TextBoxX As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxY As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxZ As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TextBoxX = New System.Windows.Forms.TextBox
        Me.TextBoxY = New System.Windows.Forms.TextBox
        Me.TextBoxZ = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.ButtonAccept = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'TextBoxX
        '
        Me.TextBoxX.Location = New System.Drawing.Point(40, 72)
        Me.TextBoxX.Name = "TextBoxX"
        Me.TextBoxX.Size = New System.Drawing.Size(80, 20)
        Me.TextBoxX.TabIndex = 0
        Me.TextBoxX.Text = "0"
        '
        'TextBoxY
        '
        Me.TextBoxY.Location = New System.Drawing.Point(152, 72)
        Me.TextBoxY.Name = "TextBoxY"
        Me.TextBoxY.Size = New System.Drawing.Size(80, 20)
        Me.TextBoxY.TabIndex = 1
        Me.TextBoxY.Text = "0"
        '
        'TextBoxZ
        '
        Me.TextBoxZ.Location = New System.Drawing.Point(264, 72)
        Me.TextBoxZ.Name = "TextBoxZ"
        Me.TextBoxZ.Size = New System.Drawing.Size(80, 20)
        Me.TextBoxZ.TabIndex = 2
        Me.TextBoxZ.Text = "0"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "X"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(152, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Y"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(272, 40)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(16, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Z"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Location = New System.Drawing.Point(152, 136)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(96, 32)
        Me.ButtonCancel.TabIndex = 7
        Me.ButtonCancel.Text = "Cancel"
        '
        'ButtonAccept
        '
        Me.ButtonAccept.Location = New System.Drawing.Point(280, 136)
        Me.ButtonAccept.Name = "ButtonAccept"
        Me.ButtonAccept.Size = New System.Drawing.Size(88, 32)
        Me.ButtonAccept.TabIndex = 8
        Me.ButtonAccept.Text = "Accept"
        '
        'FormOffsetInput
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(400, 206)
        Me.Controls.Add(Me.ButtonAccept)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBoxZ)
        Me.Controls.Add(Me.TextBoxY)
        Me.Controls.Add(Me.TextBoxX)
        Me.Name = "FormOffsetInput"
        Me.Text = "Form Offset Input"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FormOffsetInput_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub


    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    Private Sub ButtonAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAccept.Click

        If IsNumeric(TextBoxX.Text) And IsNumeric(TextBoxY.Text) And IsNumeric(TextBoxZ.Text) Then
            Me.DialogResult = DialogResult.OK
        Else
            MyMsgBox("Please input valid number", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Input number invalid")
        End If
    End Sub

    Private Sub TextBoxX_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxX.TextChanged

    End Sub

    Private Sub TextBoxY_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxY.TextChanged

    End Sub

    Private Sub TextBoxZ_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBoxZ.TextChanged

    End Sub
End Class
