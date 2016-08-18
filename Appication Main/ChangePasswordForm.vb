Public Class ChangePasswordForm
    Inherits System.Windows.Forms.Form
    Public currentPass As String = ""
    Public newPassword As String = ""
    Public passed As Boolean = False
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
    Friend WithEvents tbOldPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tbNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tbConfirmNewPassword As System.Windows.Forms.TextBox
    Friend WithEvents btChangePassword As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbOldPassword = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.tbNewPassword = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbConfirmNewPassword = New System.Windows.Forms.TextBox
        Me.btChangePassword = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'tbOldPassword
        '
        Me.tbOldPassword.Location = New System.Drawing.Point(14, 32)
        Me.tbOldPassword.Name = "tbOldPassword"
        Me.tbOldPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.tbOldPassword.Size = New System.Drawing.Size(264, 20)
        Me.tbOldPassword.TabIndex = 0
        Me.tbOldPassword.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Old Password"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "New Password"
        '
        'tbNewPassword
        '
        Me.tbNewPassword.Location = New System.Drawing.Point(14, 88)
        Me.tbNewPassword.Name = "tbNewPassword"
        Me.tbNewPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.tbNewPassword.Size = New System.Drawing.Size(264, 20)
        Me.tbNewPassword.TabIndex = 2
        Me.tbNewPassword.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 23)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Confirm New Password"
        '
        'tbConfirmNewPassword
        '
        Me.tbConfirmNewPassword.Location = New System.Drawing.Point(14, 144)
        Me.tbConfirmNewPassword.Name = "tbConfirmNewPassword"
        Me.tbConfirmNewPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.tbConfirmNewPassword.Size = New System.Drawing.Size(264, 20)
        Me.tbConfirmNewPassword.TabIndex = 4
        Me.tbConfirmNewPassword.Text = ""
        '
        'btChangePassword
        '
        Me.btChangePassword.Location = New System.Drawing.Point(62, 200)
        Me.btChangePassword.Name = "btChangePassword"
        Me.btChangePassword.Size = New System.Drawing.Size(80, 48)
        Me.btChangePassword.TabIndex = 6
        Me.btChangePassword.Text = "Change Password"
        '
        'btCancel
        '
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Location = New System.Drawing.Point(150, 200)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(80, 48)
        Me.btCancel.TabIndex = 7
        Me.btCancel.Text = "Cancel"
        '
        'ChangePasswordForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btChangePassword)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.tbConfirmNewPassword)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbNewPassword)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbOldPassword)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "ChangePasswordForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change Password"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btChangePassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btChangePassword.Click
        If Me.tbNewPassword.Text = "" Then
            MessageBox.Show("Please insert new password.")
            Return
        ElseIf Me.tbConfirmNewPassword.Text = "" Then
            MessageBox.Show("Please insert confirm new password.")
            Return
        ElseIf Me.tbOldPassword.Text = "" Then
            MessageBox.Show("Please insert old password.")
            Return
        End If
        If Not (Me.tbNewPassword.Text = Me.tbConfirmNewPassword.Text) Then
            MessageBox.Show("New password not match, please make sure you key in the same new password.")
            Return
        End If
        If Not (Me.tbOldPassword.Text = currentPass) Then
            MessageBox.Show("Invalid old password.Please key in valid password.")
            Return
        End If
        passed = True
        newPassword = tbNewPassword.Text
        Close()
    End Sub

    Private Sub ChangePasswordForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tbNewPassword.Text = ""
        tbOldPassword.Text = ""
        tbConfirmNewPassword.Text = ""
        newPassword = ""
    End Sub
End Class
