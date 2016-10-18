Imports System.IO
Public Class Login
    Inherits System.Windows.Forms.Form
    Public pass As String = ""
    Public passed As Boolean = False
    Public loginMode As Integer = 0 ' default is normal login
    Public logging As Boolean = False
    Dim chgPass As ChangePasswordForm = New ChangePasswordForm

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
    Friend WithEvents tbPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btLogin As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents btChangePassword As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbPassword = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btLogin = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
        Me.btChangePassword = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'tbPassword
        '
        Me.tbPassword.Location = New System.Drawing.Point(13, 64)
        Me.tbPassword.Name = "tbPassword"
        Me.tbPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.tbPassword.Size = New System.Drawing.Size(296, 20)
        Me.tbPassword.TabIndex = 0
        Me.tbPassword.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Password : "
        '
        'btLogin
        '
        Me.btLogin.Location = New System.Drawing.Point(76, 96)
        Me.btLogin.Name = "btLogin"
        Me.btLogin.TabIndex = 2
        Me.btLogin.Text = "Login"
        '
        'btCancel
        '
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Location = New System.Drawing.Point(172, 96)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.TabIndex = 3
        Me.btCancel.Text = "Cancel"
        '
        'btChangePassword
        '
        Me.btChangePassword.Location = New System.Drawing.Point(192, 8)
        Me.btChangePassword.Name = "btChangePassword"
        Me.btChangePassword.Size = New System.Drawing.Size(120, 23)
        Me.btChangePassword.TabIndex = 4
        Me.btChangePassword.Text = "Change Password"
        '
        'Login
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(322, 136)
        Me.Controls.Add(Me.btChangePassword)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btLogin)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbPassword)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btLogin.Click
        If Not (pass = tbPassword.Text) Then
            MessageBox.Show("Invalid password")
            Return
        End If
        passed = True
        Close()
    End Sub

    Private Sub btCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancel.Click

    End Sub

    Private Sub Login_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        logging = True
        pass = ""
        tbPassword.Text = ""
        tbPassword.Focus()
        passed = False
        If loginMode = 0 Or loginMode = 2 Then
            ReadFile()
        Else
            ReadSetupPassFile()
        End If
        tbPassword.Focus()
    End Sub
    Private Sub ReadFile()
        If Not (Directory.Exists("C:\Program Files\TIDS\")) Then
            Directory.CreateDirectory("C:\Program Files\TIDS\")
        End If
        Dim filePath As String = ""
        If loginMode = 0 Then
            filePath = "C:\Program Files\TIDS\p.dat"
        ElseIf loginMode = 2 Then
            filePath = "C:\Program Files\TIDS\wp.dat"
        End If
        If Not (File.Exists(filePath)) Then
            Dim writer As StreamWriter = New StreamWriter(filePath)
            writer.WriteLine("default")
            writer.Close()
        End If
        Dim reader As StreamReader = New StreamReader(filePath)
        pass = reader.ReadLine()
        reader.Close()
    End Sub
    Private Sub WritePassword(ByVal password As String)
        Dim filePath As String = "C:\Program Files\TIDS\p.dat"
        If loginMode = 2 Then
            filePath = "C:\Program Files\TIDS\wp.dat"
        End If
        Dim writer As StreamWriter = New StreamWriter(filePath)
        writer.WriteLine(password)
        writer.Close()
    End Sub
    Private Sub ReadSetupPassFile()
        Dim regKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\IDS\", True)
        pass = regKey.GetValue("")
        regKey.Close()
    End Sub

    Private Sub tbPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            If Not (pass = tbPassword.Text) Then
                MessageBox.Show("Invalid password")
                Return
            End If
            passed = True
            Close()
        End If
    End Sub

    Private Sub Login_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    End Sub

    Private Sub btChangePassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btChangePassword.Click
        chgPass.passed = False
        chgPass.currentPass = pass
        chgPass.newPassword = ""
        chgPass.ShowDialog()
        If chgPass.passed Then
            WritePassword(chgPass.newPassword)
            pass = chgPass.newPassword
            MessageBox.Show("Password changed successfully")
        End If
    End Sub

    Public Sub ShowChangePassword(ByVal toshow As Boolean)
        Me.btChangePassword.Visible = toshow
    End Sub
End Class
