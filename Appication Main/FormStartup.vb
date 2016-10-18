Imports DLL_Export_Service
Imports DLL_DataManager

Public Class FormStartup
    Inherits System.Windows.Forms.Form

#Region " APIs "
    Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End Region

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
    Friend WithEvents BtnExit As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents loginPanel As System.Windows.Forms.Panel
    Friend WithEvents btSetup As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btAccessWindows As System.Windows.Forms.Button
    Friend WithEvents btExitSoftware As System.Windows.Forms.Button
    Friend WithEvents ExitButtonTimer As System.Windows.Forms.Timer
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents AccessWindowsTimer As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormStartup))
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.loginPanel = New System.Windows.Forms.Panel
        Me.btSetup = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btAccessWindows = New System.Windows.Forms.Button
        Me.btExitSoftware = New System.Windows.Forms.Button
        Me.ExitButtonTimer = New System.Windows.Forms.Timer(Me.components)
        Me.Button3 = New System.Windows.Forms.Button
        Me.AccessWindowsTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnExit.Location = New System.Drawing.Point(1144, 16)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(112, 64)
        Me.BtnExit.TabIndex = 1
        Me.BtnExit.Text = "Shut Down"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(549, 264)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(194, 58)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Production"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(549, 352)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(194, 58)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "Teach Program"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(-24, 920)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(1340, 119)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 3
        Me.PictureBox2.TabStop = False
        '
        'loginPanel
        '
        Me.loginPanel.Location = New System.Drawing.Point(56, 408)
        Me.loginPanel.Name = "loginPanel"
        Me.loginPanel.Size = New System.Drawing.Size(368, 488)
        Me.loginPanel.TabIndex = 4
        '
        'btSetup
        '
        Me.btSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btSetup.Location = New System.Drawing.Point(549, 448)
        Me.btSetup.Name = "btSetup"
        Me.btSetup.Size = New System.Drawing.Size(194, 58)
        Me.btSetup.TabIndex = 5
        Me.btSetup.Text = "Setup Program"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 72.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(177, Byte))
        Me.Label1.Location = New System.Drawing.Point(246, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(800, 128)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "IDS2000"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btAccessWindows
        '
        Me.btAccessWindows.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btAccessWindows.Location = New System.Drawing.Point(56, 16)
        Me.btAccessWindows.Name = "btAccessWindows"
        Me.btAccessWindows.Size = New System.Drawing.Size(112, 64)
        Me.btAccessWindows.TabIndex = 7
        Me.btAccessWindows.Text = "Access Windows"
        Me.btAccessWindows.Visible = False
        '
        'btExitSoftware
        '
        Me.btExitSoftware.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btExitSoftware.Location = New System.Drawing.Point(1144, 176)
        Me.btExitSoftware.Name = "btExitSoftware"
        Me.btExitSoftware.Size = New System.Drawing.Size(112, 64)
        Me.btExitSoftware.TabIndex = 8
        Me.btExitSoftware.Text = "Exit IDS"
        Me.btExitSoftware.Visible = False
        '
        'ExitButtonTimer
        '
        Me.ExitButtonTimer.Interval = 30000
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(1144, 96)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(112, 64)
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Restart"
        '
        'AccessWindowsTimer
        '
        Me.AccessWindowsTimer.Enabled = True
        Me.AccessWindowsTimer.Interval = 50
        '
        'FormStartup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1292, 1036)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btExitSoftware)
        Me.Controls.Add(Me.btAccessWindows)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btSetup)
        Me.Controls.Add(Me.loginPanel)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormStartup"
        Me.Text = "IDS2000"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Form Events "
#Region " ... OnHandleCreated "
    Protected Overrides Sub OnHandleCreated(ByVal e As System.EventArgs)
        '---------------------------------------------------------------------------------
        '     Date          Developer                      Code Change       
        '  ---------- -------------------- -----------------------------------------------
        '  12/10/2005 G Gilbert            Original code
        '---------------------------------------------------------------------------------

        '---------------------------------------------------------------------------------
        ' Get all of the active processes for the application
        '---------------------------------------------------------------------------------
        Dim CurrentProcesses() As Process
        CurrentProcesses = Process.GetProcessesByName _
                           (Process.GetCurrentProcess.ProcessName)

        '---------------------------------------------------------------------------------
        ' If there is no previous copy of the application running, bail
        '---------------------------------------------------------------------------------
        If CurrentProcesses.GetUpperBound(0) <= 0 Then
            Exit Sub
        End If

        '---------------------------------------------------------------------------------
        ' Make the prior instance the active window
        '---------------------------------------------------------------------------------
        Dim i As Integer
        Dim Result As Long
        Dim ProcessHandle As Long
        For i = 0 To CurrentProcesses.GetUpperBound(0)
            ProcessHandle = CurrentProcesses(i).MainWindowHandle.ToInt32
            If ProcessHandle <> 0 Then
                '** Restore and activate the copy already running
                Result = OpenIcon(ProcessHandle)
                Result = SetForegroundWindow(ProcessHandle)
            End If
        Next

        '---------------------------------------------------------------------------------
        ' Terminate this instance
        '---------------------------------------------------------------------------------
        End

    End Sub
#End Region
#End Region

    Shared formlg As New FormLogin
    Dim loginForm As New Login


    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click
        If MessageBox.Show("Are you sure you want to shut down the PC?", "Warning", MessageBoxButtons.YesNo) = DialogResult.No Then
            Return
        End If
        formlg.Dispose()
        KeyboardControl.ReleaseControls()
        Taskbar.ShowTaskBar(True)
        Shell("shutdown -s -t 00")
        '   Xue Wen                                     '
        '   Testing (Kill the application directly)     '
        Dim procRunning() As Process
        procRunning = Process.GetProcesses

        For Each procs As Process In procRunning
            Try
                If procs.ProcessName.Equals("IDS Application") Then
                    procs.Kill()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next

        '   Xue Wen                                                     '
        '   Sometimes, application is still running at background.      '
        '   Need to find out because we have instantated may objects.   '

    End Sub

    Private Sub FormStartup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        KeyboardControl.GainControls()
        'Taskbar.ShowTaskBar(False)
        'formlg.StartPosition = FormStartPosition.CenterScreen
        'formlg.Show()
        'formlg.Hide()

        formlg.StartPosition = FormStartPosition.Manual
        formlg.TopLevel = False
        formlg.Parent = Me
        loginPanel.Controls.Add(formlg)
        formlg.Show()
        formlg.Hide()
    End Sub

    Public frmLogin As New FormLogin

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Cursor Is Cursors.WaitCursor Then Exit Sub
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        IDS.Data.Admin.User.RunApplication = "Operator"
        Cursor = Cursors.WaitCursor
        KeyboardControl.BlockWinKey = True
        ExitButtonTimer.Stop()
        ExitButtonTimer.Enabled = False
        btExitSoftware.Visible = False
        Me.AccessWindowsTimer.Enabled = False
        Call frmLogin.LoadProgrammerMaintenance()
        Me.AccessWindowsTimer.Enabled = True
        Cursor = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoginName = "Application_Maintenance"
        formlg.StartPosition = FormStartPosition.CenterScreen
        formlg.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        LoginName = "Application_Programmer"
        'formlg.StartPosition = FormStartPosition.CenterScreen
        'formlg.ShowDialog()
        If loginForm.logging Then Return
        loginForm.ShowChangePassword(True)
        loginForm.logging = True
        loginForm.loginMode = 0
        loginForm.ShowDialog()
        If loginForm.passed Then
            If Cursor Is Cursors.WaitCursor Then
                Return
            End If
            IDS.Data.ParameterID.RecordID = "FactoryDefault"
            IDSData.Admin.Folder.FileExtension = "Pat"
            IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
            IDS.Data.OpenData()
            IDS.Data.Admin.User.RunApplication = "Programmer"
            Cursor = Cursors.WaitCursor
            KeyboardControl.BlockWinKey = True
            ExitButtonTimer.Stop()
            ExitButtonTimer.Enabled = False
            btExitSoftware.Visible = False
            Me.AccessWindowsTimer.Enabled = False
            Call frmLogin.LoadProgrammerMaintenance()
            Me.AccessWindowsTimer.Enabled = True
            Cursor = Cursors.Default
            Me.BringToFront()
        End If
        loginForm.logging = False
    End Sub

    Private Sub btSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btSetup.Click
        If loginForm.logging Then Return
        loginForm.logging = True
        loginForm.ShowChangePassword(False)
        loginForm.loginMode = 1
        loginForm.ShowDialog()
        If loginForm.passed Then
            IDS.Data.ParameterID.RecordID = "FactoryDefault"
            IDSData.Admin.Folder.FileExtension = "Pat"
            IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
            IDS.Data.OpenData()
            IDS.Data.Admin.User.RunApplication = "System"
            KeyboardControl.BlockWinKey = True
            ExitButtonTimer.Stop()
            ExitButtonTimer.Enabled = False
            btExitSoftware.Visible = False
            Me.AccessWindowsTimer.Enabled = False
            Call frmLogin.LoadProgrammerMaintenance()
            Me.AccessWindowsTimer.Enabled = True
        End If
        loginForm.logging = False
    End Sub

    Private Sub FormStartup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Taskbar.ShowTaskBar(True)
    End Sub

    Private Sub btAccessWindows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btAccessWindows.Click
        If loginForm.logging Then Return
        loginForm.logging = True
        loginForm.ShowChangePassword(True)
        loginForm.loginMode = 2
        loginForm.ShowDialog()
        If loginForm.passed Then
            Taskbar.ShowTaskBar(True)
            KeyboardControl.BlockWinKey = False
            btExitSoftware.Visible = True
            ExitButtonTimer.Enabled = True
        End If
        loginForm.logging = False
    End Sub

    Private Sub btExitSoftware_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btExitSoftware.Click
        If MessageBox.Show("Are you sure you want to exit this software?", "Warning", MessageBoxButtons.YesNo) = DialogResult.No Then
            Return
        End If
        formlg.Dispose()
        KeyboardControl.ReleaseControls()
        Taskbar.ShowTaskBar(True)
        '   Xue Wen                                     '
        '   Testing (Kill the application directly)     '
        Dim procRunning() As Process
        procRunning = Process.GetProcesses

        For Each procs As Process In procRunning
            Try
                If procs.ProcessName.Equals("IDS Application") Then
                    procs.Kill()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
    End Sub

    Private Sub ExitButtonTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButtonTimer.Tick
        Me.btExitSoftware.Visible = False
        ExitButtonTimer.Stop()
        ExitButtonTimer.Enabled = False
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If MessageBox.Show("Are you sure you want to restart the PC?", "Warning", MessageBoxButtons.YesNo) = DialogResult.No Then
            Return
        End If
        formlg.Dispose()
        KeyboardControl.ReleaseControls()
        Taskbar.ShowTaskBar(True)
        Shell("shutdown -r -t 00")
        '   Xue Wen                                     '
        '   Testing (Kill the application directly)     '
        Dim procRunning() As Process
        procRunning = Process.GetProcesses

        For Each procs As Process In procRunning
            Try
                If procs.ProcessName.Equals("IDS Application") Then
                    procs.Kill()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
    End Sub
    Dim GPressed As Boolean = False
    Dim TPressed As Boolean = False
    Dim WPressed As Boolean = False
    Dim showingLoginWindows As Boolean = False
    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        If showingLoginWindows Then Return
        If e.KeyCode = Keys.G Then
            GPressed = True
        ElseIf e.KeyCode = Keys.T Then
            TPressed = True
        ElseIf e.KeyCode = Keys.W Then
            WPressed = True
        End If
        If GPressed And TPressed And WPressed Then
            showingLoginWindows = True
            If loginForm.logging Then Return
            loginForm.logging = True
            loginForm.ShowChangePassword(True)
            loginForm.loginMode = 2
            loginForm.ShowDialog()
            If loginForm.passed Then
                Taskbar.ShowTaskBar(True)
                KeyboardControl.BlockWinKey = False
                btExitSoftware.Visible = True
                ExitButtonTimer.Enabled = True
            End If
            loginForm.logging = False
            showingLoginWindows = False
        End If
    End Sub

    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.G Then
            GPressed = False
        ElseIf e.KeyCode = Keys.T Then
            TPressed = False
        ElseIf e.KeyCode = Keys.W Then
            WPressed = False
        End If
    End Sub

    Private Sub AccessWindowsTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccessWindowsTimer.Tick
        If KeyboardControl.AccessWindow Then
            AccessWindowsTimer.Stop()
            AccessWindowsTimer.Enabled = False
            If showingLoginWindows Then Return
            showingLoginWindows = True
            If loginForm.logging Then Return
            loginForm.logging = True
            loginForm.ShowChangePassword(True)
            loginForm.loginMode = 2
            loginForm.ShowDialog()
            If loginForm.passed Then
                Taskbar.ShowTaskBar(True)
                KeyboardControl.BlockWinKey = False
                btExitSoftware.Visible = True
                ExitButtonTimer.Enabled = True
            End If
            loginForm.logging = False
            showingLoginWindows = False
            AccessWindowsTimer.Enabled = True
        End If
    End Sub
End Class
