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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormStartup))
        Me.BtnExit = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.loginPanel = New System.Windows.Forms.Panel
        Me.SuspendLayout()
        '
        'BtnExit
        '
        Me.BtnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnExit.Location = New System.Drawing.Point(1144, 16)
        Me.BtnExit.Name = "BtnExit"
        Me.BtnExit.Size = New System.Drawing.Size(112, 64)
        Me.BtnExit.TabIndex = 1
        Me.BtnExit.Text = "Exit"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(544, 248)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(194, 58)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Production"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(16, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(194, 58)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "Teach Program"
        Me.Button2.Visible = False
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
        Me.loginPanel.Location = New System.Drawing.Point(462, 328)
        Me.loginPanel.Name = "loginPanel"
        Me.loginPanel.Size = New System.Drawing.Size(368, 488)
        Me.loginPanel.TabIndex = 4
        '
        'FormStartup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1292, 1036)
        Me.Controls.Add(Me.loginPanel)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.BtnExit)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormStartup"
        Me.Text = "FormStartup"
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


    Private Sub BtnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExit.Click

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

        '   Xue Wen                                                     '
        '   Sometimes, application is still running at background.      '
        '   Need to find out because we have instantated may objects.   '

    End Sub

    Private Sub FormStartup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        KeyboardControl.GainControls()
        Taskbar.ShowTaskBar(False)
        'formlg.StartPosition = FormStartPosition.CenterScreen
        'formlg.Show()
        'formlg.Hide()
        formlg.StartPosition = FormStartPosition.Manual
        formlg.TopLevel = False
        formlg.Parent = Me
        loginPanel.Controls.Add(formlg)
        formlg.Show()
    End Sub

    Public frmLogin As New FormLogin

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        IDS.Data.Admin.User.RunApplication = "Operator"
        Call frmLogin.LoadProgrammerMaintenance()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        LoginName = "Application_Maintenance"
        formlg.StartPosition = FormStartPosition.CenterScreen
        formlg.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        LoginName = "Application_Programmer"
        formlg.StartPosition = FormStartPosition.CenterScreen
        formlg.ShowDialog()
    End Sub
End Class
