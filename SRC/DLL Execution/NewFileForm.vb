Imports System.IO

Public Class NewFileForm
    Inherits System.Windows.Forms.Form
    Public teachingMode As String = "Vision"
    Public NewFileName As String = ""
    Private abortExit As Boolean = False
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
    Friend WithEvents tbFileName As System.Windows.Forms.TextBox
    Friend WithEvents btOK As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents cbTeachingMode As System.Windows.Forms.ComboBox
    Friend WithEvents lbExistingProject As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gbProjectName As System.Windows.Forms.GroupBox
    Friend WithEvents gbTeachingMode As System.Windows.Forms.GroupBox
    Friend WithEvents cbTeachMode As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbFileName = New System.Windows.Forms.TextBox
        Me.btOK = New System.Windows.Forms.Button
        Me.btCancel = New System.Windows.Forms.Button
        'Me.cbTeachingMode = New System.Windows.Forms.ComboBox
        Me.lbExistingProject = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.gbProjectName = New System.Windows.Forms.GroupBox
        Me.gbTeachingMode = New System.Windows.Forms.GroupBox
        Me.cbTeachMode = New System.Windows.Forms.ComboBox
        Me.gbProjectName.SuspendLayout()
        Me.gbTeachingMode.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbFileName
        '
        Me.tbFileName.Location = New System.Drawing.Point(8, 24)
        Me.tbFileName.Name = "tbFileName"
        Me.tbFileName.Size = New System.Drawing.Size(440, 20)
        Me.tbFileName.TabIndex = 0
        Me.tbFileName.Text = ""
        '
        'btOK
        '
        Me.btOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btOK.Location = New System.Drawing.Point(112, 368)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(96, 48)
        Me.btOK.TabIndex = 2
        Me.btOK.Text = "OK"
        '
        'btCancel
        '
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(256, 368)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(96, 48)
        Me.btCancel.TabIndex = 3
        Me.btCancel.Text = "Cancel"
        '
        'lbExistingProject
        '
        Me.lbExistingProject.Location = New System.Drawing.Point(9, 16)
        Me.lbExistingProject.Name = "lbExistingProject"
        Me.lbExistingProject.Size = New System.Drawing.Size(456, 199)
        Me.lbExistingProject.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Existing Project:"
        '
        'gbProjectName
        '
        Me.gbProjectName.Controls.Add(Me.tbFileName)
        Me.gbProjectName.Location = New System.Drawing.Point(8, 224)
        Me.gbProjectName.Name = "gbProjectName"
        Me.gbProjectName.Size = New System.Drawing.Size(456, 56)
        Me.gbProjectName.TabIndex = 6
        Me.gbProjectName.TabStop = False
        Me.gbProjectName.Text = "Please insert new project name"
        '
        'gbTeachingMode
        '
        Me.gbTeachingMode.Controls.Add(Me.cbTeachMode)
        Me.gbTeachingMode.Location = New System.Drawing.Point(8, 296)
        Me.gbTeachingMode.Name = "gbTeachingMode"
        Me.gbTeachingMode.Size = New System.Drawing.Size(456, 56)
        Me.gbTeachingMode.TabIndex = 7
        Me.gbTeachingMode.TabStop = False
        Me.gbTeachingMode.Text = "Teaching Mode"
        '
        'cbTeachMode
        '
        Me.cbTeachMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbTeachMode.Items.AddRange(New Object() {"Vision", "Needle"})
        Me.cbTeachMode.Location = New System.Drawing.Point(16, 24)
        Me.cbTeachMode.Name = "cbTeachMode"
        Me.cbTeachMode.Size = New System.Drawing.Size(432, 21)
        Me.cbTeachMode.TabIndex = 0
        '
        'NewFileForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(474, 424)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbExistingProject)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Controls.Add(Me.gbProjectName)
        Me.Controls.Add(Me.gbTeachingMode)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "NewFileForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Create Project Folder"
        Me.gbProjectName.ResumeLayout(False)
        Me.gbTeachingMode.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub NewFileForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cbTeachMode.SelectedIndex = 0
        Dim str() As String = Directory.GetDirectories("C:\IDS\Pattern_Dir")
        Dim i As Integer = 0
        Dim index As Integer = 0
        For i = 0 To str.Length - 1
            index = str(i).LastIndexOf("\")

            If index > 0 Then
                If str(i).Substring(index + 1).ToUpper = "SYSSWAPDATA" Then
                Else
                    lbExistingProject.Items.Add(str(i).Substring(index + 1))
                End If

            Else
                    lbExistingProject.Items.Add(str(i))
                End If
        Next
    End Sub

    Private Sub cbTeachingMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbTeachMode.SelectedIndexChanged
        If cbTeachMode.SelectedIndex >= 0 Then
            teachingMode = cbTeachMode.SelectedItem
        End If
    End Sub

    Private Sub btOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btOK.Click
        If tbFileName.Text = "" Then
            MessageBox.Show("Please insert File Name")
            Return
        End If
        abortExit = False
        Dim i As Integer = 0
        Dim fileExist As Boolean = False
        For i = 0 To lbExistingProject.Items.Count - 1
            If tbFileName.Text = lbExistingProject.Items(i) Then
                fileExist = True
            End If
        Next
        If fileExist Then
            If MessageBox.Show("Project exist! Are you sure you want to overwrite the existing project?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.No Then
                abortExit = True
                Return
            End If
        End If
        NewFileName = tbFileName.Text
    End Sub

    Private Sub NewFileForm_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If abortExit Then e.Cancel = True
    End Sub

    Private Sub lbExistingProject_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbExistingProject.SelectedIndexChanged
        If lbExistingProject.SelectedIndex >= 0 Then
            tbFileName.Text = lbExistingProject.SelectedItem
        End If
    End Sub
End Class
