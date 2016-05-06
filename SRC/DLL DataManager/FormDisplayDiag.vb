Public Class FormSelectPatternFile
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
    Friend WithEvents DirListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents FileListBox1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnApply As System.Windows.Forms.Button
    Friend WithEvents textFilename As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DirListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Me.DriveListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
        Me.FileListBox1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
        Me.textFilename = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnApply = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DirListBox1
        '
        Me.DirListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DirListBox1.IntegralHeight = False
        Me.DirListBox1.Location = New System.Drawing.Point(8, 88)
        Me.DirListBox1.Name = "DirListBox1"
        Me.DirListBox1.Size = New System.Drawing.Size(184, 240)
        Me.DirListBox1.TabIndex = 1
        '
        'DriveListBox1
        '
        Me.DriveListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DriveListBox1.Location = New System.Drawing.Point(8, 336)
        Me.DriveListBox1.Name = "DriveListBox1"
        Me.DriveListBox1.Size = New System.Drawing.Size(192, 23)
        Me.DriveListBox1.TabIndex = 2
        '
        'FileListBox1
        '
        Me.FileListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileListBox1.Location = New System.Drawing.Point(208, 88)
        Me.FileListBox1.Name = "FileListBox1"
        Me.FileListBox1.Pattern = "*.pat"
        Me.FileListBox1.Size = New System.Drawing.Size(264, 260)
        Me.FileListBox1.TabIndex = 3
        '
        'textFilename
        '
        Me.textFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textFilename.Location = New System.Drawing.Point(96, 24)
        Me.textFilename.Name = "textFilename"
        Me.textFilename.Size = New System.Drawing.Size(280, 22)
        Me.textFilename.TabIndex = 4
        Me.textFilename.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "FileName"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnApply)
        Me.GroupBox1.Controls.Add(Me.textFilename)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(472, 64)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'BtnApply
        '
        Me.BtnApply.Location = New System.Drawing.Point(384, 16)
        Me.BtnApply.Name = "BtnApply"
        Me.BtnApply.Size = New System.Drawing.Size(72, 40)
        Me.BtnApply.TabIndex = 6
        Me.BtnApply.Text = "Apply"
        '
        'FormSelectPatternFile
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 365)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.FileListBox1)
        Me.Controls.Add(Me.DriveListBox1)
        Me.Controls.Add(Me.DirListBox1)
        Me.Name = "FormSelectPatternFile"
        Me.Text = "PatternFile"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub FileListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileListBox1.SelectedIndexChanged
        textFilename.Text = FileListBox1.Path + "\" + FileListBox1.FileName
    End Sub

    Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox1.SelectedIndexChanged

        DirListBox1.Path = DriveListBox1.Drive
        DirListBox1_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub DirListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox1.SelectedIndexChanged
        Dim str As String = DirListBox1.Path
        Dim fil() As Char = DirListBox1.SelectedItem
        Dim Index As Integer = str.IndexOf(fil)

        If Index >= 0 Then
            Dim Dir As String = str.Remove(Index, str.Length - Index)

            FileListBox1.Path = Dir + DirListBox1.SelectedItem
            FileListBox1_SelectedIndexChanged(sender, e)
        End If

        
    End Sub

    Private Sub BtnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnApply.Click
        Try
            Dim str1 As String = FileListBox1.FileName
            Dim str As String() = str1.Split(".")
            IDSData.ParameterID.RecordID = str(0)
            IDSData.Admin.Folder.PatternPath = FileListBox1.Path

            If IDSData.IsSaveAsFile = True Then

                If SavePathFileName(textFilename.Text) = True Then
                    Me.Close()
                End If

            ElseIf IDSData.IsOpenFile = True Then

                If OpenPathFileName(textFilename.Text) = True Then
                    Me.Close()
                End If

            ElseIf IDSData.IsNewFile = True Then
                System.IO.File.Copy(FileListBox1.Path + "\FactoryDefault.Pat", FileListBox1.Path + "\Temp.Pat")
                If System.IO.File.Exists(textFilename.Text) = True Then
                    System.IO.File.Delete(textFilename.Text)
                End If
                System.IO.File.Move(FileListBox1.Path + "\temp.Pat", textFilename.Text)

                If OpenPathFileName(textFilename.Text) = True Then
                    Me.Close()
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

        
    End Sub

    Private Sub FormSelectPatternFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        System.IO.Directory.SetCurrentDirectory(IDSData.Admin.Folder.PatternPath)
        FileListBox1.ClearSelected()
        DirListBox1.SelectedItem = IDSData.Admin.Folder.PatternPath
        DirListBox1.Path = IDSData.Admin.Folder.PatternPath
        DirListBox1_SelectedIndexChanged(sender, e)


        If IDSData.IsSaveAsFile = True Then
            Me.Text = "Save as Pattern File"

        ElseIf IDSData.IsOpenFile = True Then
            Me.Text = "Open Pattern File"

        ElseIf IDSData.IsNewFile = True Then
            Me.Text = "New Pattern File"
        End If

    End Sub
End Class
