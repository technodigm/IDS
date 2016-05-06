Imports System.IO.File
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
    'Friend WithEvents DirListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    'Friend WithEvents DriveListBox1 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    'Friend WithEvents FileListBox1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents textFilename As System.Windows.Forms.TextBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents BtnOK As System.Windows.Forms.Button
    Friend WithEvents BtnCancel As System.Windows.Forms.Button
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents DirListBox2 As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
    Friend WithEvents DriveListBox2 As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
    Friend WithEvents FileListBox2 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormSelectPatternFile))
        Me.textFilename = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.BtnOK = New System.Windows.Forms.Button
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.BtnCancel = New System.Windows.Forms.Button
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.DirListBox2 = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
        Me.DriveListBox2 = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
        Me.FileListBox2 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
        Me.SuspendLayout()
        '
        'textFilename
        '
        Me.textFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.textFilename.Location = New System.Drawing.Point(8, 336)
        Me.textFilename.Name = "textFilename"
        Me.textFilename.Size = New System.Drawing.Size(264, 20)
        Me.textFilename.TabIndex = 4
        Me.textFilename.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 312)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "FilePathName :"
        '
        'BtnOK
        '
        Me.BtnOK.Enabled = False
        Me.BtnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnOK.Location = New System.Drawing.Point(288, 320)
        Me.BtnOK.Name = "BtnOK"
        Me.BtnOK.Size = New System.Drawing.Size(80, 40)
        Me.BtnOK.TabIndex = 6
        Me.BtnOK.Text = "Ok"
        '
        'PictureBox1
        '
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(488, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(296, 352)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'BtnCancel
        '
        Me.BtnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCancel.Location = New System.Drawing.Point(384, 320)
        Me.BtnCancel.Name = "BtnCancel"
        Me.BtnCancel.Size = New System.Drawing.Size(80, 40)
        Me.BtnCancel.TabIndex = 8
        Me.BtnCancel.Text = "Cancel"
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(488, 8)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(296, 352)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 9
        Me.PictureBox2.TabStop = False
        '
        'DirListBox2
        '
        Me.DirListBox2.IntegralHeight = False
        Me.DirListBox2.Location = New System.Drawing.Point(8, 48)
        Me.DirListBox2.Name = "DirListBox2"
        Me.DirListBox2.Size = New System.Drawing.Size(256, 256)
        Me.DirListBox2.TabIndex = 10
        '
        'DriveListBox2
        '
        Me.DriveListBox2.Location = New System.Drawing.Point(8, 16)
        Me.DriveListBox2.Name = "DriveListBox2"
        Me.DriveListBox2.Size = New System.Drawing.Size(256, 21)
        Me.DriveListBox2.TabIndex = 11
        '
        'FileListBox2
        '
        Me.FileListBox2.Location = New System.Drawing.Point(272, 16)
        Me.FileListBox2.Name = "FileListBox2"
        Me.FileListBox2.Pattern = "*.txt"
        Me.FileListBox2.Size = New System.Drawing.Size(208, 290)
        Me.FileListBox2.TabIndex = 12
        '
        'FormSelectPatternFile
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(800, 373)
        Me.Controls.Add(Me.FileListBox2)
        Me.Controls.Add(Me.DriveListBox2)
        Me.Controls.Add(Me.DirListBox2)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.BtnCancel)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.BtnOK)
        Me.Controls.Add(Me.textFilename)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormSelectPatternFile"
        Me.Text = "PatternFile"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

#End Region

    '''''''''''''''''
    '   Xue Wen     '
    '''''''''''''''''
    'Save CDROM, Floppy Disk
    Dim CD_Drive, Flop_Drive(5) As String
    Dim Flop_Count As Integer = 0

    Private Sub DriveListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox2.SelectedIndexChanged

        '''''''''''''''''
        '   Xue Wen     '
        '''''''''''''''''
        'Check "CDROM" and "Floppy Disk", then do nothing.
        'Note: If there is no CD and Disk, system will prompt error and close the application.
        '
        '   Comment: Find another => check CD inserted or not?
        Dim str As String
        Dim str_Drive() As String

        str = DriveListBox2.Drive
        str_Drive = Split(str, ":")

        If (UCase(str_Drive(0)) = CD_Drive) Then
            Exit Sub
        ElseIf (Flop_Count >= 1) Then
            Dim count As Integer

            For count = 0 To (Flop_Count - 1)
                If (UCase(str_Drive(0)) = Flop_Drive(count)) Then
                    Exit Sub
                End If
            Next
        End If

        DirListBox2.Path = DriveListBox2.Drive
        DirListBox2_SelectedIndexChanged(sender, e)

    End Sub

    Private Sub DirListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox2.SelectedIndexChanged
        FileListDisplay()
    End Sub

    Private Sub FileListDisplay()
        Dim PathFileName As String

        'Me.Text = DirListBox1.Path + " ____ " + DirListBox1.SelectedItem + " ____" _
        '+ DirListBox1.DirList(DirListBox1.SelectedIndex) + " _____"

        FileListBox2.Path = DirListBox2.Path
        If FileListBox2.Items.Contains(DirListBox2.SelectedItem + ".txt") = True Then

            PathFileName = FileListBox2.Path + "\" + DirListBox2.SelectedItem + ".bmp"

            If System.IO.File.Exists(PathFileName) = True Then
                PictureBox2.SendToBack()
                PictureBox1.Image = PictureBox1.Image.FromFile(PathFileName)
            End If
        Else
            PictureBox2.BringToFront()
        End If
    End Sub

    Private Sub FileListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileListBox2.SelectedIndexChanged
        textFilename.Text = FileListBox2.Path + "\" + FileListBox2.FileName
        If System.IO.File.Exists(textFilename.Text) = True Then
            BtnOK.Enabled = True
        Else
            BtnOK.Enabled = False
        End If

        Dim str1 As String = textFilename.Text
        Dim str As String() = str1.Split(".")
        Dim ImagefileName As String = str(0) + ".bmp"

        If System.IO.File.Exists(ImagefileName) = True Then
            PictureBox2.SendToBack()
            PictureBox1.Image = PictureBox1.Image.FromFile(ImagefileName)
        Else
            PictureBox2.BringToFront()
        End If

    End Sub

    'Private Sub FileListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileListBox2.SelectedIndexChanged
    '    textFilename.Text = FileListBox2.Path + "\" + FileListBox2.FileName


    '    Dim str1 As String = textFilename.Text
    '    Dim str As String() = str1.Split(".")
    '    Dim ImagefileName As String = str(0) + ".bmp"

    '    If System.IO.File.Exists(ImagefileName) = True Then
    '        PictureBox2.SendToBack()
    '        PictureBox1.Image = PictureBox1.Image.FromFile(ImagefileName)
    '    Else
    '        PictureBox2.BringToFront()
    '    End If

    'End Sub

    'Private Sub DriveListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DriveListBox2.SelectedIndexChanged

    '    DirListBox2.Path = DriveListBox2.Drive
    '    DirListBox2_SelectedIndexChanged(sender, e)
    'End Sub

    'Private Sub DirListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirListBox2.SelectedIndexChanged
    '    Dim PathFileName As String

    '    'Me.Text = DirListBox1.Path + " ____ " + DirListBox1.SelectedItem + " ____" _
    '    '+ DirListBox1.DirList(DirListBox1.SelectedIndex) + " _____"


    '    FileListBox2.Path = DirListBox2.Path
    '    If FileListBox2.Items.Contains(DirListBox2.SelectedItem + ".ptp") = True Then

    '        PathFileName = FileListBox2.Path + "\" + DirListBox2.SelectedItem + ".bmp"

    '        If System.IO.File.Exists(PathFileName) = True Then
    '            PictureBox2.SendToBack()
    '            PictureBox1.Image = PictureBox1.Image.FromFile(PathFileName)
    '        End If
    '    Else
    '        PictureBox2.BringToFront()
    '    End If

    'End Sub
    Function Path() As String
        Return TempFileName2
    End Function
    Dim str1 As String
    Private Sub BtnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOK.Click
        ' open file
        Try
            Me.Close()
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub

    Private Sub FormSelectPatternFile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '''''''''''''''''
        '   Xue Wen     '
        '''''''''''''''''

        'Check all drive and drive types
        Dim oFSO As Scripting.FileSystemObject
        Dim oDrive As Scripting.Drive

        oFSO = New Scripting.FileSystemObject

        CD_Drive = ""

        For Each oDrive In oFSO.Drives
            If oDrive.DriveType = Scripting.DriveTypeConst.CDRom Then
                If (oDrive.IsReady = False) Then
                    CD_Drive = UCase(oDrive.DriveLetter)
                End If
            ElseIf oDrive.DriveType = Scripting.DriveTypeConst.Removable Then
                If (oDrive.IsReady = False) Then
                    Flop_Drive(Flop_Count) = UCase(oDrive.DriveLetter)

                    Flop_Count += 1
                End If
            End If
        Next

        ' initialize to the predefine directory
        'System.IO.Directory.SetCurrentDirectory(IDSData.Admin.Folder.PatternPath)
        'FileListBox1.ClearSelected()
        'DirListBox1.SelectedItem = IDSData.Admin.Folder.PatternPath
        'DirListBox1.Path = IDSData.Admin.Folder.PatternPath
        'DirListBox1_SelectedIndexChanged(sender, e)
        DirListBox2.Path = "C:\IDS\Pattern_Dir"

    End Sub

    Private Sub BtnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCancel.Click
        TempFileName2 = Nothing
        Me.Close()
    End Sub

    Dim DirName As String
    Dim TempFileName As String
    Dim TempFileName2 As String
    Private Sub DirListBox2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DirListBox2.DoubleClick
        DirName = DirListBox2.Path
        TempFileName = DirName.Substring(DirName.LastIndexOf("\") + 1)
        TempFileName2 = DirName + "\" + TempFileName + ".txt"
        textFilename.Text = DirName
        If System.IO.File.Exists(TempFileName2) = True Then
            BtnOK.Enabled = True
            textFilename.Text = TempFileName2
        Else
            BtnOK.Enabled = False
        End If

        Dim TempFileName3 As String = DirName + "\" + TempFileName + ".bmp"
        If System.IO.File.Exists(TempFileName3) = True Then
            PictureBox2.SendToBack()
            PictureBox1.Image = PictureBox1.Image.FromFile(TempFileName3)
        Else
            PictureBox2.BringToFront()
        End If
        FileListDisplay()
    End Sub

    Private Sub FileListBox2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles FileListBox2.DoubleClick
        Try
            Me.Close()
        Catch ex As Exception
            ExceptionDisplay(ex)
        End Try
    End Sub
End Class
