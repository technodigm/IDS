'Copyright (C) 2002 Microsoft Corporation
'All rights reserved.
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

'Requires the Trial or Release version of Visual Studio .NET Professional (or greater).

Option Strict On
Imports System.Windows.Forms
Imports System.Diagnostics ' For Process.Start
Imports System.IO

Class FileListView
    Inherits ListView
    Private strDirectory As String
    Private startupDir As String
    Private FileNameCrtrl As System.Windows.Forms.Control
    'Private components As System.ComponentModel.IContainer
    Private GraphicCtrl As System.Windows.Forms.Control
    Public Property initDir() As String
        Get
            Return startupDir
        End Get
        Set(ByVal Value As String)
            startupDir = Value
        End Set
    End Property
    ' This is the Class constructor.
    Public Sub New()
        ' Set the default View enumeration to Details.
        Me.View = View.Details

        ' Get images as icons for some of the common file types.
        Dim img As New ImageList
        '''''''SJ test
        With img.Images
            .Add(New Bitmap("C:\IDS\Appication Main\bin\DOC.BMP"))
            .Add(New Bitmap("C:\IDS\Appication Main\bin\EXE.BMP"))
        End With

        ' The Small and Large image lists for the ListView use the same set of
        ' images.
        Me.SmallImageList = img
        Me.LargeImageList = img

        ' Create the columns.
        With Columns
            .Add("Name", 100, HorizontalAlignment.Left)
            .Add("Size", 50, HorizontalAlignment.Right)
            .Add("Modified", 100, HorizontalAlignment.Right)
            '.Add("Attribute", 100, HorizontalAlignment.Left)
        End With


    End Sub

    ' Overrides the base class OnItemActivate event handler. Extends the base
    ' class implementation to run any .exe or file with an associated executable.
    Protected Overrides Sub OnItemActivate(ByVal ea As EventArgs)
        MyBase.OnItemActivate(ea)

        Dim lvi As ListViewItem
        For Each lvi In SelectedItems
            Try
                'Process.Start(Path.Combine(strDirectory, lvi.Text))

            Catch
                ' Do nothing. Just pass to Next lvi.
                Exit Try
            End Try
        Next lvi
    End Sub

    ' This subroutine is used to display a list of all files in the directory
    ' currently selected by the user from the custom TreeView control.
    Public Sub ShowFiles(ByVal strDirectory As String)
        ' Save the directory name as a field.
        Me.strDirectory = startupDir & strDirectory

        Items.Clear()

        Dim diDirectories As New DirectoryInfo(strDirectory)
        Dim afiFiles() As FileInfo

        Try
            ' Call the convenient GetFiles method to get an array of all files
            ' in the directory.
            afiFiles = diDirectories.GetFiles()
        Catch
            Return
        End Try

        Dim fi As FileInfo
        For Each fi In afiFiles
            ' Create ListViewItem.
            Dim lvi As New ListViewItem(fi.Name)
            '  If Path.GetExtension(fi.Name).ToUpper() = ".PTP" Then 'should be set by user
            If Path.GetExtension(fi.Name).ToUpper() = ".TXT" Then 'should be set by user

                ' Assign ImageIndex based on filename extension.
                ' Select Case Path.GetExtension(fi.Name).ToUpper()
                '    Case ".EXE"
                '   lvi.ImageIndex = 1
                '  Case ".PTP"
                lvi.ImageIndex = 0
                ' Case Else

                'End Select

                ' Add file length and last modified time sub-items.
                lvi.SubItems.Add(fi.Length.ToString("N0"))
                lvi.SubItems.Add(fi.LastWriteTime.ToString())

                ' Add attribute subitem.
                '  Dim strAttr As String = ""

                '  If (fi.Attributes And FileAttributes.Archive) <> 0 Then
                ' strAttr += "A"
                ' End If
                ' If (fi.Attributes And FileAttributes.Hidden) <> 0 Then
                '     strAttr += "H"
                ' End If
                'If (fi.Attributes And FileAttributes.ReadOnly) <> 0 Then
                '    strAttr += "R"
                'End If
                ' If (fi.Attributes And FileAttributes.System) <> 0 Then
                '    strAttr += "S"
                '' End If
                'lvi.SubItems.Add(strAttr)

                ' Add completed ListViewItem to FileListView.
                Items.Add(lvi)
            End If
        Next fi
    End Sub

    Protected Overrides Sub OnItemCheck(ByVal ice As System.Windows.Forms.ItemCheckEventArgs)

    End Sub


    Private Sub FileListView_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SelectedIndexChanged
        Dim i As Integer
        If Me.SelectedItems.Count = 1 Then
            ' Dim formup As System.Windows.Forms.Form = Me.FindForm
            ' Dim cntl As System.Windows.Forms.Control
            'For Each cntl In formup.Controls
            'If cntl.Name = "TextBoxFile" Then
            'cntl.Text = Me.SelectedItems(0).Text
            'End If
            '   Next
            FileNameCrtrl.Text = Me.SelectedItems(0).Text
            DrawPattern()
        End If
    End Sub

    Public Sub DrawPattern()
        Dim graph As System.Drawing.Graphics = GraphicCtrl.CreateGraphics()
        Dim color As New System.Drawing.Color
        graph.Clear(color.Black)
        Dim pen As New Pen(color.Blue)

        graph.DrawLine(pen, 0, 0, 100, 100)
        pen.Width = 3
        Dim rect As System.Drawing.Rectangle
        rect.X = 0
        rect.Y = 0
        rect.Height = GraphicCtrl.Height - 1
        rect.Width = GraphicCtrl.Width - 1
        Static i As Integer = 0
        graph.DrawImage(Me.LargeImageList.Images(i), New PointF(0, 0))
        i = i + 1
        If i > 1 Then
            i = 0
        End If
        ' graph.DrawRectangle(pen, rect)
    End Sub

    Private Sub FileListView_HandleCreated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.HandleCreated
        Dim formup As System.Windows.Forms.Form = Me.FindForm
        Dim cntl As System.Windows.Forms.Control
        For Each cntl In formup.Controls
            If cntl.Name = "TextBoxFile" Then
                FileNameCrtrl = cntl
            ElseIf cntl.Name = "PicturePattern" Then
                GraphicCtrl = cntl
            End If
        Next
    End Sub

End Class