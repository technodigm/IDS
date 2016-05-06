'Copyright (C) 2002 Microsoft Corporation
'All rights reserved.
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

'Requires the Trial or Release version of Visual Studio .NET Professional (or greater).

Option Strict On

Imports System.IO

Class DirectoryTreeView
    Inherits TreeView
    Private startupDir As String
    Public Property initDir() As String
        Get
            Return startupDir
        End Get
        Set(ByVal Value As String)
            startupDir = Value
        End Set
    End Property
    ' This is the Class constructor.
    Public Sub New(ByVal initDir As String)
        ' Make a little more room for long directory names.
        Me.Width *= 2
        startupDir = initDir
        ' Get images for tree.
        Me.ImageList = New ImageList
        With Me.ImageList.Images
            .Add(New Bitmap("C:\IDS\Appication Main\bin\35FLOPPY.BMP"))
            .Add(New Bitmap("C:\IDS\Appication Main\bin\CLSDFOLD.BMP"))
            .Add(New Bitmap("C:\IDS\Appication Main\bin\OPENFOLD.BMP"))
        End With

        ' Construct tree.
        RefreshTree()
    End Sub

    ' Handles the BeforeExpand event for the subclassed TreeView. See further 
    ' comments about the Before_____ and After_______ TreeView event pairs in 
    ' /DirectoryScanner/DirectoryScanner.vb.
    Protected Overrides Sub OnBeforeExpand(ByVal tvcea As TreeViewCancelEventArgs)
        MyBase.OnBeforeExpand(tvcea)

        ' For performance reasons and to avoid TreeView "flickering" during an 
        ' large node update, it is best to wrap the update code in BeginUpdate...
        ' EndUpdate statements.
        Me.BeginUpdate()

        Dim tn As TreeNode
        ' Add child nodes for each child node in the node clicked by the user. 
        ' For performance reasons each node in the DirectoryTreeView 
        ' contains only the next level of child nodes in order to display the + sign 
        ' to indicate whether the user can expand the node. So when the user expands
        ' a node, in order for the + sign to be appropriately displayed for the next
        ' level of child nodes, *their* child nodes have to be added.
        For Each tn In tvcea.Node.Nodes
            AddDirectories(tn)
        Next tn

        Me.EndUpdate()
    End Sub

    ' This subroutine is used to add a child node for every directory under its
    ' parent node, which is passed as an argument. See further comments in the
    ' OnBeforeExpand event handler.
    Sub AddDirectories(ByVal tn As TreeNode)
        tn.Nodes.Clear()

        Dim strPath As String = tn.FullPath
        Dim diDirectory As New DirectoryInfo(strPath)
        Dim adiDirectories() As DirectoryInfo

        Try
            ' Get an array of all sub-directories as DirectoryInfo objects.
            adiDirectories = diDirectory.GetDirectories()
        Catch exp As Exception
            Exit Sub
        End Try

        Dim di As DirectoryInfo
        For Each di In adiDirectories
            ' Create a child node for every sub-directory, passing in the directory
            ' name and the images its node will use.
            Dim tnDir As New TreeNode(di.Name, 1, 2)
            ' Add the new child node to the parent node.
            tn.Nodes.Add(tnDir)

            ' We could now fill up the whole tree by recursively calling 
            ' AddDirectories():
            '
            'AddDirectories(tnDir)
            '
            ' This is way too slow, however. Give it a try!
        Next
    End Sub

    ' This subroutine clears out the existing TreeNode objects and rebuilds the 
    ' DirectoryTreeView, showing the logical drives.
    Public Sub RefreshTree()

        ' For performance reasons and to avoid TreeView "flickering" during an 
        ' large node update, it is best to wrap the update code in BeginUpdate...
        ' EndUpdate statements.
        BeginUpdate()

        Nodes.Clear()

        ' Make disk drives the root nodes. 
        ' Dim astrDrives As String() = Directory.GetDirectories(startupDir) 'GetLogicalDrives()

        Dim strDrive As String
        strDrive = startupDir
        '   For Each strDrive In astrDrives
        'If strDrive = "C:\IDS\Pattern_Dir" Then
        ' Dim text As String = strDrive.Remove(0, 7)
        strDrive = strDrive.TrimEnd("\"c)
        'Dim index As Integer = strDrive.LastIndexOf("\"c)
        'Dim dispText As String = strDrive.Remove(0, index + 1)

        Dim tnDrive As New TreeNode(strDrive, 1, 2)
        tnDrive.Text = strDrive

        Nodes.Add(tnDrive)
        AddDirectories(tnDrive)

        ' Set the C drive as the default selection.

        Me.SelectedNode = tnDrive
        'End If
        '  Next

        EndUpdate()
    End Sub

    Private Sub InitializeComponent()

    End Sub
End Class
