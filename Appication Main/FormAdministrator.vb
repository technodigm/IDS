
Imports System.IO
Imports System
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports DLL_DataManager
Imports DLL_Export_Service

Public Class FormAdministration
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents TabControlMain As System.Windows.Forms.TabControl
    Friend WithEvents PageUsers As System.Windows.Forms.TabPage
    Friend WithEvents TabControlView As System.Windows.Forms.TabControl
    Friend WithEvents tabGroup As System.Windows.Forms.TabPage
    Friend WithEvents tabPrivilege As System.Windows.Forms.TabPage
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TreePrivilege As System.Windows.Forms.TreeView
    Friend WithEvents TextName As System.Windows.Forms.TextBox
    Friend WithEvents TextDepartment As System.Windows.Forms.TextBox
    Friend WithEvents TextPassWord As System.Windows.Forms.TextBox
    Friend WithEvents TextContact As System.Windows.Forms.TextBox
    Friend WithEvents btnUserApply As System.Windows.Forms.Button
    Friend WithEvents BtnUserCancel As System.Windows.Forms.Button
    Friend WithEvents BtnNewUser As System.Windows.Forms.Button
    Friend WithEvents BtnDelUser As System.Windows.Forms.Button
    Friend WithEvents ListAllGroup As System.Windows.Forms.ListBox
    Friend WithEvents ListUserGroup As System.Windows.Forms.ListBox
    Friend WithEvents BtnAddGrouptoUser As System.Windows.Forms.Button
    Friend WithEvents BtnDelGrouptoUser As System.Windows.Forms.Button
    Friend WithEvents btnGroupCancel As System.Windows.Forms.Button
    Friend WithEvents btnGroupApply As System.Windows.Forms.Button
    Friend WithEvents BtnDelUsertoGroup As System.Windows.Forms.Button
    Friend WithEvents BtnAddUsertoGroup As System.Windows.Forms.Button
    Friend WithEvents ListGroupUser As System.Windows.Forms.ListBox
    Friend WithEvents ListAllUser As System.Windows.Forms.ListBox
    Friend WithEvents BtnNewGroup As System.Windows.Forms.Button
    Friend WithEvents BtnDelGroup As System.Windows.Forms.Button
    Friend WithEvents TextGroupInfo As System.Windows.Forms.TextBox
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents CMMain As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents TBAdmin As System.Windows.Forms.ToolBar
    Friend WithEvents TBNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBDelete As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBPrint As System.Windows.Forms.ToolBarButton
    Friend WithEvents TreeGroupID As System.Windows.Forms.TreeView
    Friend WithEvents TreeUserID As System.Windows.Forms.TreeView
    Friend WithEvents TextUserID As System.Windows.Forms.TextBox
    Friend WithEvents TextGroupID As System.Windows.Forms.TextBox
    Friend WithEvents PageGroups As System.Windows.Forms.TabPage
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents CMData As System.Windows.Forms.ContextMenu
    Friend WithEvents CMGroupTable As System.Windows.Forms.MenuItem
    Friend WithEvents CMPrivilegeTable As System.Windows.Forms.MenuItem
    Friend WithEvents CMUserTable As System.Windows.Forms.MenuItem
    Friend WithEvents CMHardwareTable As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents CMUpdate As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents CMExecutionTable As System.Windows.Forms.MenuItem
    Friend WithEvents CMPatternTable As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormAdministration))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.TabControlMain = New System.Windows.Forms.TabControl
        Me.PageGroups = New System.Windows.Forms.TabPage
        Me.TabControlView = New System.Windows.Forms.TabControl
        Me.tabGroup = New System.Windows.Forms.TabPage
        Me.TextGroupID = New System.Windows.Forms.TextBox
        Me.TextGroupInfo = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.BtnNewGroup = New System.Windows.Forms.Button
        Me.BtnDelGroup = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.BtnDelUsertoGroup = New System.Windows.Forms.Button
        Me.BtnAddUsertoGroup = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.ListGroupUser = New System.Windows.Forms.ListBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ListAllUser = New System.Windows.Forms.ListBox
        Me.tabPrivilege = New System.Windows.Forms.TabPage
        Me.TreePrivilege = New System.Windows.Forms.TreeView
        Me.TreeGroupID = New System.Windows.Forms.TreeView
        Me.btnGroupCancel = New System.Windows.Forms.Button
        Me.btnGroupApply = New System.Windows.Forms.Button
        Me.PageUsers = New System.Windows.Forms.TabPage
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.TextUserID = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BtnDelGrouptoUser = New System.Windows.Forms.Button
        Me.BtnAddGrouptoUser = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.ListUserGroup = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListAllGroup = New System.Windows.Forms.ListBox
        Me.BtnNewUser = New System.Windows.Forms.Button
        Me.BtnDelUser = New System.Windows.Forms.Button
        Me.BtnUserCancel = New System.Windows.Forms.Button
        Me.TextName = New System.Windows.Forms.TextBox
        Me.TextDepartment = New System.Windows.Forms.TextBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.TextPassWord = New System.Windows.Forms.TextBox
        Me.TextContact = New System.Windows.Forms.TextBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.btnUserApply = New System.Windows.Forms.Button
        Me.TreeUserID = New System.Windows.Forms.TreeView
        Me.CMData = New System.Windows.Forms.ContextMenu
        Me.CMGroupTable = New System.Windows.Forms.MenuItem
        Me.CMPrivilegeTable = New System.Windows.Forms.MenuItem
        Me.CMUserTable = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.CMHardwareTable = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.CMExecutionTable = New System.Windows.Forms.MenuItem
        Me.CMPatternTable = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.CMUpdate = New System.Windows.Forms.MenuItem
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.CMMain = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.TBAdmin = New System.Windows.Forms.ToolBar
        Me.TBNew = New System.Windows.Forms.ToolBarButton
        Me.TBDelete = New System.Windows.Forms.ToolBarButton
        Me.TBPrint = New System.Windows.Forms.ToolBarButton
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.TabControlMain.SuspendLayout()
        Me.PageGroups.SuspendLayout()
        Me.TabControlView.SuspendLayout()
        Me.tabGroup.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.tabPrivilege.SuspendLayout()
        Me.PageUsers.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TabControlMain)
        Me.Panel1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel1.Location = New System.Drawing.Point(4, 43)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(681, 523)
        Me.Panel1.TabIndex = 59
        '
        'TabControlMain
        '
        Me.TabControlMain.Controls.Add(Me.PageGroups)
        Me.TabControlMain.Controls.Add(Me.PageUsers)
        Me.TabControlMain.Location = New System.Drawing.Point(4, 9)
        Me.TabControlMain.Name = "TabControlMain"
        Me.TabControlMain.Padding = New System.Drawing.Point(20, 3)
        Me.TabControlMain.SelectedIndex = 0
        Me.TabControlMain.Size = New System.Drawing.Size(671, 498)
        Me.TabControlMain.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabControlMain.TabIndex = 8
        '
        'PageGroups
        '
        Me.PageGroups.AllowDrop = True
        Me.PageGroups.Controls.Add(Me.TabControlView)
        Me.PageGroups.Controls.Add(Me.TreeGroupID)
        Me.PageGroups.Controls.Add(Me.btnGroupCancel)
        Me.PageGroups.Controls.Add(Me.btnGroupApply)
        Me.PageGroups.Location = New System.Drawing.Point(4, 24)
        Me.PageGroups.Name = "PageGroups"
        Me.PageGroups.Size = New System.Drawing.Size(663, 470)
        Me.PageGroups.TabIndex = 0
        Me.PageGroups.Text = "Groups"
        '
        'TabControlView
        '
        Me.TabControlView.Controls.Add(Me.tabGroup)
        Me.TabControlView.Controls.Add(Me.tabPrivilege)
        Me.TabControlView.Location = New System.Drawing.Point(194, -2)
        Me.TabControlView.Name = "TabControlView"
        Me.TabControlView.SelectedIndex = 0
        Me.TabControlView.Size = New System.Drawing.Size(470, 429)
        Me.TabControlView.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabControlView.TabIndex = 59
        '
        'tabGroup
        '
        Me.tabGroup.Controls.Add(Me.TextGroupID)
        Me.tabGroup.Controls.Add(Me.TextGroupInfo)
        Me.tabGroup.Controls.Add(Me.Label6)
        Me.tabGroup.Controls.Add(Me.BtnNewGroup)
        Me.tabGroup.Controls.Add(Me.BtnDelGroup)
        Me.tabGroup.Controls.Add(Me.Label5)
        Me.tabGroup.Controls.Add(Me.GroupBox2)
        Me.tabGroup.Location = New System.Drawing.Point(4, 24)
        Me.tabGroup.Name = "tabGroup"
        Me.tabGroup.Size = New System.Drawing.Size(462, 401)
        Me.tabGroup.TabIndex = 0
        Me.tabGroup.Text = "Group Properties"
        '
        'TextGroupID
        '
        Me.TextGroupID.Location = New System.Drawing.Point(99, 20)
        Me.TextGroupID.Name = "TextGroupID"
        Me.TextGroupID.Size = New System.Drawing.Size(176, 21)
        Me.TextGroupID.TabIndex = 72
        Me.TextGroupID.Text = ""
        '
        'TextGroupInfo
        '
        Me.TextGroupInfo.Location = New System.Drawing.Point(14, 289)
        Me.TextGroupInfo.Multiline = True
        Me.TextGroupInfo.Name = "TextGroupInfo"
        Me.TextGroupInfo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextGroupInfo.Size = New System.Drawing.Size(438, 107)
        Me.TextGroupInfo.TabIndex = 71
        Me.TextGroupInfo.Text = "textGroupInfo"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(12, 270)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(76, 16)
        Me.Label6.TabIndex = 70
        Me.Label6.Text = "Information"
        '
        'BtnNewGroup
        '
        Me.BtnNewGroup.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNewGroup.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnNewGroup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnNewGroup.Location = New System.Drawing.Point(291, 15)
        Me.BtnNewGroup.Name = "BtnNewGroup"
        Me.BtnNewGroup.Size = New System.Drawing.Size(72, 28)
        Me.BtnNewGroup.TabIndex = 68
        Me.BtnNewGroup.Text = "Clear"
        '
        'BtnDelGroup
        '
        Me.BtnDelGroup.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelGroup.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnDelGroup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnDelGroup.Location = New System.Drawing.Point(370, 15)
        Me.BtnDelGroup.Name = "BtnDelGroup"
        Me.BtnDelGroup.Size = New System.Drawing.Size(72, 28)
        Me.BtnDelGroup.TabIndex = 67
        Me.BtnDelGroup.Text = "Del"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(12, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 16)
        Me.Label5.TabIndex = 65
        Me.Label5.Text = "Group Name"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Button2)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.BtnDelUsertoGroup)
        Me.GroupBox2.Controls.Add(Me.BtnAddUsertoGroup)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.ListGroupUser)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.ListAllUser)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 54)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(446, 207)
        Me.GroupBox2.TabIndex = 64
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "User List"
        '
        'Button2
        '
        Me.Button2.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button2.Location = New System.Drawing.Point(192, 160)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(64, 28)
        Me.Button2.TabIndex = 64
        Me.Button2.Text = "Save"
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.SystemColors.Highlight
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.Location = New System.Drawing.Point(192, 128)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(64, 28)
        Me.Button1.TabIndex = 63
        Me.Button1.Text = "Get"
        Me.Button1.Visible = False
        '
        'BtnDelUsertoGroup
        '
        Me.BtnDelUsertoGroup.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelUsertoGroup.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnDelUsertoGroup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnDelUsertoGroup.Location = New System.Drawing.Point(192, 72)
        Me.BtnDelUsertoGroup.Name = "BtnDelUsertoGroup"
        Me.BtnDelUsertoGroup.Size = New System.Drawing.Size(64, 28)
        Me.BtnDelUsertoGroup.TabIndex = 62
        Me.BtnDelUsertoGroup.Text = "Del"
        '
        'BtnAddUsertoGroup
        '
        Me.BtnAddUsertoGroup.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddUsertoGroup.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddUsertoGroup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnAddUsertoGroup.Location = New System.Drawing.Point(192, 40)
        Me.BtnAddUsertoGroup.Name = "BtnAddUsertoGroup"
        Me.BtnAddUsertoGroup.Size = New System.Drawing.Size(64, 28)
        Me.BtnAddUsertoGroup.TabIndex = 61
        Me.BtnAddUsertoGroup.Text = ">>"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(264, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(76, 16)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "Group User"
        '
        'ListGroupUser
        '
        Me.ListGroupUser.ItemHeight = 15
        Me.ListGroupUser.Location = New System.Drawing.Point(264, 40)
        Me.ListGroupUser.Name = "ListGroupUser"
        Me.ListGroupUser.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListGroupUser.Size = New System.Drawing.Size(168, 154)
        Me.ListGroupUser.Sorted = True
        Me.ListGroupUser.TabIndex = 52
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(12, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(76, 16)
        Me.Label4.TabIndex = 51
        Me.Label4.Text = "All Users"
        '
        'ListAllUser
        '
        Me.ListAllUser.BackColor = System.Drawing.SystemColors.Info
        Me.ListAllUser.ItemHeight = 15
        Me.ListAllUser.Location = New System.Drawing.Point(11, 41)
        Me.ListAllUser.Name = "ListAllUser"
        Me.ListAllUser.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListAllUser.Size = New System.Drawing.Size(168, 154)
        Me.ListAllUser.Sorted = True
        Me.ListAllUser.TabIndex = 0
        '
        'tabPrivilege
        '
        Me.tabPrivilege.BackColor = System.Drawing.Color.Transparent
        Me.tabPrivilege.Controls.Add(Me.TreePrivilege)
        Me.tabPrivilege.Location = New System.Drawing.Point(4, 24)
        Me.tabPrivilege.Name = "tabPrivilege"
        Me.tabPrivilege.Size = New System.Drawing.Size(462, 401)
        Me.tabPrivilege.TabIndex = 1
        Me.tabPrivilege.Text = "Group Privileges"
        Me.tabPrivilege.Visible = False
        '
        'TreePrivilege
        '
        Me.TreePrivilege.BackColor = System.Drawing.Color.WhiteSmoke
        Me.TreePrivilege.CheckBoxes = True
        Me.TreePrivilege.ImageIndex = -1
        Me.TreePrivilege.Location = New System.Drawing.Point(4, 1)
        Me.TreePrivilege.Name = "TreePrivilege"
        Me.TreePrivilege.Nodes.AddRange(New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("All Groups", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("1", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node0")}), New System.Windows.Forms.TreeNode("2", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node3"), New System.Windows.Forms.TreeNode("Node5")}), New System.Windows.Forms.TreeNode("3", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node6"), New System.Windows.Forms.TreeNode("Node7"), New System.Windows.Forms.TreeNode("Node8")})})})
        Me.TreePrivilege.SelectedImageIndex = -1
        Me.TreePrivilege.Size = New System.Drawing.Size(456, 435)
        Me.TreePrivilege.Sorted = True
        Me.TreePrivilege.TabIndex = 59
        '
        'TreeGroupID
        '
        Me.TreeGroupID.HotTracking = True
        Me.TreeGroupID.ImageIndex = -1
        Me.TreeGroupID.Location = New System.Drawing.Point(0, 0)
        Me.TreeGroupID.Name = "TreeGroupID"
        Me.TreeGroupID.Nodes.AddRange(New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node0", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node1")}), New System.Windows.Forms.TreeNode("Node2", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node3", New System.Windows.Forms.TreeNode() {New System.Windows.Forms.TreeNode("Node4")})})})
        Me.TreeGroupID.SelectedImageIndex = -1
        Me.TreeGroupID.Size = New System.Drawing.Size(196, 466)
        Me.TreeGroupID.Sorted = True
        Me.TreeGroupID.TabIndex = 6
        '
        'btnGroupCancel
        '
        Me.btnGroupCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGroupCancel.ForeColor = System.Drawing.SystemColors.Highlight
        Me.btnGroupCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnGroupCancel.Location = New System.Drawing.Point(578, 434)
        Me.btnGroupCancel.Name = "btnGroupCancel"
        Me.btnGroupCancel.Size = New System.Drawing.Size(72, 28)
        Me.btnGroupCancel.TabIndex = 63
        Me.btnGroupCancel.Text = "Cancel"
        '
        'btnGroupApply
        '
        Me.btnGroupApply.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGroupApply.ForeColor = System.Drawing.SystemColors.Highlight
        Me.btnGroupApply.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnGroupApply.Location = New System.Drawing.Point(497, 435)
        Me.btnGroupApply.Name = "btnGroupApply"
        Me.btnGroupApply.Size = New System.Drawing.Size(72, 28)
        Me.btnGroupApply.TabIndex = 62
        Me.btnGroupApply.Text = "Apply"
        '
        'PageUsers
        '
        Me.PageUsers.Controls.Add(Me.TabControl1)
        Me.PageUsers.Controls.Add(Me.TreeUserID)
        Me.PageUsers.Location = New System.Drawing.Point(4, 24)
        Me.PageUsers.Name = "PageUsers"
        Me.PageUsers.Size = New System.Drawing.Size(663, 470)
        Me.PageUsers.TabIndex = 1
        Me.PageUsers.Text = "Users"
        Me.PageUsers.Visible = False
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.ItemSize = New System.Drawing.Size(109, 20)
        Me.TabControl1.Location = New System.Drawing.Point(195, -2)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(468, 469)
        Me.TabControl1.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabControl1.TabIndex = 60
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.TextUserID)
        Me.TabPage3.Controls.Add(Me.GroupBox1)
        Me.TabPage3.Controls.Add(Me.BtnNewUser)
        Me.TabPage3.Controls.Add(Me.BtnDelUser)
        Me.TabPage3.Controls.Add(Me.BtnUserCancel)
        Me.TabPage3.Controls.Add(Me.TextName)
        Me.TabPage3.Controls.Add(Me.TextDepartment)
        Me.TabPage3.Controls.Add(Me.Label19)
        Me.TabPage3.Controls.Add(Me.Label16)
        Me.TabPage3.Controls.Add(Me.Label15)
        Me.TabPage3.Controls.Add(Me.Label20)
        Me.TabPage3.Controls.Add(Me.TextPassWord)
        Me.TabPage3.Controls.Add(Me.TextContact)
        Me.TabPage3.Controls.Add(Me.Label17)
        Me.TabPage3.Controls.Add(Me.btnUserApply)
        Me.TabPage3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabPage3.Location = New System.Drawing.Point(4, 24)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(460, 441)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "User Properties"
        Me.TabPage3.Visible = False
        '
        'TextUserID
        '
        Me.TextUserID.Location = New System.Drawing.Point(95, 21)
        Me.TextUserID.Name = "TextUserID"
        Me.TextUserID.Size = New System.Drawing.Size(176, 21)
        Me.TextUserID.TabIndex = 62
        Me.TextUserID.Text = ""
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnDelGrouptoUser)
        Me.GroupBox1.Controls.Add(Me.BtnAddGrouptoUser)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ListUserGroup)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ListAllGroup)
        Me.GroupBox1.Location = New System.Drawing.Point(9, 189)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(446, 207)
        Me.GroupBox1.TabIndex = 61
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Group List"
        '
        'BtnDelGrouptoUser
        '
        Me.BtnDelGrouptoUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelGrouptoUser.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnDelGrouptoUser.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnDelGrouptoUser.Location = New System.Drawing.Point(184, 72)
        Me.BtnDelGrouptoUser.Name = "BtnDelGrouptoUser"
        Me.BtnDelGrouptoUser.Size = New System.Drawing.Size(64, 28)
        Me.BtnDelGrouptoUser.TabIndex = 62
        Me.BtnDelGrouptoUser.Text = "Del"
        '
        'BtnAddGrouptoUser
        '
        Me.BtnAddGrouptoUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnAddGrouptoUser.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnAddGrouptoUser.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnAddGrouptoUser.Location = New System.Drawing.Point(184, 40)
        Me.BtnAddGrouptoUser.Name = "BtnAddGrouptoUser"
        Me.BtnAddGrouptoUser.Size = New System.Drawing.Size(64, 28)
        Me.BtnAddGrouptoUser.TabIndex = 61
        Me.BtnAddGrouptoUser.Text = ">>"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(264, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 16)
        Me.Label2.TabIndex = 53
        Me.Label2.Text = "User Group"
        '
        'ListUserGroup
        '
        Me.ListUserGroup.ItemHeight = 15
        Me.ListUserGroup.Location = New System.Drawing.Point(256, 40)
        Me.ListUserGroup.Name = "ListUserGroup"
        Me.ListUserGroup.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListUserGroup.Size = New System.Drawing.Size(176, 154)
        Me.ListUserGroup.Sorted = True
        Me.ListUserGroup.TabIndex = 52
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 16)
        Me.Label1.TabIndex = 51
        Me.Label1.Text = "All Groups"
        '
        'ListAllGroup
        '
        Me.ListAllGroup.BackColor = System.Drawing.SystemColors.Info
        Me.ListAllGroup.ItemHeight = 15
        Me.ListAllGroup.Location = New System.Drawing.Point(11, 41)
        Me.ListAllGroup.Name = "ListAllGroup"
        Me.ListAllGroup.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListAllGroup.Size = New System.Drawing.Size(168, 154)
        Me.ListAllGroup.Sorted = True
        Me.ListAllGroup.TabIndex = 0
        '
        'BtnNewUser
        '
        Me.BtnNewUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnNewUser.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnNewUser.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnNewUser.Location = New System.Drawing.Point(287, 15)
        Me.BtnNewUser.Name = "BtnNewUser"
        Me.BtnNewUser.Size = New System.Drawing.Size(72, 28)
        Me.BtnNewUser.TabIndex = 60
        Me.BtnNewUser.Text = "Clear"
        '
        'BtnDelUser
        '
        Me.BtnDelUser.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelUser.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnDelUser.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnDelUser.Location = New System.Drawing.Point(366, 15)
        Me.BtnDelUser.Name = "BtnDelUser"
        Me.BtnDelUser.Size = New System.Drawing.Size(72, 28)
        Me.BtnDelUser.TabIndex = 59
        Me.BtnDelUser.Text = "Del"
        '
        'BtnUserCancel
        '
        Me.BtnUserCancel.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnUserCancel.ForeColor = System.Drawing.SystemColors.Highlight
        Me.BtnUserCancel.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnUserCancel.Location = New System.Drawing.Point(379, 406)
        Me.BtnUserCancel.Name = "BtnUserCancel"
        Me.BtnUserCancel.Size = New System.Drawing.Size(72, 28)
        Me.BtnUserCancel.TabIndex = 58
        Me.BtnUserCancel.Text = "Cancel"
        '
        'TextName
        '
        Me.TextName.Location = New System.Drawing.Point(96, 56)
        Me.TextName.Name = "TextName"
        Me.TextName.Size = New System.Drawing.Size(176, 21)
        Me.TextName.TabIndex = 48
        Me.TextName.Text = ""
        '
        'TextDepartment
        '
        Me.TextDepartment.Location = New System.Drawing.Point(96, 88)
        Me.TextDepartment.Name = "TextDepartment"
        Me.TextDepartment.Size = New System.Drawing.Size(176, 21)
        Me.TextDepartment.TabIndex = 52
        Me.TextDepartment.Text = ""
        '
        'Label19
        '
        Me.Label19.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label19.Location = New System.Drawing.Point(14, 59)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(40, 16)
        Me.Label19.TabIndex = 53
        Me.Label19.Text = "Name"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(16, 92)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(80, 16)
        Me.Label16.TabIndex = 55
        Me.Label16.Text = "Department"
        '
        'Label15
        '
        Me.Label15.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Label15.Location = New System.Drawing.Point(14, 121)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 57
        Me.Label15.Text = "Contact No"
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(15, 155)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(64, 16)
        Me.Label20.TabIndex = 50
        Me.Label20.Text = "Password"
        '
        'TextPassWord
        '
        Me.TextPassWord.Location = New System.Drawing.Point(96, 154)
        Me.TextPassWord.Name = "TextPassWord"
        Me.TextPassWord.Size = New System.Drawing.Size(176, 21)
        Me.TextPassWord.TabIndex = 54
        Me.TextPassWord.Text = ""
        '
        'TextContact
        '
        Me.TextContact.Location = New System.Drawing.Point(96, 121)
        Me.TextContact.Name = "TextContact"
        Me.TextContact.Size = New System.Drawing.Size(176, 21)
        Me.TextContact.TabIndex = 56
        Me.TextContact.Text = ""
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(16, 24)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(48, 16)
        Me.Label17.TabIndex = 49
        Me.Label17.Text = "User ID"
        '
        'btnUserApply
        '
        Me.btnUserApply.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUserApply.ForeColor = System.Drawing.SystemColors.Highlight
        Me.btnUserApply.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnUserApply.Location = New System.Drawing.Point(297, 408)
        Me.btnUserApply.Name = "btnUserApply"
        Me.btnUserApply.Size = New System.Drawing.Size(72, 28)
        Me.btnUserApply.TabIndex = 47
        Me.btnUserApply.Text = "Apply"
        '
        'TreeUserID
        '
        Me.TreeUserID.ImageIndex = -1
        Me.TreeUserID.Location = New System.Drawing.Point(-2, 0)
        Me.TreeUserID.Name = "TreeUserID"
        Me.TreeUserID.SelectedImageIndex = -1
        Me.TreeUserID.Size = New System.Drawing.Size(196, 465)
        Me.TreeUserID.Sorted = True
        Me.TreeUserID.TabIndex = 7
        '
        'CMData
        '
        Me.CMData.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.CMGroupTable, Me.CMPrivilegeTable, Me.CMUserTable, Me.MenuItem7, Me.CMHardwareTable, Me.MenuItem6, Me.CMExecutionTable, Me.CMPatternTable, Me.MenuItem3, Me.CMUpdate})
        '
        'CMGroupTable
        '
        Me.CMGroupTable.Index = 0
        Me.CMGroupTable.Text = "GroupTable"
        '
        'CMPrivilegeTable
        '
        Me.CMPrivilegeTable.Index = 1
        Me.CMPrivilegeTable.Text = "PrivilegeTable"
        '
        'CMUserTable
        '
        Me.CMUserTable.Index = 2
        Me.CMUserTable.Text = "UserTable"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 3
        Me.MenuItem7.Text = "-"
        '
        'CMHardwareTable
        '
        Me.CMHardwareTable.Index = 4
        Me.CMHardwareTable.Text = "HardwareTable"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 5
        Me.MenuItem6.Text = "-"
        '
        'CMExecutionTable
        '
        Me.CMExecutionTable.Index = 6
        Me.CMExecutionTable.Text = "ExecutionTable"
        '
        'CMPatternTable
        '
        Me.CMPatternTable.Index = 7
        Me.CMPatternTable.Text = "PatternTable"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 8
        Me.MenuItem3.Text = "-"
        '
        'CMUpdate
        '
        Me.CMUpdate.Index = 9
        Me.CMUpdate.Text = "Update Data"
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Data Source=""C:\IDS\BIN\EEtver1.mdb"";Mode=Share Deny None;Jet OLED" & _
        "B:Engine Type=5;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLEDB:System database=;Je" & _
        "t OLEDB:SFP=False;persist security info=False;Extended Properties=;Jet OLEDB:Com" & _
        "pact Without Replica Repair=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Cre" & _
        "ate System Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;User ID=A" & _
        "dmin;Jet OLEDB:Global Bulk Transactions=1"
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'CMMain
        '
        Me.CMMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Add"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Delete"
        '
        'TBAdmin
        '
        Me.TBAdmin.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBNew, Me.TBDelete, Me.TBPrint})
        Me.TBAdmin.ButtonSize = New System.Drawing.Size(40, 35)
        Me.TBAdmin.DropDownArrows = True
        Me.TBAdmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TBAdmin.ImageList = Me.ImageList1
        Me.TBAdmin.Location = New System.Drawing.Point(0, 0)
        Me.TBAdmin.Name = "TBAdmin"
        Me.TBAdmin.ShowToolTips = True
        Me.TBAdmin.Size = New System.Drawing.Size(688, 41)
        Me.TBAdmin.TabIndex = 60
        Me.TBAdmin.Visible = False
        '
        'TBNew
        '
        Me.TBNew.ImageIndex = 0
        Me.TBNew.Text = "Demo"
        Me.TBNew.ToolTipText = "Demonstration of using Data Manager"
        '
        'TBDelete
        '
        Me.TBDelete.ImageIndex = 1
        Me.TBDelete.ToolTipText = "Delete"
        '
        'TBPrint
        '
        Me.TBPrint.ImageIndex = 2
        Me.TBPrint.ToolTipText = "Print"
        '
        'TabPage1
        '
        Me.TabPage1.Location = New System.Drawing.Point(4, 24)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(663, 470)
        Me.TabPage1.TabIndex = 2
        Me.TabPage1.Text = "PageData"
        '
        'FormAdministration
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 553)
        Me.Controls.Add(Me.TBAdmin)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "FormAdministration"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Administration"
        Me.Panel1.ResumeLayout(False)
        Me.TabControlMain.ResumeLayout(False)
        Me.PageGroups.ResumeLayout(False)
        Me.TabControlView.ResumeLayout(False)
        Me.tabGroup.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.tabPrivilege.ResumeLayout(False)
        Me.PageUsers.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim DataDS As New DataSet
#Region "Support methods"
     '''''''''''''''''''''''''
    '' refresh form and display components
     '''''''''''''''''''''''''
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call RefreshAll()           '' load info form database to all display component
        Call ClearAll()             '' clear all display component
        Call LoadBlankPrivilege()
        'Dim x As Integer = TreeGroupID.Nodes(0).SelectedImageIndex()
    End Sub

     '''''''''''''''''''''''''
    '' load items into group list box from database
     '''''''''''''''''''''''''
    Private Sub RefreshAllGroup()
        Dim I As Integer

        DBView = New DataView(DS.Tables("GroupTable"))          'open group table
        DBView.RowFilter = ""                                   'no filter
        ListAllGroup.Items.Clear()                              'clear list box 
        If DBView.Count > 0 Then
            For I = 0 To DBView.Count - 1
                Dim Row As DataRow = DBView(I).Row
                ListAllGroup.Items.Add(Row("GroupID"))          'add row info into the group list box
            Next
        End If

    End Sub

     '''''''''''''''''''''''''
    '' load items into user list box from database
     '''''''''''''''''''''''''
    Private Sub RefreshAllUser()
        Dim I As Integer

        DBView = New DataView(DS.Tables("UserTable"))           'open user table
        DBView.RowFilter = ""                                   'no filter
        ListAllUser.Items.Clear()                               'clear list box
        If DBView.Count > 0 Then
            For I = 0 To DBView.Count - 1
                Dim Row As DataRow = DBView(I).Row
                ListAllUser.Items.Add(Row("UserID"))             'add row info into the user list box 
            Next
        End If

    End Sub

     '''''''''''''''''''''''''
    '' load user tree-component with data from database
    '
    '   ids.data.group -> group id -> user id
    '  (primarynode)   (subnode)  (subsubnode)
    '
     '''''''''''''''''''''''''
    Private Sub RefreshTreeGroup()
        Dim I As Integer
        Dim J As Integer

        TreeGroupID.Nodes.Clear()                           'clear subnodes
        TreeGroupID.HotTracking = True                      'on hottracking
        TreeGroupID.Nodes.Add("IDS.data Groups")            'add the primary node

        DBView = New DataView(DS.Tables("GroupTable"))      'open group table
        DBView.Sort = "GroupID Asc"                         'in ascending order
        DBView.RowFilter = ""
        I = 0
        For I = 0 To DBView.Count - 1
            Dim RowGroup As DataRow = DBView(I).Row

            Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable")) 'open user table that is under the particular group
            SubView.Sort = "UserID Asc"                                         'in ascending order
            SubView.RowFilter = "GroupID = '" + RowGroup("GroupID") + "'"       'filter for user id that fall uder selected group id

            TreeGroupID.Nodes(0).Nodes.Add(RowGroup("GroupID"))                 'add the group id node
            If SubView.Count > 0 Then
                J = 0
                For J = 0 To SubView.Count - 1
                    Dim Row As DataRow = SubView(J).Row                         'get the user id row info
                    TreeGroupID.Nodes(0).Nodes(I).Nodes.Add(Row("UserID"))      'add the user id node under the group node
                Next J
            End If
        Next I

        'TreeGroupID.ExpandAll()
        TreeGroupID.Nodes(0).Toggle()                                           'toggle on the tree


    End Sub

     '''''''''''''''''''''''''
    '' load group tree-component with data from database
    '
    '   ids.data.user-> user id -> group id
    '  (primarynode)   (subnode)  (subsubnode)
    '
     '''''''''''''''''''''''''
    Private Sub RefreshTreeUser()
        Dim I As Integer
        Dim J As Integer


        TreeUserID.Nodes.Clear()                                        'clear all subnodes
        TreeUserID.HotTracking = True                                   'on hot tracking
        TreeUserID.Nodes.Add("IDS.data Users")                          'add primary node

        DBView = New DataView(DS.Tables("UserTable"))                   'open user table
        DBView.Sort = "UserID Asc"                                      'in ascending order
        DBView.RowFilter = ""

        I = 0
        For I = 0 To DBView.Count - 1
            Dim RowUser As DataRow = DBView(I).Row

            Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable"))     'open group user table
            SubView.Sort = "GroupID Asc"                                            'in ascending order
            SubView.RowFilter = "UserID = '" + RowUser("UserID") + "'"              'filter for only groupid uder the selected user id

            TreeUserID.Nodes(0).Nodes.Add(RowUser("UserID"))                        'add subnode of user id 

            If SubView.Count > 0 Then
                J = 0
                For J = 0 To SubView.Count - 1
                    Dim Row As DataRow = SubView(J).Row
                    TreeUserID.Nodes(0).Nodes(I).Nodes.Add(Row("GroupID"))          'add subsubnode of group id
                Next J
            End If
        Next I
        'TreeUserID.ExpandAll()
        TreeUserID.Nodes(0).Toggle()                                                 'toggle the primary node   


    End Sub
     '''''''''''''''''''''''''
    '' load info form database to all display component
     '''''''''''''''''''''''''
    Private Sub RefreshAll()
        Call RefreshTreeUser()
        Call RefreshTreeGroup()
        Call RefreshAllUser()
        Call RefreshAllGroup()

    End Sub

     '''''''''''''''''''''''''
    '' clear all display component
     '''''''''''''''''''''''''
    Private Sub ClearAll()
        ListGroupUser.Items.Clear()
        ListUserGroup.Items.Clear()
        TextGroupInfo.Text = ""
        TextGroupID.Text = ""

        TextUserID.Text = ""
        TextName.Text = ""
        TextDepartment.Text = ""
        TextContact.Text = ""
        TextPassWord.Text = ""

    End Sub
#End Region

#Region "Tree Privilege"


      '''
    ''initialize the tree privilege with privilege type nodes and privilege id subnodes
    '
    '           privilege tree -> privilege type -> privilege id
    '             (primarynode)   (subnode)         (subsubnode)'                           
      '''


    Private Sub LoadBlankPrivilege()

         '''''''
        '' Load Privilege item into the TreePrivilege
         '''''''
        Dim I As Integer
        Dim J As Integer
        Dim DBView As New DataView

        'get the privilege type first
        DBView = New DataView(DS.Tables("PrivilegeTable"))                      'open privilege table
        DBView.Sort = "PrivilegeType Asc"                                       'in ascending order
        DBView.RowFilter = ""
        I = 0
        Dim ListType As New ListBox                                             'create instance of temp listtype
        ListType.Items.Clear()                                                  'clear listtype item
        Dim ExistFlag As Boolean                                                'mark the status is exist or not
        For I = 0 To DBView.Count - 1
            Dim Row As DataRow = DBView(I).Row                                  'get row info

            If ListType.Items.Count = 0 Then                                    'if list count = 0
                ListType.Items.Add(Row("PrivilegeType"))                        'add the row item
            ElseIf ListType.Items.Count > 0 Then                                'if list count > 0
                J = 0
                For J = 0 To ListType.Items.Count - 1                           '
                    If CStr(Row("PrivilegeType")) = CStr(ListType.Items(J)) Then 'if the privilege type is already in the listtype item
                        ExistFlag = True                                        'set the flag to true
                        Exit For
                    Else
                        ExistFlag = False                                       'else set to false
                    End If
                Next J
                If ExistFlag = False Then                                       'if privilege type is not in the listtype item, 
                    ListType.Items.Add(Row("PrivilegeType"))                    'add the privilege type into the listtype item
                End If
            End If
        Next I

         ''''''''''''''''''''''''''
        '  adding tree items
         '''''''''''''''''''''''''''
        TreePrivilege.Nodes.Clear()                                             'clear the treeprivilege nodes
        I = 0
        For I = 0 To ListType.Items.Count - 1                                   'iterate thru the temp list type item
            TreePrivilege.Nodes.Add(ListType.Items(I))                          'create tree privilege node with listtype item

            DBView = New DataView(DS.Tables("PrivilegeTable"))                      'open privilege table    
            DBView.RowFilter = "PrivilegeType = '" + CStr(ListType.Items(I)) + "'"  'filter for listype = privilege type

            If DBView.Count > 0 Then
                J = 0
                For J = 0 To DBView.Count - 1                                               'iterate thru the data table
                    Dim Row As DataRow = DBView(J).Row                                      'get the row info

                    Dim SSView As DataView = New DataView(DS.Tables("SystemConfigureTable"))    'open the system configuretable
                    SSView.RowFilter = "Hardware = '" + row("PrivilegeID") + "'"                'filter for hardwareid = privilege id of privilege table

                    If SSView.Count = 0 Then
                        TreePrivilege.Nodes(I).Nodes.Add(row("PrivilegeID"))                    'if not found, add privilege id to the privilege type as the subsubnode
                    Else

                        Dim SSRow As DataRow = SSView(0).Row                                    'if found,
                        If SSRow("Selected") = True Then                                        'if the hardware flag is in selected mode

                            TreePrivilege.Nodes(I).Nodes.Add(row("PrivilegeID"))                'treeprivilege add privilege id into the subnode of privilege type node
                            If SSRow("Hardware") = "HeightSensor" Then                          'for heighsenor, check it if it is found at the subnode of privilege type node
                                Dim x = TreePrivilege.Nodes(I).GetNodeCount(False)
                                TreePrivilege.Nodes(I).Nodes(x - 1).Checked = True
                            End If
                        End If

                    End If
                Next J
            End If
        Next I
        TreePrivilege.ExpandAll()                                                               'expand the tree privilege nodes


    End Sub

     '''''''''''''''''''''''''
    '' when user check on the item of the tree privilege node
     '''''''''''''''''''''''''
    Private Sub TreePrivilege_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreePrivilege.AfterCheck
        Dim Mynode As TreeNode = e.Node
        Dim NodeCount As Integer = Mynode.GetNodeCount(False)   'tree privilege count number of item

        If NodeCount > 0 Then                                   'if is greater than 0
            Dim I As Integer
            For I = 0 To NodeCount - 1                          'iterate thru the nodes
                If Mynode.Checked = True Then
                    Mynode.Nodes(I).Checked = True              'check it on
                Else
                    Mynode.Nodes(I).Checked = False             'check it off
                End If
            Next
        End If
    End Sub

     '''''''''''''''''''''''''
    '' load group privilege info into tree privilege nodes
    '
    '           privilege tree -> privilege type -> privilege id
    '             (primarynode)   (subnode)         (subsubnode)
     '''''''''''''''''''''''''
    Private Sub LoadGroupPrivilege()
        Dim I As Integer
        Dim J As Integer
        Dim K As Integer
        Dim DBView As New DataView

         '''''''
        '' Load Privilege item into the TreePrivilege
         '''''''
        'get the privilege type first
        DBView = New DataView(DS.Tables("GroupPrivilegeTable"))                 'open group privilege table
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"               'filter for the select group id

        Call LoadBlankPrivilege()                                               'call the blank tree privilege 
        If DBView.Count = 0 Then
            Return
        Else
            I = 0
            For I = 0 To TreePrivilege.GetNodeCount(False) - 1                  'iterate thru tree privilege items
                J = 0
                For J = 0 To TreePrivilege.Nodes(I).GetNodeCount(False) - 1     'iterate thru tree privilege subnode count
                    K = 0
                    For K = 0 To DBView.Count - 1
                        Dim Row As DataRow = DBView(K).Row                      'get the row info from group privilege table
                        Dim Str = CStr(Row("PrivilegeID"))

                        If TreePrivilege.Nodes(I).Nodes(J).Text = Str Then      'if privilege id = subsubnode id
                            TreePrivilege.Nodes(I).Nodes(J).Checked = True      'check it
                            Exit For                                            'exit for loop
                        Else
                            TreePrivilege.Nodes(I).Nodes(J).Checked = False     'uncheck it
                        End If
                    Next K
                Next J
            Next I
        End If

    End Sub
#End Region

#Region "TREE GroupID"

     ''
    '  after checking tree group item 
     '''

    Private Sub TreeGroupID_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeGroupID.AfterSelect
        Dim I As Integer

        If TreeGroupID.SelectedNode.Text = "IDS.data Groups" Then 'if selected node is the primary node (ids.data.groups)
            Return
        End If

        DBView = New DataView(DS.Tables("GroupTable"))          'open group table
        DBView.RowFilter = ""
        For I = 0 To DBView.Count - 1                           'iterate thru the group table items
            Dim Row As DataRow = DBView(I).Row                  'get the row info

            'if tree group selected node (groupId) = row (groupid) from group table
            If TreeGroupID.SelectedNode.Text = Row("GroupID") Or TreeGroupID.SelectedNode.Parent.Text = Row("GroupID") Then
                TextGroupID.Text = Row("GroupID")                       'display the group info at the form

                If TextGroupID.Text.ToUpper = "ADMINISTRATION" Then     'if the group is under administration
                    tabGroup.Enabled = False                            'disable the group tab display
                Else
                    tabGroup.Enabled = True                             'else enable it
                End If

                If IsDBNull(Row("Remark")) = False Then                 'any remarks for the group id?
                    TextGroupInfo.Text = Row("Remark")                  'yes, display it
                Else
                    TextGroupInfo.Text = ""                             'no, display blank
                End If


                Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable"))     'open groupuser table
                SubView.RowFilter = "GroupID = '" + Row("GroupID") + "'"                'filter for the selected above groupid

                ListGroupUser.Items.Clear()                                             'clear the groupuser listbox
                If SubView.Count > 0 Then
                    Dim J As Integer
                    For J = 0 To SubView.Count - 1                                      'iterate thru the groupuser table
                        Dim UserRow As DataRow = SubView(J).Row                         'get the row info of user id
                        ListGroupUser.Items.Add(UserRow("UserID"))                      'add the user id into the group user listbox
                    Next

                    Call LoadGroupPrivilege()                                           '' load group privilege info into tree privilege nodes

                End If

            End If
        Next
    End Sub
     '''''''''''''''''''''''''
    '' add user to group user list box
     '''''''''''''''''''''''''
    Private Sub BtnAddUsertoGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddUsertoGroup.Click
        Dim I As Integer
        If ListAllUser.SelectedItems.Count > 0 Then 'if there are selected users from list all users 

            I = 0
            For I = 0 To ListAllUser.SelectedItems.Count - 1    'iterater thru the selected users

                If ListGroupUser.Items.Contains(ListAllUser.SelectedItems.Item(I)) = True Then
                Else
                    ListGroupUser.Items.Add(ListAllUser.SelectedItems.Item(I))  'add the selected user from all user list box into the group user list box
                End If
            Next
        End If
        ListAllUser.ClearSelected()  'clear the selection
    End Sub
     '''''''''''''''''''''''''
    '' delete user from the group user list box
     '''''''''''''''''''''''''
    Private Sub BtnDelUsertoGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelUsertoGroup.Click
        Dim I As Integer
        If ListGroupUser.SelectedItems.Count > 0 Then   'if there are selected users from the listgroupuser box
            Dim A(ListGroupUser.SelectedItems.Count - 1) As String    'declare array
            ListGroupUser.SelectedItems.CopyTo(A, 0)                   'copy selected items of listgroupuser to array 

            I = 0
            For I = 0 To A.Length - 1
                ListGroupUser.Items.Remove(A(I))                        'clear the array
            Next I
        End If

    End Sub
     '''''''''''''''''''''''''
    '' create new group
     '''''''''''''''''''''''''
    Private Sub BtnNewGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNewGroup.Click

        Call RefreshAll()       '' load info form database to all display component
        Call ClearAll()         '' clear all display component

    End Sub
     '''''''''''''''''''''''''
    '' delete group
     '''''''''''''''''''''''''
    Private Sub BtnDelGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelGroup.Click

        Dim I As Integer

         '''''''''''''''''''''''''
        '' validate the group id for error
         '''''''''''''''''''''''''

        IDS.Data.MsgErr = ""
        If TextGroupID.Text = "" Then
            IDS.Data.MsgErr = "GroupID cant be blank"
            MessageBox.Show(IDS.Data.MsgErr)
            Return
        ElseIf TextGroupID.Text.ToUpper = "ADMINISTRATION" Then
            IDS.Data.MsgErr = "You are not allow to delete Group Administraion"
            MessageBox.Show(IDS.Data.MsgErr)
            Return
        End If

         '''''''''''''''''''''''''
        '' delete group and update to database
         '''''''''''''''''''''''''
        DBView = New DataView(DS.Tables("GroupTable"))
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"

        If DBView.Count = 0 Then
        Else
            Dim Respond As DialogResult = MessageBox.Show("Are you sure to Delete GroupID " + TextGroupID.Text + "?", "", MessageBoxButtons.YesNo)
            If Respond <> DialogResult.Yes Then
                Return
            Else
                Dim row As DataRow = DBView(0).Row
                row.Delete()
            End If
        End If
        UpdateData("SELECT * FROM GroupTable", "GroupTable")

        Call RefreshAll()       '' load info form database to all display component
        Call ClearAll()         '' clear all display component
    End Sub
     '''''''''''''''''''''''''
    '' cancel the changes of group panel
     '''''''''''''''''''''''''
    Private Sub btnGroupCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupCancel.Click
        Dim I As Integer
        Dim J As Integer


          '''
        '' open group table and fill the group panel with the selected group id
          '''

        DBView = New DataView(DS.Tables("GroupTable"))
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"

        If DBView.Count > 0 Then
            Dim Row As DataRow = DBView(0).Row
            If IsDBNull(Row("Remark")) = False Then
                TextGroupInfo.Text = Row("Remark")
            Else
                TextGroupInfo.Text = ""
            End If

              '''
            '' open group user table and fill the listgroupuser with user id that under the selected group
              '''

            Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable"))
            SubView.RowFilter = "GroupID = '" + Row("GroupID") + "'"

            ListGroupUser.Items.Clear()
            If SubView.Count > 0 Then
                J = 0
                For J = 0 To SubView.Count - 1
                    Dim UserRow As DataRow = SubView(J).Row
                    ListGroupUser.Items.Add(UserRow("UserID"))
                Next
            End If
        End If
        Call LoadGroupPrivilege()    '' load group privilege info into tree privilege nodes

    End Sub
     '''''''''''''''''''''''''
    '' to save changes of group info to the database
     '''''''''''''''''''''''''
    Private Sub btnGroupApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupApply.Click
        Dim I As Integer
        Dim J As Integer

         '''''''''''''''''''''''''
        '' validate the group id for error
         '''''''''''''''''''''''''

        IDS.Data.MsgErr = ""
        If TextGroupID.Text = "" Then
            IDS.Data.MsgErr = "GroupID cant be blank"
            MessageBox.Show(IDS.Data.MsgErr)
            Return
        End If
         '''''''''''''''''''''''''
         '''''''''''''''''''''''''

        DBView = New DataView(DS.Tables("GroupTable"))
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"
        IDS.Data.MsgErr = ""

        If DBView.Count = 0 Then
        Else
            ' there is an exisiting group id in the database, user choose whether to overwrite it.
            Dim Respond As DialogResult = MessageBox.Show("There is an existing record of " + TextGroupID.Text + ". Do you want to overwrite it?", "", MessageBoxButtons.YesNo)
            If Respond <> DialogResult.Yes Then
                Return                                                  'dunno overwrite the exisiting record, exit from function
            Else
                Dim row As DataRow = DBView(0).Row                      'get the group id info
                row.Delete()                                            'del the group id info
                UpdateData("SELECT * FROM GroupTable", "GroupTable")    'update to the database
            End If
        End If

        Dim NewRow As DataRow = DBView.Table.NewRow()                   'create a new row
        NewRow("GroupID") = TextGroupID.Text                            'save the group id
        NewRow("remark") = TextGroupInfo.Text                           'save the remark
        DBView.Table.Rows.Add(NewRow)                                   'add the row
        UpdateData("SELECT * FROM GroupTable", "GroupTable")            'update to the database

        DBView = New DataView(DS.Tables("GroupUserTable"))              'open group user table
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"       'filter for selected group id

        If ListGroupUser.Items.Count > 0 Then                           '
            I = 0
            For I = 0 To ListGroupUser.Items.Count - 1                  'iterate thru the listgroupuser items
                Dim Row As DataRow = DBView.Table.NewRow()              'create new row
                Row("GroupID") = TextGroupID.Text                       'save group id
                Row("UserID") = ListGroupUser.Items(I)                  'save user id
                DBView.Table.Rows.Add(Row)                              'add row
            Next
            UpdateData("SELECT * FROM GroupUserTable", "GroupUserTable") 'update to the database
        End If

         ''''
        '   update group privilege to the database
         ''''
        Dim PrivilegeArray As New ArrayList                             'create a temp privilege array
        I = 0
        For I = 0 To TreePrivilege.Nodes.Count - 1                      'iterate thru the tree privilege nodes (privilege type)
            J = 0
            For J = 0 To TreePrivilege.Nodes(I).Nodes.Count - 1         'iterate thru the tree privilege subnobes (privildege id)
                If TreePrivilege.Nodes(I).Nodes(J).Checked = True Then  'if privilege id is checked    
                    PrivilegeArray.Add(TreePrivilege.Nodes(I).Nodes(J).Text)    'add the privilege id to the temp privilege array
                End If
            Next J
        Next I


        DBView = New DataView(DS.Tables("GroupPrivilegeTable"))         'open group privilege table
        DBView.RowFilter = "GroupID = '" + TextGroupID.Text + "'"       'filter for selected group id
        If PrivilegeArray.Count > 0 Then                                'if there is item in the temp privilege array
            I = 0
            For I = 0 To PrivilegeArray.Count - 1
                Dim Row As DataRow = DBView.Table.NewRow()              'create new row
                Row("GroupID") = TextGroupID.Text                       'save group id
                Row("PrivilegeID") = CStr(PrivilegeArray.Item(I))       'save privilege id
                DBView.Table.Rows.Add(Row)                              'add row to table
            Next
            UpdateData("SELECT * FROM GroupPrivilegeTable", "GroupPrivilegeTable") 'update to the database
        End If

        Call RefreshAll()          '' load info form database to all display component

    End Sub
#End Region

#Region "Tree UserID"

     ''
    '  after checking tree userid item 
     ''''' 
    Private Sub TreeUserID_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeUserID.AfterSelect
        Dim I As Integer

        If TreeUserID.SelectedNode.Text = "IDS.data Users" Then 'if selected node is the primary node
            Return                                              'exit function
        End If

        DBView = New DataView(DS.Tables("UserTable"))           'open user table
        DBView.RowFilter = ""
        For I = 0 To DBView.Count - 1                           'iterate thru the user table items
            Dim Row As DataRow = DBView(I).Row

            'selected node from treeuserid = userid from the user table
            If TreeUserID.SelectedNode.Text = Row("UserID") Or TreeUserID.SelectedNode.Parent.Text = Row("UserID") Then

                ''''''''''''''''''''''''''''''
                ' validate the user info
                ''''''''''''''''''''''''''''''

                TextUserID.Text = Row("UserID")

                If IsDBNull(Row("UserName")) = False Then
                    TextName.Text = Row("UserName")
                Else
                    TextName.Text = ""
                End If
                If IsDBNull(Row("Department")) = False Then
                    TextDepartment.Text = Row("Department")
                Else
                    TextDepartment.Text = ""
                End If
                If IsDBNull(Row("ContactNo")) = False Then
                    TextContact.Text = Row("ContactNo")
                Else
                    TextContact.Text = ""
                End If
                If IsDBNull(Row("UserPassword")) = False Then
                    TextPassWord.Text = Row("UserPassword")
                Else
                    TextPassWord.Text = ""
                End If


                 '''''''''''''''''''
                'open group user table and filter for selected user id
                 '''''''''''''''''''
                Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable"))
                SubView.RowFilter = "UserID = '" + Row("UserID") + "'"

                ListUserGroup.Items.Clear()  'clear item
                If SubView.Count > 0 Then
                    Dim J As Integer
                    Dim Flag As Boolean = False
                    For J = 0 To SubView.Count - 1  'iterate thru the groups that fall under the selected user id
                        Dim GroupRow As DataRow = SubView(J).Row
                        ListUserGroup.Items.Add(GroupRow("GroupID"))    'add the group id to the list box

                         '
                        ' check if userid are administrator
                         '
                        If Flag = False Then
                            Dim s As String = GroupRow("GroupID")

                            If s.ToUpper = "ADMINISTRATION" Then 'if yes
                                BtnNewUser.Enabled = False      'disable the buttons
                                BtnDelUser.Enabled = False      'disable the button
                                Flag = True
                            Else
                                BtnNewUser.Enabled = True
                                BtnDelUser.Enabled = True
                            End If
                        End If

                         
                    Next
                End If
            End If
        Next
    End Sub
     '''''''''''''''''''''''''
    '' add groups to listusergroup list box
     '''''''''''''''''''''''''
    Private Sub BtnAddGrouptoUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddGrouptoUser.Click
        Dim I As Integer
        If ListAllGroup.SelectedItems.Count > 0 Then 'more than 1 item is selected

            I = 0
            For I = 0 To ListAllGroup.SelectedItems.Count - 1 'iterate thru the selected items

                If ListUserGroup.Items.Contains(ListAllGroup.SelectedItems.Item(I)) = True Then 'if the selected item is already in the listusergroup 
                Else
                    If ListAllGroup.SelectedItems.Item(I).ToUpper = "ADMINISTRATION" Then 'not exist, check is selected group under adminstration
                        'dun add administration 
                    Else
                        ListUserGroup.Items.Add(ListAllGroup.SelectedItems.Item(I)) 'no, add to the listusergroup item
                    End If
                End If
            Next
        End If
        ListAllGroup.ClearSelected()
    End Sub
     '''''''''''''''''''''''''
    '' delete groups from listusergroup list box
     '''''''''''''''''''''''''
    Private Sub BtnDelGrouptoUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelGrouptoUser.Click
        Dim I As Integer
        If ListUserGroup.SelectedItems.Count > 0 Then   'more than 1 item is selected
            Dim A(ListUserGroup.SelectedItems.Count - 1) As String 'declare a temp array
            ListUserGroup.SelectedItems.CopyTo(A, 0)    'copy selected item to array

            I = 0
            For I = 0 To A.Length - 1                   'iterate thru the array 
                If A(I).ToUpper = "ADMINISTRATION" Then
                Else
                    ListUserGroup.Items.Remove(A(I))    'del the item if user is not adminstration
                End If
            Next I
        End If
    End Sub
     '''''''''''''''''''''''''
    '' create new user id
     '''''''''''''''''''''''''
    Private Sub BtnNewUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnNewUser.Click
        Call RefreshAll()       '' load info form database to all display component
        Call ClearAll()         '' clear all display component
    End Sub
     '''''''''''''''''''''''''
    '' del user id form the database
     '''''''''''''''''''''''''
    Private Sub BtnDelUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelUser.Click
        Dim I As Integer


        ''''''''''''''''''''''''''''''
        ' validate the user info
        ''''''''''''''''''''''''''''''

        IDS.Data.MsgErr = ""
        If TextUserID.Text = "" Then
            IDS.Data.MsgErr = "UserID cant be blank"
            MessageBox.Show(IDS.Data.MsgErr)
            Return
        End If
        ''''''''''''''''''''''''''''''
        
        DBView = New DataView(DS.Tables("UserTable"))       'open user table
        DBView.RowFilter = "UserID = '" + TextUserID.Text + "'" 'filter for selected user id

        If DBView.Count = 0 Then
        Else

            Dim Respond As DialogResult = MessageBox.Show("Are you sure to Delete UserID " + TextUserID.Text + "?", "", MessageBoxButtons.YesNo)
            If Respond <> DialogResult.Yes Then
                Return
            Else
                Dim row As DataRow = DBView(0).Row          'get row info
                row.Delete()                                'del row
            End If
        End If
        UpdateData("SELECT * FROM UserTable", "UserTable")  'update to database

        Call RefreshAll()                                   '' load info form database to all display component
        Call ClearAll()                                     '' clear all display component
    End Sub
     '''''''''''''''''''''''''
    '' cancel changes at the display 
     '''''''''''''''''''''''''
    Private Sub BtnUserCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUserCancel.Click
        Dim I As Integer

        DBView = New DataView(DS.Tables("UserTable"))           'open user table
        DBView.RowFilter = "UserID = '" + TextUserID.Text + "'" 'filter for selected user id

        If DBView.Count > 0 Then                                'more than 1 record found
            Dim Row As DataRow = DBView(0).Row
            TextUserID.Text = Row("UserID")

            ''''''''''''''''''''''''''''''
            ' validate the user info
            '''''''''''''''''''''''''''''

            If IsDBNull(Row("UserName")) = False Then
                TextName.Text = Row("UserName")
            Else
                TextName.Text = ""
            End If
            If IsDBNull(Row("Department")) = False Then
                TextDepartment.Text = Row("Department")
            Else
                TextDepartment.Text = ""
            End If
            If IsDBNull(Row("ContactNo")) = False Then
                TextContact.Text = Row("ContactNo")
            Else
                TextContact.Text = ""
            End If
            If IsDBNull(Row("UserPassword")) = False Then
                TextPassWord.Text = Row("UserPassword")
            Else
                TextPassWord.Text = ""
            End If


             ''''''''''''''''''
            ' open group user table filter with the selected user id
             '''''''''''''''''''
            Dim SubView As DataView = New DataView(DS.Tables("GroupUserTable"))
            SubView.RowFilter = "UserID = '" + Row("UserID") + "'"

            ListUserGroup.Items.Clear() 'clear the listusergroup items
            If SubView.Count > 0 Then
                Dim J As Integer
                For J = 0 To SubView.Count - 1
                    Dim GroupRow As DataRow = SubView(J).Row    'get the row info (groupid)
                    ListUserGroup.Items.Add(GroupRow("GroupID")) 'add to the listusergroup items
                Next
            End If
        End If
    End Sub
     '''''''''''''''''''''''''
    '' apply the changes at the display and update to the database
     '''''''''''''''''''''''''
    Private Sub btnUserApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserApply.Click
        Dim I As Integer


        ''''''''''''''''''''''''''''''
        ' validate the user info
        ''''''''''''''''''''''''''''''
        IDS.Data.MsgErr = ""
        If TextUserID.Text = "" Then
            IDS.Data.MsgErr = "UserID cant be blank"
            MessageBox.Show(IDS.Data.MsgErr)
            Return
        End If


        DBView = New DataView(DS.Tables("UserTable"))   'open user table with filter for selected user id
        DBView.RowFilter = "UserID = '" + TextUserID.Text + "'"
        IDS.Data.MsgErr = ""
        If DBView.Count = 0 Then
        Else
            'existing record found, do u wan to overwrite
            Dim Respond As DialogResult = MessageBox.Show("There is an existing UserID of " + TextUserID.Text + ". Do you want to overwrite it?", "", MessageBoxButtons.YesNo)
            If Respond <> DialogResult.Yes Then
                Return  'no
            Else
                Dim row As DataRow = DBView(0).Row
                row.Delete() 'yes, delete it 
                UpdateData("SELECT * FROM UserTable", "UserTable")  'and update to database
            End If
        End If

         ''''''''''''''''
        ' create new row and fill it with info from the dislay
         ''''''''''''''''
        Dim Newrow As DataRow = DBView.Table.NewRow()
        Newrow("UserID") = TextUserID.Text
        Newrow("UserName") = TextName.Text
        Newrow("Department") = TextDepartment.Text
        Newrow("ContactNo") = TextContact.Text
        Newrow("UserPassword") = TextPassWord.Text
        DBView.Table.Rows.Add(Newrow)                       'add row to the table     
        UpdateData("SELECT * FROM UserTable", "UserTable") 'update table to the database


         ''''''''''''''''
        ' open group user table with filter for selected user id
         ''''''''''''''''
        DBView = New DataView(DS.Tables("GroupUserTable"))
        DBView.RowFilter = "UserID = '" + TextUserID.Text + "'"

        If ListUserGroup.Items.Count > 0 Then
            I = 0
            For I = 0 To ListUserGroup.Items.Count - 1
                Dim Row As DataRow = DBView.Table.NewRow() 'create new row
                Row("UserID") = TextUserID.Text             'save user id info
                Row("GroupID") = ListUserGroup.Items(I)     'save group id info
                DBView.Table.Rows.Add(Row)                  'add row to table
            Next
            UpdateData("SELECT * FROM GroupUserTable", "GroupUserTable") 'update table to the database
        End If
        Call RefreshAll()       '' load info form database to all display component
    End Sub
#End Region

    Private Sub FormAdministration_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        btnGroupCancel_Click(sender, e)
    End Sub
End Class
