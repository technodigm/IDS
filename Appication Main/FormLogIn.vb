'Option Strict On
Option Explicit On 
Imports Microsoft.Win32
Imports System.IO
Imports System
Imports System.Data.OleDb
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports DLL_Export_Service
Imports DLL_Execution
Imports DLL_SetupAndCalibrate
Imports System.Math
Imports DLL_DataManager

Public Class FormLogin
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()



        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'IDS.Data.ParameterID.RecordID = "FactoryDefault"
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents TextLoginPW As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents TextCurrentUSerID As System.Windows.Forms.TextBox
    Friend WithEvents TextNewUserID As System.Windows.Forms.TextBox
    Friend WithEvents TextConfirmUserID As System.Windows.Forms.TextBox
    Friend WithEvents BtnWelcomeChangePW As System.Windows.Forms.Button
    Friend WithEvents BtnWelcomeChangeUserID As System.Windows.Forms.Button
    Friend WithEvents BtnLogin As System.Windows.Forms.Button
    Friend WithEvents LUserName As System.Windows.Forms.Label
    Friend WithEvents LDepartment As System.Windows.Forms.Label
    Friend WithEvents LContact As System.Windows.Forms.Label
    Friend WithEvents LUserID As System.Windows.Forms.Label
    Friend WithEvents LGroupID As System.Windows.Forms.Label
    Friend WithEvents PanelIDSLogin As System.Windows.Forms.Panel
    Friend WithEvents Panelwelcome As System.Windows.Forms.Panel
    Friend WithEvents BtnWelcomeLogout As System.Windows.Forms.Button
    Friend WithEvents PanelChangeUserID As System.Windows.Forms.Panel
    Friend WithEvents PanelChangePW As System.Windows.Forms.Panel
    Friend WithEvents BtnChangeUserIDBack As System.Windows.Forms.Button
    Friend WithEvents BtnchangeUserID As System.Windows.Forms.Button
    Friend WithEvents BtnChangePWBack As System.Windows.Forms.Button
    Friend WithEvents btnchangePW As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextcurrentPW As System.Windows.Forms.TextBox
    Friend WithEvents TextNewPW As System.Windows.Forms.TextBox
    Friend WithEvents TextConfirmPW As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents CBUserID As System.Windows.Forms.ComboBox
    Friend WithEvents CBGroupID As System.Windows.Forms.ComboBox
    Friend WithEvents PanelSelectApplication As System.Windows.Forms.Panel
    Friend WithEvents BtnSelectapplcationBack As System.Windows.Forms.Button
    Friend WithEvents BtnSelectApplicationOK As System.Windows.Forms.Button
    Friend WithEvents RadioOperator As System.Windows.Forms.RadioButton
    Friend WithEvents RadioProgrammer As System.Windows.Forms.RadioButton
    Friend WithEvents RadioMaintenance As System.Windows.Forms.RadioButton
    Friend WithEvents RadioSystem As System.Windows.Forms.RadioButton
    Friend WithEvents PanelSSetup As System.Windows.Forms.Panel
    Friend WithEvents BtnSSetupBack As System.Windows.Forms.Button
    Friend WithEvents BtnSSetupOK As System.Windows.Forms.Button
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextSSetupPassword As System.Windows.Forms.TextBox
    Friend WithEvents TextSSetupGroupID As System.Windows.Forms.TextBox
    Friend WithEvents TextSSetupUserID As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents CMPrimaryKey As System.Windows.Forms.ContextMenu
    Friend WithEvents CMGroupIDPrimaryKey As System.Windows.Forms.MenuItem
    Friend WithEvents CMUserIDPrimaryKey As System.Windows.Forms.MenuItem
    Friend WithEvents PanelUserID As System.Windows.Forms.Panel
    Friend WithEvents PanelgroupID As System.Windows.Forms.Panel
    Friend WithEvents ButtonSetupRegistry As System.Windows.Forms.Button
    Public WithEvents BtnWelcomeOK As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormLogin))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextLoginPW = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.CBGroupID = New System.Windows.Forms.ComboBox
        Me.TextCurrentUSerID = New System.Windows.Forms.TextBox
        Me.TextNewUserID = New System.Windows.Forms.TextBox
        Me.TextConfirmUserID = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.CMPrimaryKey = New System.Windows.Forms.ContextMenu
        Me.CMGroupIDPrimaryKey = New System.Windows.Forms.MenuItem
        Me.CMUserIDPrimaryKey = New System.Windows.Forms.MenuItem
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LGroupID = New System.Windows.Forms.Label
        Me.LUserName = New System.Windows.Forms.Label
        Me.LDepartment = New System.Windows.Forms.Label
        Me.LContact = New System.Windows.Forms.Label
        Me.LUserID = New System.Windows.Forms.Label
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.BtnWelcomeChangePW = New System.Windows.Forms.Button
        Me.BtnWelcomeChangeUserID = New System.Windows.Forms.Button
        Me.BtnLogin = New System.Windows.Forms.Button
        Me.PanelIDSLogin = New System.Windows.Forms.Panel
        Me.ButtonSetupRegistry = New System.Windows.Forms.Button
        Me.PanelUserID = New System.Windows.Forms.Panel
        Me.CBUserID = New System.Windows.Forms.ComboBox
        Me.PanelgroupID = New System.Windows.Forms.Panel
        Me.Panelwelcome = New System.Windows.Forms.Panel
        Me.BtnWelcomeOK = New System.Windows.Forms.Button
        Me.BtnWelcomeLogout = New System.Windows.Forms.Button
        Me.PanelChangeUserID = New System.Windows.Forms.Panel
        Me.BtnChangeUserIDBack = New System.Windows.Forms.Button
        Me.BtnchangeUserID = New System.Windows.Forms.Button
        Me.PanelChangePW = New System.Windows.Forms.Panel
        Me.BtnChangePWBack = New System.Windows.Forms.Button
        Me.btnchangePW = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.TextcurrentPW = New System.Windows.Forms.TextBox
        Me.TextNewPW = New System.Windows.Forms.TextBox
        Me.TextConfirmPW = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.PanelSelectApplication = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.RadioOperator = New System.Windows.Forms.RadioButton
        Me.RadioProgrammer = New System.Windows.Forms.RadioButton
        Me.RadioMaintenance = New System.Windows.Forms.RadioButton
        Me.RadioSystem = New System.Windows.Forms.RadioButton
        Me.BtnSelectapplcationBack = New System.Windows.Forms.Button
        Me.BtnSelectApplicationOK = New System.Windows.Forms.Button
        Me.PanelSSetup = New System.Windows.Forms.Panel
        Me.Label11 = New System.Windows.Forms.Label
        Me.TextSSetupUserID = New System.Windows.Forms.TextBox
        Me.TextSSetupGroupID = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.TextSSetupPassword = New System.Windows.Forms.TextBox
        Me.BtnSSetupBack = New System.Windows.Forms.Button
        Me.BtnSSetupOK = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.PanelIDSLogin.SuspendLayout()
        Me.PanelUserID.SuspendLayout()
        Me.PanelgroupID.SuspendLayout()
        Me.Panelwelcome.SuspendLayout()
        Me.PanelChangeUserID.SuspendLayout()
        Me.PanelChangePW.SuspendLayout()
        Me.PanelSelectApplication.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.PanelSSetup.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(152, 24)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(40, 24)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 52
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "UserID"
        '
        'TextLoginPW
        '
        Me.TextLoginPW.Location = New System.Drawing.Point(88, 152)
        Me.TextLoginPW.Name = "TextLoginPW"
        Me.TextLoginPW.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TextLoginPW.Size = New System.Drawing.Size(176, 21)
        Me.TextLoginPW.TabIndex = 2
        Me.TextLoginPW.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 152)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Password"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 16)
        Me.Label6.TabIndex = 49
        Me.Label6.Text = "Group ID"
        '
        'CBGroupID
        '
        Me.CBGroupID.Location = New System.Drawing.Point(72, 8)
        Me.CBGroupID.Name = "CBGroupID"
        Me.CBGroupID.Size = New System.Drawing.Size(176, 23)
        Me.CBGroupID.Sorted = True
        Me.CBGroupID.TabIndex = 0
        Me.CBGroupID.Text = "Application_ALL"
        '
        'TextCurrentUSerID
        '
        Me.TextCurrentUSerID.Location = New System.Drawing.Point(152, 48)
        Me.TextCurrentUSerID.Name = "TextCurrentUSerID"
        Me.TextCurrentUSerID.Size = New System.Drawing.Size(136, 21)
        Me.TextCurrentUSerID.TabIndex = 0
        Me.TextCurrentUSerID.Text = ""
        '
        'TextNewUserID
        '
        Me.TextNewUserID.Location = New System.Drawing.Point(152, 88)
        Me.TextNewUserID.Name = "TextNewUserID"
        Me.TextNewUserID.Size = New System.Drawing.Size(136, 21)
        Me.TextNewUserID.TabIndex = 1
        Me.TextNewUserID.Text = ""
        '
        'TextConfirmUserID
        '
        Me.TextConfirmUserID.Location = New System.Drawing.Point(152, 128)
        Me.TextConfirmUserID.Name = "TextConfirmUserID"
        Me.TextConfirmUserID.Size = New System.Drawing.Size(136, 21)
        Me.TextConfirmUserID.TabIndex = 2
        Me.TextConfirmUserID.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 16)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Current UserID"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(32, 88)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(112, 16)
        Me.Label9.TabIndex = 5
        Me.Label9.Text = "New UserID"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Confirm UserID"
        '
        'CMPrimaryKey
        '
        Me.CMPrimaryKey.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.CMGroupIDPrimaryKey, Me.CMUserIDPrimaryKey})
        '
        'CMGroupIDPrimaryKey
        '
        Me.CMGroupIDPrimaryKey.Index = 0
        Me.CMGroupIDPrimaryKey.Text = "Group ID as Primary Key"
        '
        'CMUserIDPrimaryKey
        '
        Me.CMUserIDPrimaryKey.Index = 1
        Me.CMUserIDPrimaryKey.Text = "User ID as Primary Key"
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
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LGroupID)
        Me.GroupBox1.Controls.Add(Me.LUserName)
        Me.GroupBox1.Controls.Add(Me.LDepartment)
        Me.GroupBox1.Controls.Add(Me.LContact)
        Me.GroupBox1.Controls.Add(Me.LUserID)
        Me.GroupBox1.Controls.Add(Me.PictureBox2)
        Me.GroupBox1.Controls.Add(Me.BtnWelcomeChangePW)
        Me.GroupBox1.Controls.Add(Me.BtnWelcomeChangeUserID)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(4, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(312, 208)
        Me.GroupBox1.TabIndex = 58
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "User Details"
        '
        'LGroupID
        '
        Me.LGroupID.Location = New System.Drawing.Point(74, 32)
        Me.LGroupID.Name = "LGroupID"
        Me.LGroupID.Size = New System.Drawing.Size(206, 16)
        Me.LGroupID.TabIndex = 68
        Me.LGroupID.Text = "GroupID"
        '
        'LUserName
        '
        Me.LUserName.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.LUserName.Location = New System.Drawing.Point(74, 86)
        Me.LUserName.Name = "LUserName"
        Me.LUserName.Size = New System.Drawing.Size(206, 16)
        Me.LUserName.TabIndex = 63
        Me.LUserName.Text = "UserName"
        '
        'LDepartment
        '
        Me.LDepartment.Location = New System.Drawing.Point(74, 113)
        Me.LDepartment.Name = "LDepartment"
        Me.LDepartment.Size = New System.Drawing.Size(206, 16)
        Me.LDepartment.TabIndex = 65
        Me.LDepartment.Text = "Department"
        '
        'LContact
        '
        Me.LContact.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.LContact.Location = New System.Drawing.Point(74, 140)
        Me.LContact.Name = "LContact"
        Me.LContact.Size = New System.Drawing.Size(206, 16)
        Me.LContact.TabIndex = 67
        Me.LContact.Text = "Contact No"
        '
        'LUserID
        '
        Me.LUserID.Location = New System.Drawing.Point(74, 59)
        Me.LUserID.Name = "LUserID"
        Me.LUserID.Size = New System.Drawing.Size(206, 16)
        Me.LUserID.TabIndex = 59
        Me.LUserID.Text = "User ID"
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(20, 36)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 55
        Me.PictureBox2.TabStop = False
        '
        'BtnWelcomeChangePW
        '
        Me.BtnWelcomeChangePW.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnWelcomeChangePW.Location = New System.Drawing.Point(164, 168)
        Me.BtnWelcomeChangePW.Name = "BtnWelcomeChangePW"
        Me.BtnWelcomeChangePW.Size = New System.Drawing.Size(116, 32)
        Me.BtnWelcomeChangePW.TabIndex = 0
        Me.BtnWelcomeChangePW.Text = "ChangePassword"
        '
        'BtnWelcomeChangeUserID
        '
        Me.BtnWelcomeChangeUserID.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnWelcomeChangeUserID.Location = New System.Drawing.Point(33, 167)
        Me.BtnWelcomeChangeUserID.Name = "BtnWelcomeChangeUserID"
        Me.BtnWelcomeChangeUserID.Size = New System.Drawing.Size(105, 32)
        Me.BtnWelcomeChangeUserID.TabIndex = 1
        Me.BtnWelcomeChangeUserID.Text = "ChangeUserID"
        '
        'BtnLogin
        '
        Me.BtnLogin.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnLogin.Location = New System.Drawing.Point(112, 192)
        Me.BtnLogin.Name = "BtnLogin"
        Me.BtnLogin.Size = New System.Drawing.Size(88, 32)
        Me.BtnLogin.TabIndex = 3
        Me.BtnLogin.Text = "Login >>"
        '
        'PanelIDSLogin
        '
        Me.PanelIDSLogin.ContextMenu = Me.CMPrimaryKey
        Me.PanelIDSLogin.Controls.Add(Me.BtnLogin)
        Me.PanelIDSLogin.Controls.Add(Me.ButtonSetupRegistry)
        Me.PanelIDSLogin.Controls.Add(Me.PanelUserID)
        Me.PanelIDSLogin.Controls.Add(Me.PanelgroupID)
        Me.PanelIDSLogin.Controls.Add(Me.Label2)
        Me.PanelIDSLogin.Controls.Add(Me.PictureBox1)
        Me.PanelIDSLogin.Controls.Add(Me.TextLoginPW)
        Me.PanelIDSLogin.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelIDSLogin.Location = New System.Drawing.Point(0, 0)
        Me.PanelIDSLogin.Name = "PanelIDSLogin"
        Me.PanelIDSLogin.Size = New System.Drawing.Size(328, 296)
        Me.PanelIDSLogin.TabIndex = 59
        Me.ToolTip1.SetToolTip(Me.PanelIDSLogin, "Right click to select User ID or Group ID as SleSelect GroupID as Primary Key")
        '
        'ButtonSetupRegistry
        '
        Me.ButtonSetupRegistry.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonSetupRegistry.Location = New System.Drawing.Point(112, 192)
        Me.ButtonSetupRegistry.Name = "ButtonSetupRegistry"
        Me.ButtonSetupRegistry.Size = New System.Drawing.Size(80, 32)
        Me.ButtonSetupRegistry.TabIndex = 55
        Me.ButtonSetupRegistry.Text = "Setup &Registry"
        '
        'PanelUserID
        '
        Me.PanelUserID.Controls.Add(Me.Label1)
        Me.PanelUserID.Controls.Add(Me.CBUserID)
        Me.PanelUserID.Location = New System.Drawing.Point(16, 104)
        Me.PanelUserID.Name = "PanelUserID"
        Me.PanelUserID.Size = New System.Drawing.Size(264, 40)
        Me.PanelUserID.TabIndex = 54
        '
        'CBUserID
        '
        Me.CBUserID.ItemHeight = 15
        Me.CBUserID.Location = New System.Drawing.Point(72, 8)
        Me.CBUserID.Name = "CBUserID"
        Me.CBUserID.Size = New System.Drawing.Size(176, 23)
        Me.CBUserID.Sorted = True
        Me.CBUserID.TabIndex = 1
        Me.CBUserID.Text = "a"
        '
        'PanelgroupID
        '
        Me.PanelgroupID.Controls.Add(Me.CBGroupID)
        Me.PanelgroupID.Controls.Add(Me.Label6)
        Me.PanelgroupID.Location = New System.Drawing.Point(16, 64)
        Me.PanelgroupID.Name = "PanelgroupID"
        Me.PanelgroupID.Size = New System.Drawing.Size(264, 40)
        Me.PanelgroupID.TabIndex = 53
        '
        'Panelwelcome
        '
        Me.Panelwelcome.Controls.Add(Me.BtnWelcomeOK)
        Me.Panelwelcome.Controls.Add(Me.BtnWelcomeLogout)
        Me.Panelwelcome.Controls.Add(Me.GroupBox1)
        Me.Panelwelcome.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panelwelcome.Location = New System.Drawing.Point(344, 0)
        Me.Panelwelcome.Name = "Panelwelcome"
        Me.Panelwelcome.Size = New System.Drawing.Size(328, 296)
        Me.Panelwelcome.TabIndex = 60
        Me.Panelwelcome.Visible = False
        '
        'BtnWelcomeOK
        '
        Me.BtnWelcomeOK.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnWelcomeOK.Location = New System.Drawing.Point(232, 224)
        Me.BtnWelcomeOK.Name = "BtnWelcomeOK"
        Me.BtnWelcomeOK.Size = New System.Drawing.Size(80, 32)
        Me.BtnWelcomeOK.TabIndex = 0
        Me.BtnWelcomeOK.Text = "OK"
        '
        'BtnWelcomeLogout
        '
        Me.BtnWelcomeLogout.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnWelcomeLogout.Location = New System.Drawing.Point(136, 224)
        Me.BtnWelcomeLogout.Name = "BtnWelcomeLogout"
        Me.BtnWelcomeLogout.Size = New System.Drawing.Size(88, 32)
        Me.BtnWelcomeLogout.TabIndex = 1
        Me.BtnWelcomeLogout.Text = "Logout<<"
        '
        'PanelChangeUserID
        '
        Me.PanelChangeUserID.Controls.Add(Me.BtnChangeUserIDBack)
        Me.PanelChangeUserID.Controls.Add(Me.BtnchangeUserID)
        Me.PanelChangeUserID.Controls.Add(Me.Label9)
        Me.PanelChangeUserID.Controls.Add(Me.Label4)
        Me.PanelChangeUserID.Controls.Add(Me.TextCurrentUSerID)
        Me.PanelChangeUserID.Controls.Add(Me.TextNewUserID)
        Me.PanelChangeUserID.Controls.Add(Me.TextConfirmUserID)
        Me.PanelChangeUserID.Controls.Add(Me.Label3)
        Me.PanelChangeUserID.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelChangeUserID.Location = New System.Drawing.Point(0, 304)
        Me.PanelChangeUserID.Name = "PanelChangeUserID"
        Me.PanelChangeUserID.Size = New System.Drawing.Size(328, 272)
        Me.PanelChangeUserID.TabIndex = 61
        Me.PanelChangeUserID.Visible = False
        '
        'BtnChangeUserIDBack
        '
        Me.BtnChangeUserIDBack.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnChangeUserIDBack.Location = New System.Drawing.Point(48, 184)
        Me.BtnChangeUserIDBack.Name = "BtnChangeUserIDBack"
        Me.BtnChangeUserIDBack.Size = New System.Drawing.Size(96, 32)
        Me.BtnChangeUserIDBack.TabIndex = 4
        Me.BtnChangeUserIDBack.Text = "Back <<"
        '
        'BtnchangeUserID
        '
        Me.BtnchangeUserID.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnchangeUserID.Location = New System.Drawing.Point(160, 184)
        Me.BtnchangeUserID.Name = "BtnchangeUserID"
        Me.BtnchangeUserID.Size = New System.Drawing.Size(120, 32)
        Me.BtnchangeUserID.TabIndex = 3
        Me.BtnchangeUserID.Text = "ChangeUserID"
        '
        'PanelChangePW
        '
        Me.PanelChangePW.Controls.Add(Me.BtnChangePWBack)
        Me.PanelChangePW.Controls.Add(Me.btnchangePW)
        Me.PanelChangePW.Controls.Add(Me.Label5)
        Me.PanelChangePW.Controls.Add(Me.Label12)
        Me.PanelChangePW.Controls.Add(Me.TextcurrentPW)
        Me.PanelChangePW.Controls.Add(Me.TextNewPW)
        Me.PanelChangePW.Controls.Add(Me.TextConfirmPW)
        Me.PanelChangePW.Controls.Add(Me.Label13)
        Me.PanelChangePW.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelChangePW.Location = New System.Drawing.Point(344, 304)
        Me.PanelChangePW.Name = "PanelChangePW"
        Me.PanelChangePW.Size = New System.Drawing.Size(328, 272)
        Me.PanelChangePW.TabIndex = 62
        Me.PanelChangePW.Visible = False
        '
        'BtnChangePWBack
        '
        Me.BtnChangePWBack.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnChangePWBack.Location = New System.Drawing.Point(56, 184)
        Me.BtnChangePWBack.Name = "BtnChangePWBack"
        Me.BtnChangePWBack.Size = New System.Drawing.Size(96, 32)
        Me.BtnChangePWBack.TabIndex = 4
        Me.BtnChangePWBack.Text = "Back <<"
        '
        'btnchangePW
        '
        Me.btnchangePW.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnchangePW.Location = New System.Drawing.Point(168, 184)
        Me.btnchangePW.Name = "btnchangePW"
        Me.btnchangePW.Size = New System.Drawing.Size(120, 32)
        Me.btnchangePW.TabIndex = 3
        Me.btnchangePW.Text = "ChangePassword"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(36, 92)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 16)
        Me.Label5.TabIndex = 57
        Me.Label5.Text = "New Password"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(36, 132)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(116, 16)
        Me.Label12.TabIndex = 55
        Me.Label12.Text = "Confirm Password"
        '
        'TextcurrentPW
        '
        Me.TextcurrentPW.Location = New System.Drawing.Point(156, 52)
        Me.TextcurrentPW.Name = "TextcurrentPW"
        Me.TextcurrentPW.Size = New System.Drawing.Size(136, 21)
        Me.TextcurrentPW.TabIndex = 0
        Me.TextcurrentPW.Text = ""
        '
        'TextNewPW
        '
        Me.TextNewPW.Location = New System.Drawing.Point(156, 92)
        Me.TextNewPW.Name = "TextNewPW"
        Me.TextNewPW.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TextNewPW.Size = New System.Drawing.Size(136, 21)
        Me.TextNewPW.TabIndex = 1
        Me.TextNewPW.Text = ""
        '
        'TextConfirmPW
        '
        Me.TextConfirmPW.Location = New System.Drawing.Point(156, 132)
        Me.TextConfirmPW.Name = "TextConfirmPW"
        Me.TextConfirmPW.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TextConfirmPW.Size = New System.Drawing.Size(136, 21)
        Me.TextConfirmPW.TabIndex = 2
        Me.TextConfirmPW.Text = ""
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(36, 52)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(112, 16)
        Me.Label13.TabIndex = 59
        Me.Label13.Text = "Current Password"
        '
        'PanelSelectApplication
        '
        Me.PanelSelectApplication.Controls.Add(Me.GroupBox2)
        Me.PanelSelectApplication.Controls.Add(Me.BtnSelectapplcationBack)
        Me.PanelSelectApplication.Controls.Add(Me.BtnSelectApplicationOK)
        Me.PanelSelectApplication.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelSelectApplication.Location = New System.Drawing.Point(672, 0)
        Me.PanelSelectApplication.Name = "PanelSelectApplication"
        Me.PanelSelectApplication.Size = New System.Drawing.Size(328, 272)
        Me.PanelSelectApplication.TabIndex = 63
        Me.PanelSelectApplication.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RadioOperator)
        Me.GroupBox2.Controls.Add(Me.RadioProgrammer)
        Me.GroupBox2.Controls.Add(Me.RadioMaintenance)
        Me.GroupBox2.Controls.Add(Me.RadioSystem)
        Me.GroupBox2.Location = New System.Drawing.Point(16, 40)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(296, 152)
        Me.GroupBox2.TabIndex = 9
        Me.GroupBox2.TabStop = False
        '
        'RadioOperator
        '
        Me.RadioOperator.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioOperator.Location = New System.Drawing.Point(24, 88)
        Me.RadioOperator.Name = "RadioOperator"
        Me.RadioOperator.Size = New System.Drawing.Size(232, 24)
        Me.RadioOperator.TabIndex = 8
        Me.RadioOperator.Text = "Go to Operator Application"
        '
        'RadioProgrammer
        '
        Me.RadioProgrammer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioProgrammer.Location = New System.Drawing.Point(24, 56)
        Me.RadioProgrammer.Name = "RadioProgrammer"
        Me.RadioProgrammer.Size = New System.Drawing.Size(232, 24)
        Me.RadioProgrammer.TabIndex = 7
        Me.RadioProgrammer.Text = "Go to Programmer Application"
        '
        'RadioMaintenance
        '
        Me.RadioMaintenance.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioMaintenance.Location = New System.Drawing.Point(24, 24)
        Me.RadioMaintenance.Name = "RadioMaintenance"
        Me.RadioMaintenance.Size = New System.Drawing.Size(232, 24)
        Me.RadioMaintenance.TabIndex = 6
        Me.RadioMaintenance.Text = "Go to Maintenance Application"
        '
        'RadioSystem
        '
        Me.RadioSystem.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioSystem.Location = New System.Drawing.Point(24, 120)
        Me.RadioSystem.Name = "RadioSystem"
        Me.RadioSystem.Size = New System.Drawing.Size(256, 24)
        Me.RadioSystem.TabIndex = 5
        Me.RadioSystem.Text = "Go to System Configuration Application"
        Me.RadioSystem.Visible = False
        '
        'BtnSelectapplcationBack
        '
        Me.BtnSelectapplcationBack.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSelectapplcationBack.Location = New System.Drawing.Point(104, 208)
        Me.BtnSelectapplcationBack.Name = "BtnSelectapplcationBack"
        Me.BtnSelectapplcationBack.Size = New System.Drawing.Size(96, 32)
        Me.BtnSelectapplcationBack.TabIndex = 4
        Me.BtnSelectapplcationBack.Text = "Back <<"
        '
        'BtnSelectApplicationOK
        '
        Me.BtnSelectApplicationOK.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSelectApplicationOK.Location = New System.Drawing.Point(208, 208)
        Me.BtnSelectApplicationOK.Name = "BtnSelectApplicationOK"
        Me.BtnSelectApplicationOK.Size = New System.Drawing.Size(80, 32)
        Me.BtnSelectApplicationOK.TabIndex = 3
        Me.BtnSelectApplicationOK.Text = "OK"
        '
        'PanelSSetup
        '
        Me.PanelSSetup.Controls.Add(Me.Label11)
        Me.PanelSSetup.Controls.Add(Me.TextSSetupUserID)
        Me.PanelSSetup.Controls.Add(Me.TextSSetupGroupID)
        Me.PanelSSetup.Controls.Add(Me.Label8)
        Me.PanelSSetup.Controls.Add(Me.Label10)
        Me.PanelSSetup.Controls.Add(Me.TextSSetupPassword)
        Me.PanelSSetup.Controls.Add(Me.BtnSSetupBack)
        Me.PanelSSetup.Controls.Add(Me.BtnSSetupOK)
        Me.PanelSSetup.Controls.Add(Me.Label7)
        Me.PanelSSetup.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PanelSSetup.Location = New System.Drawing.Point(680, 304)
        Me.PanelSSetup.Name = "PanelSSetup"
        Me.PanelSSetup.Size = New System.Drawing.Size(328, 272)
        Me.PanelSSetup.TabIndex = 64
        Me.PanelSSetup.Visible = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(40, 32)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(248, 24)
        Me.Label11.TabIndex = 58
        Me.Label11.Text = "System Setup Authentication"
        '
        'TextSSetupUserID
        '
        Me.TextSSetupUserID.BackColor = System.Drawing.SystemColors.Info
        Me.TextSSetupUserID.Location = New System.Drawing.Point(128, 120)
        Me.TextSSetupUserID.Name = "TextSSetupUserID"
        Me.TextSSetupUserID.ReadOnly = True
        Me.TextSSetupUserID.Size = New System.Drawing.Size(152, 21)
        Me.TextSSetupUserID.TabIndex = 57
        Me.TextSSetupUserID.Text = ""
        '
        'TextSSetupGroupID
        '
        Me.TextSSetupGroupID.BackColor = System.Drawing.SystemColors.Info
        Me.TextSSetupGroupID.Location = New System.Drawing.Point(128, 80)
        Me.TextSSetupGroupID.Name = "TextSSetupGroupID"
        Me.TextSSetupGroupID.ReadOnly = True
        Me.TextSSetupGroupID.Size = New System.Drawing.Size(152, 21)
        Me.TextSSetupGroupID.TabIndex = 56
        Me.TextSSetupGroupID.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(32, 80)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.TabIndex = 55
        Me.Label8.Text = "Group ID"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(32, 120)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 16)
        Me.Label10.TabIndex = 51
        Me.Label10.Text = "UserID"
        '
        'TextSSetupPassword
        '
        Me.TextSSetupPassword.Location = New System.Drawing.Point(128, 160)
        Me.TextSSetupPassword.Name = "TextSSetupPassword"
        Me.TextSSetupPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.TextSSetupPassword.Size = New System.Drawing.Size(152, 21)
        Me.TextSSetupPassword.TabIndex = 54
        Me.TextSSetupPassword.Text = ""
        '
        'BtnSSetupBack
        '
        Me.BtnSSetupBack.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSSetupBack.Location = New System.Drawing.Point(104, 208)
        Me.BtnSSetupBack.Name = "BtnSSetupBack"
        Me.BtnSSetupBack.Size = New System.Drawing.Size(96, 32)
        Me.BtnSSetupBack.TabIndex = 4
        Me.BtnSSetupBack.Text = "Back <<"
        '
        'BtnSSetupOK
        '
        Me.BtnSSetupOK.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.BtnSSetupOK.Location = New System.Drawing.Point(208, 208)
        Me.BtnSSetupOK.Name = "BtnSSetupOK"
        Me.BtnSSetupOK.Size = New System.Drawing.Size(80, 32)
        Me.BtnSSetupOK.TabIndex = 3
        Me.BtnSSetupOK.Text = "OK"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(32, 160)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 32)
        Me.Label7.TabIndex = 53
        Me.Label7.Text = "Authentication Code"
        '
        'FormLogin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(992, 753)
        Me.Controls.Add(Me.PanelSSetup)
        Me.Controls.Add(Me.PanelSelectApplication)
        Me.Controls.Add(Me.PanelChangePW)
        Me.Controls.Add(Me.PanelChangeUserID)
        Me.Controls.Add(Me.Panelwelcome)
        Me.Controls.Add(Me.PanelIDSLogin)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.MinimizeBox = False
        Me.Name = "FormLogin"
        Me.Text = "Login"
        Me.GroupBox1.ResumeLayout(False)
        Me.PanelIDSLogin.ResumeLayout(False)
        Me.PanelUserID.ResumeLayout(False)
        Me.PanelgroupID.ResumeLayout(False)
        Me.Panelwelcome.ResumeLayout(False)
        Me.PanelChangeUserID.ResumeLayout(False)
        Me.PanelChangePW.ResumeLayout(False)
        Me.PanelSelectApplication.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.PanelSSetup.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " APIs "
    Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End Region

#Region " Form Events "
#Region " ... OnHandleCreated "
    Protected Overrides Sub OnHandleCreated(ByVal e As System.EventArgs)
        '---------------------------------------------------------------------------------
        '     Date          Developer                      Code Change       
        '  ---------- -------------------- -----------------------------------------------
        '  12/10/2005 
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

#Region "Main Login Panel "
    Dim SetGroupIDPriKey As Boolean 'working flag - to set group id as primary selection field
    Dim SetUserIDPriKey As Boolean  'working flag - to set user id as primary selection field
    '''''''''''''''''''''''''
    'select GroupID comobox at the top for selection
    '''''''''''''''''''''''''
    Private Sub CMGroupIDPrimaryKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMGroupIDPrimaryKey.Click

        ' CBGroupID.DataSource = DS
        ' CBGroupID.DisplayMember = "GroupTable.GroupID"
        SetGroupIDPriKey = True
        SetUserIDPriKey = False
        RefreshCBGroupID()
        CBUserID.Items.Clear()
        CBUserID.Text = ""
        CBGroupID.Text = ""

        Dim pt As Point
        pt.X = 16
        pt.Y = 64
        PanelgroupID.Location = pt

        pt.X = 16
        pt.Y = 104
        PanelUserID.Location = pt

    End Sub
    '''''''''''''''''''''''''
    'select UserID comobox at the top for selection
    '''''''''''''''''''''''''
    Private Sub CMUserIDPrimaryKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CMUserIDPrimaryKey.Click

        'CBUserID.DataSource = DS
        ' CBUserID.DisplayMember = "UserTable.UserID"
        SetGroupIDPriKey = False
        SetUserIDPriKey = True
        RefreshCBGroupID()
        CBGroupID.Items.Clear()
        CBGroupID.Text = ""
        CBUserID.Text = ""

        Dim pt As Point
        pt.X = 16
        pt.Y = 104
        PanelgroupID.Location = pt

        pt.X = 16
        pt.Y = 64
        PanelUserID.Location = pt

    End Sub
    'refresh the GroupID/userid combobox data item
    Private Sub RefreshCBGroupID()

        If SetGroupIDPriKey = True Then                     'groupid as primary selection field

            DBView = New DataView(DS.Tables("GroupTable"))  'load grouptable to dataview
            Dim I As Integer
            CBGroupID.Items.Clear()                         'clear combo box           
            If DBView.Count > 0 Then                        'load item from the dataview row into the combo box
                For I = 0 To DBView.Count - 1
                    Dim Row As DataRow = DBView(I).Row
                    CBGroupID.Items.Add(Row("GroupID"))
                Next
            End If

        ElseIf SetUserIDPriKey = True Then                  'userid as primary selection field

            DBView = New DataView(DS.Tables("UserTable"))   'load usertable to dataview
            Dim I As Integer
            CBUserID.Items.Clear()                          'clear combo box
            If DBView.Count > 0 Then                        'load item from the dataview row into the combo box
                For I = 0 To DBView.Count - 1
                    Dim Row As DataRow = DBView(I).Row
                    CBUserID.Items.Add(Row("userID"))
                Next
            End If
        End If

    End Sub
    '''''''''''''''''''''''''
    'reset the panel positon and visibility
    '''''''''''''''''''''''''
    Private Sub PaneLReset()
        IDS.Data.SystemAtLogin = True

        PanelIDSLogin.Top = 0
        PanelIDSLogin.Left = 0
        PanelIDSLogin.Visible = False
        Panelwelcome.Top = 0
        Panelwelcome.Left = 0
        Panelwelcome.Visible = False
        PanelChangeUserID.Top = 0
        PanelChangeUserID.Left = 0
        PanelChangeUserID.Visible = False
        PanelChangePW.Top = 0
        PanelChangePW.Left = 0
        PanelChangePW.Visible = False

        PanelSelectApplication.Top = 0
        PanelSelectApplication.Left = 0
        PanelSelectApplication.Visible = False

        PanelSSetup.Top = 0
        PanelSSetup.Left = 0
        PanelSSetup.Visible = False
        'Me.Width = PanelIDSLogin.Width
        Me.Size = PanelIDSLogin.Size

    End Sub
    'Public Ftop As New Boolean
    'Public formstup As New Form  'Xu Long 23/02/2012

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        'formstup.TopMost = True  'Xu Long 23/02/2012
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        IDS.Data.SystemAtLogin = True

        '   load data from database to ds
        '   and refresh it to the display component

        Conn = OleDbConnection1
        Call DataLoad()                                         'load data from database to DS
        IDSData = IDS.Data                                      'equate the global IDSdata to IDS.data (define at ids service module)

        Call CMGroupIDPrimaryKey_Click(sender, e)               'set as groupid as the top primary selection field

        PaneLReset()                                            'reset display panel 
        PanelIDSLogin.Visible = True
        Me.Text = "IDS Login"
        Call RefreshCBGroupID()

        '   extract all the privilegs infomation from data table
        '   and save it to the ids.data

        DBView = New DataView(DS.Tables("PrivilegeTable"))
        DBView.RowFilter = ""
        DBView.Sort = "PrivilegeID DESC"

        If DBView.Count > 0 Then
            Dim I As Integer = 0
            IDS.Data.Admin.ALLPrivileges.IDArray.Clear()
            IDS.Data.Admin.ALLPrivileges.TypeArray.Clear()

            For I = 0 To DBView.Count - 1
                Dim Row As DataRow = DBView(I).Row
                IDS.Data.Admin.ALLPrivileges.IDArray.Add(Row("PrivilegeID"))
                IDS.Data.Admin.ALLPrivileges.TypeArray.Add(Row("PrivilegeType"))
            Next I
        End If

        'start the gantry motor initialization
        'Timer1.Enabled = True
        'Timer1.Start()
        'Whole timer1 removed by Xu Long 20/02/2012
        'load the default pat file to global variable

        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()

        ' select the previous saved group id and user id as default value

        If CBGroupID.Items.Contains(IDS.Data.Admin.User.Group.ID) = True Then CBGroupID.SelectedItem = IDS.Data.Admin.User.Group.ID
        If CBUserID.Items.Contains(IDS.Data.Admin.User.ID) = True Then CBUserID.SelectedItem = IDS.Data.Admin.User.ID

        If Not LoginName = "" Then
            CBGroupID.SelectedItem = LoginName
        End If

    End Sub

    Private Sub BtnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnLogin.Click

        Dim regKey As RegistryKey
        Dim ver As Decimal

        If CBUserID.Items.Contains(CBUserID.Text) = True Then
        ElseIf CBUserID.Text = "" And CBGroupID.Text = "" Then
        Else
            MessageBox.Show("User not belongs to the Group")
            Return
        End If

        IDS.Data.MsgErr = ""
        DBView = New DataView(DS.Tables("userTable"))
        DBView.RowFilter = "UserID = '" + CBUserID.Text.ToString + "'"

        If DBView.Count = 1 Then
            Dim Row As DataRow = DBView(0).Row


            ' user login in and info to display 


            If TextLoginPW.Text = CStr(Row("UserPassword")) Then
                PaneLReset()
                Panelwelcome.Visible = True
                Me.Text = "Welcome"

                LGroupID.Text = "GroupID : " + CBGroupID.Text
                LUserID.Text = "UserID : " + CStr(Row("UserID"))
                If IsDBNull(Row("UserName")) = False Then
                    LUserName.Text = "UserName : " + CStr(Row("UserName"))
                Else
                    LUserName.Text = "UserName : "
                End If
                If IsDBNull(Row("Department")) = False Then
                    LDepartment.Text = "Department : " + CStr(Row("Department"))
                Else
                    LDepartment.Text = "Department : "
                End If
                If IsDBNull(Row("ContactNo")) = False Then
                    LContact.Text = "Contact Number: " + CStr(Row("ContactNo"))
                Else
                    LContact.Text = "Contact Number: "
                End If
            Else
                IDS.Data.MsgErr = "Wrong Password (Case Sensitive)"
                MessageBox.Show(IDS.Data.MsgErr)
            End If

        ElseIf DBView.Count = 0 Then

            '   user select to do system setup 
            '   the system setup password is stored at 
            '   registry = mycomputer/local machine/software/ids
            '   run regedit.exe to view the registry

            regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\IDS\", True)
            Dim s As String = regKey.GetValue("")
            regKey.Close()

            Try
                If CBGroupID.Text = "" And CBUserID.Text = "" And TextLoginPW.Text.ToUpper = s Then
                    ' load the default pat file to global variable
                    IDS.Data.ParameterID.RecordID = "FactoryDefault"
                    IDSData.Admin.Folder.FileExtension = "Pat"
                    IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
                    IDS.Data.OpenData()
                    ' run system setup
                    IDSData.Admin.User.RunApplication = "System"
                    IDSData.SystemAtLogin = False

                    MySetup.ShowDialog()

                    PaneLReset()
                    PanelIDSLogin.Visible = True
                    Me.Text = "IDS.data Login"
                    TextLoginPW.Text = ""
                Else
                    IDS.Data.MsgErr = "UserID not found"
                    MessageBox.Show(IDS.Data.MsgErr)
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If

    End Sub

    Private Sub CBGroupID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBGroupID.SelectedIndexChanged

        Dim I As Integer

        ''''''''''''''''''''
        ' load all user id that is under the selected group id into the user id combo box
        ''''''''''''''''''''
        If SetGroupIDPriKey = True Then
            DBView = New DataView(DS.Tables("GroupUserTable"))
            DBView.RowFilter = "GroupID = '" + CBGroupID.Text.ToString + "'"
            I = 0
            CBUserID.Items.Clear()
            CBUserID.Text = ""
            TextLoginPW.Text = "" 'Xu Long 16-02-2012 to clear password
            If DBView.Count > 0 Then
                For I = 0 To DBView.Count - 1
                    Dim Row As DataRow = DBView(I).Row
                    CBUserID.Items.Add(Row("UserID"))
                Next
            End If
        End If
    End Sub

    Private Sub CBUserID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBUserID.SelectedIndexChanged

        Dim I As Integer


        ' load all group id that is under the selected user id into the group id combo box


        If SetUserIDPriKey = True Then
            DBView = New DataView(DS.Tables("GroupUserTable"))
            DBView.RowFilter = "userID = '" + CBUserID.Text.ToString + "'"
            I = 0
            CBGroupID.Items.Clear()
            CBGroupID.Text = ""
            TextLoginPW.Text = ""  'Xu Long 16-02-2012 to clear password
            If DBView.Count > 0 Then
                For I = 0 To DBView.Count - 1
                    Dim Row As DataRow = DBView(I).Row
                    CBGroupID.Items.Add(Row("GroupID"))
                Next
            End If
        End If
    End Sub

    'not use
    Public Overrides Function PreProcessMessage(ByRef msg As System.Windows.Forms.Message) As Boolean
        Dim s As String = msg.ToString
    End Function
    'not use
    Private Sub FormLogin_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.Shift = True Then
            Dim k = e.KeyCode
        End If
    End Sub
    '''''''''''''''''''''''''
    'to reset ids system password to 1234 at the system registry 
    '''''''''''''''''''''''''
    Private Sub ButtonSetupRegistry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSetupRegistry.Click
        Dim regKey As RegistryKey
        Dim ver As Decimal
        regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
        regKey.CreateSubKey("IDS")
        regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\IDS", True)
        regKey.SetValue("", "1234")
        regKey.Close()
    End Sub
    '''''''''''''''''''''''''
    'connect to trio controller
    '''''''''''''''''''''''''
    'Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    '    Timer1.Stop()
    '    Timer1.Enabled = False

    'Removed because do not want to homing here.
    'MySetup.Connect_Controller()                'Removed by xu long.

    'End Sub
    'Whole timer1 removed by Xu Long 20/02/2012


#End Region

#Region "Welcome Panel"

    Private Sub BtnWelcomeChangeUserID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWelcomeChangeUserID.Click

        ' display welcome panel

        PaneLReset()
        PanelChangeUserID.Visible = True
        TextCurrentUSerID.Text = CBUserID.Text

    End Sub

    Private Sub BtnWelcomeChangePW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWelcomeChangePW.Click


        ' change user Password 


        PaneLReset()
        PanelChangePW.Visible = True

        DBView = New DataView(DS.Tables("UserTable"))
        DBView.RowFilter = "UserID = '" + CBUserID.Text + "'"

        If DBView.Count = 1 Then
            Dim Row As DataRow = DBView(0).Row
            TextcurrentPW.Text = CStr(Row("Userpassword"))
        End If

    End Sub

    Private Sub BtnWelcomeLogout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWelcomeLogout.Click

        ' logout from welcome panel

        PaneLReset()
        PanelIDSLogin.Visible = True
        TextLoginPW.Text = ""
    End Sub
    Private Sub WecomeOkBtn()

        ' form welcome panel proceed to application selection panel

        DBView = New DataView(DS.Tables("GroupUserTable"))
        DBView.RowFilter = "GroupID = '" + CBGroupID.Text.ToString + "' And UserID = '" + CBUserID.Text + "'"
        If DBView.Count = 1 Then
        End If
    End Sub

    '  user select applcation to run

    Public Sub BtnWelcomeOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnWelcomeOK.Click
        formlg.Hide()
        'Application.DoEvents()
        ' load global variable with factorydefault value
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()

        ' get user group is under administration
        Dim I As Integer
        Dim AdminID As String = ""
        DBView = New DataView(DS.Tables("GroupUserTable"))
        DBView.RowFilter = "GroupID = '" + CBGroupID.Text.ToString + "' And UserID = '" + CBUserID.Text + "'"
        If DBView.Count = 1 Then
            Dim Row As DataRow = DBView(0).Row
            If CBGroupID.Text.ToUpper = "ADMINISTRATION" Then
                AdminID = CStr(Row("UserID"))
            End If
        End If

        ' extract login user info
        DBView = New DataView(DS.Tables("userTable"))
        DBView.RowFilter = "UserID = '" + CBUserID.Text.ToString + "'"

        If DBView.Count = 1 Then
            Dim Row As DataRow = DBView(0).Row

            IDS.Data.Admin.User.ID = CStr(Row("UserID"))
            IDS.Data.Admin.User.Name = CStr(Row("UserName"))
            IDS.Data.Admin.User.ContactNo = CStr(row("ContactNo"))
            IDS.Data.Admin.User.Department = CStr(row("Department"))

            '   extract login group info
            DBView = New DataView(DS.Tables("GroupTable"))
            DBView.RowFilter = "GroupID = '" + CBGroupID.Text.ToString + "'"
            If DBView.Count > 0 Then
                Row = DBView(0).Row
                IDS.Data.Admin.User.Group.ID = CStr(row("GroupID"))
                IDS.Data.Admin.User.Group.Remark = CStr(row("Remark"))
            End If

            '   extract system hardware configuration info
            IDS.Data.Admin.User.Group.SystemHardwareArray.Clear()
            Dim SSView As DataView = New DataView(DS.Tables("SystemConfigureTable"))
            SSView.RowFilter = "Selected = 1"
            If SSView.Count > 0 Then
                Dim J As Integer = 0
                For J = 0 To SSView.Count - 1
                    Dim SSRow As DataRow = SSView(J).Row
                    IDS.Data.Admin.User.Group.SystemHardwareArray.Add(SSRow("Hardware"))
                Next J
            End If

            '   extract login group privilege info
            DBView = New DataView(DS.Tables("GroupPrivilegeTable"))
            DBView.RowFilter = "GroupID = '" + IDS.Data.Admin.User.Group.ID + "'"

            If DBView.Count > 0 Then
                I = 0

                IDS.Data.Admin.User.Group.PrivilegeArray.Clear()
                For I = 0 To DBView.Count - 1
                    Row = DBView(I).Row

                    SSView = New DataView(DS.Tables("SystemConfigureTable"))
                    SSView.RowFilter = "Hardware = '" + row("PrivilegeID") + "'"
                    If SSView.Count = 0 Then
                        IDS.Data.Admin.User.Group.PrivilegeArray.Add(row("PrivilegeID"))
                    Else
                        'is under configurable list          
                        Dim SSRow As DataRow = SSView(0).Row
                        If SSRow("Selected") = True Then
                            IDS.Data.Admin.User.Group.PrivilegeArray.Add(row("PrivilegeID"))
                        End If

                    End If

                Next I
            End If

            ' end of login user data exttraction
            If CBGroupID.Text.ToUpper = "ADMINISTRATION" And AdminID = CBUserID.Text Then

                '' login as administration application
                PaneLReset()
                PanelIDSLogin.Visible = True
                Me.Text = "IDS.data Login"
                TextLoginPW.Text = ""

                Dim frmAdministration As FormAdministration = New FormAdministration
                frmAdministration.ShowDialog()
                Call RefreshCBGroupID()
                CBGroupID.SelectedIndex = 0

            ElseIf CBGroupID.Text.ToUpper = "ADMINISTRATION" And AdminID = "" Then

                MessageBox.Show("User is not assign as the adminstrator")
            Else
                '' select other applications
                RadioSystem.Enabled = False
                RadioProgrammer.Enabled = False
                RadioMaintenance.Enabled = False
                RadioOperator.Enabled = False

                RadioSystem.Checked = False
                RadioProgrammer.Checked = False
                RadioMaintenance.Checked = False
                RadioOperator.Checked = False

                RoleCounter = 0

                '' login as system setup (notused)
                If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("System") = True Then
                    RadioSystem.Enabled = True
                    RoleCounter += 1
                End If
                '' login as programmer
                If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Programmer") = True Then
                    RadioProgrammer.Enabled = True
                    RoleCounter += 1
                End If
                '' login as maintenance
                If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Maintenance") = True Then
                    RadioMaintenance.Enabled = True
                    RoleCounter += 1
                End If
                '' login as operator
                If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Operator") = True Then
                    RadioOperator.Enabled = True
                    RoleCounter += 1
                End If

                '' login as multiple applications; programmer, maintenance ,operator, system setup(notused)
                If RoleCounter > 1 Then
                    PaneLReset()
                    PanelSelectApplication.Visible = True
                    Me.Text = "IDS Select Application"

                Else

                    If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("System") = True Then
                        IDS.Data.Admin.User.RunApplication = "System"
                        Call LoadSSetup()
                    ElseIf IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Programmer") = True Then
                        IDS.Data.Admin.User.RunApplication = "Programmer"
                        Call LoadProgrammerMaintenance()
                    ElseIf IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Maintenance") = True Then
                        IDS.Data.Admin.User.RunApplication = "Maintenace"
                        Call LoadProgrammerMaintenance()
                    ElseIf IDS.Data.Admin.User.Group.PrivilegeArray.Contains("Operator") = True Then
                        IDS.Data.Admin.User.RunApplication = "Operator"
                        Call LoadProgrammerMaintenance()
                    End If

                End If
            End If
        End If

    End Sub


#End Region

#Region "Panel Application Selection"
    Dim RoleCounter As Integer

    '''''''''''''''''''''''''
    '' from application panel go backward to welcome panel
    '''''''''''''''''''''''''
    Private Sub BtnSelectapplcationBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSelectapplcationBack.Click

        TextConfirmPW.Text = ""
        TextNewPW.Text = ""
        PaneLReset()
        Panelwelcome.Visible = True
        Me.Text = "Welcome"

    End Sub

    '''''''''''''''''''''''''
    '' select application
    '''''''''''''''''''''''''
    Private Sub BtnSelectApplicationOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSelectApplicationOK.Click


        '''''''''''''''''''
        ' global variable get default value from the default pat file
        '''''''''''''''''''

        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()


        ''''''''''''''''''''''''''''''
        '' login as multiple applications; programmer, maintenance ,operator, system setup(notused)
        '''''''''''''''''''''''''''''''

        If RadioSystem.Checked = True Then
            IDS.Data.Admin.User.RunApplication = "System"
            Call LoadSSetup()
        ElseIf RadioProgrammer.Checked = True Then
            IDS.Data.Admin.User.RunApplication = "Programmer"
            Call LoadProgrammerMaintenance()
        ElseIf RadioMaintenance.Checked = True Then
            IDS.Data.Admin.User.RunApplication = "Maintenance"
            Call LoadProgrammerMaintenance()
        ElseIf RadioOperator.Checked = True Then
            IDS.Data.Admin.User.RunApplication = "Operator"
            Call LoadProgrammerMaintenance()
        End If


    End Sub
    ''''''''''''''''''''''''''''''
    '' load system setup panel - not used
    '''''''''''''''''''''''''''''''
    Private Sub LoadSSetup()
        PaneLReset()
        PanelSSetup.Visible = True
        Me.Text = "System Configuration Authentication"
        TextSSetupGroupID.Text = IDS.Data.Admin.User.Group.ID
        TextSSetupUserID.Text = IDS.Data.Admin.User.ID
        TextSSetupPassword.Text = ""
    End Sub
    ''''''''''''''''''''''''''''''
    '' create new instance for the process setup forms and run applications
    '''''''''''''''''''''''''''''''
    Public Sub LoadProgrammerMaintenance()

        MyConveyorSettings = New ConveyorSettings
        MyEventSettings = New EventSettings
        MySPCLogging = New SPCLogging
        MyGantrySettings = New GantrySettings
        MyHeaterSettings = New HeaterSettings
        MyVolumeCalibrationSettings = New VolumeCalibrationSettings
        MySettings = New Settings

        '  for process setup 
        '  to control the visibility of the radio button 
        '  with repects to user's login group's hardware setup list
        If IDS.Data.Admin.User.Group.SystemHardwareArray.Contains("ThermalController") = True Then
            MySettings.buttonthermalsettings.Visible = True
        Else
            MySettings.buttonthermalsettings.Visible = False
        End If
        If IDS.Data.Admin.User.Group.SystemHardwareArray.Contains("VolumeCalibration") = True Then
            MySettings.ButtonVolumeCalibSettings.Visible = True
        Else
            MySettings.ButtonVolumeCalibSettings.Visible = False
        End If

        ' to enable the radio button with repects to the login group's privilege
        If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("ThermalController") = True Then
            MySettings.buttonthermalsettings.Enabled = True
        Else
            MySettings.buttonthermalsettings.Enabled = False
        End If

        If IDS.Data.Admin.User.Group.PrivilegeArray.Contains("VolumeCalibration") = True Then
            MySettings.ButtonVolumeCalibSettings.Enabled = True
        Else
            MySettings.ButtonVolumeCalibSettings.Enabled = False
        End If

        ' create new instance for programming and production forms
        Production = New DLL_Execution.FormProduction
        Programming = New DLL_Execution.FormProgramming

        IDSData.SystemAtLogin = False
        If IDS.Data.Admin.User.RunApplication = "Programmer" Then
            Programming.CurrentMode = "Program Editor"
            Programming.ShowDialog()
        ElseIf IDS.Data.Admin.User.RunApplication = "Operator" Then
            Production.ShowDialog()
        End If

        IDSData.SystemAtLogin = True

        ' no application is selected for the login group id - go back to login panel
        If RoleCounter <= 1 Then
            PaneLReset()
            PanelIDSLogin.Visible = True
            Me.Text = "IDS.data Login"
            TextLoginPW.Text = ""
        End If

    End Sub

#End Region

#Region "System configuration"
    '''''''''''''''''''''''''''''''
    ' login to system setup application - not used
    '''''''''''''''''''''''''''''''
    Private Sub BtnSSetupOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSSetupOK.Click

        If TextSSetupPassword.Text.ToLower = "a" Then
            MySetup.ShowDialog()
            PaneLReset()
            PanelIDSLogin.Visible = True
            Me.Text = "IDS.data Login"
            TextLoginPW.Text = ""
        Else
            MessageBox.Show("Wrong Password")
        End If
    End Sub
    '''''''''''''''''''''''''''''''
    ' from system setup backward to welcome panel - notused
    '''''''''''''''''''''''''''''''
    Private Sub BtnSSetupBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSSetupBack.Click

        If RoleCounter > 1 Then
            PaneLReset()
            PanelSelectApplication.Visible = True
            Me.Text = "IDS Select Application"
        Else
            PaneLReset()
            Panelwelcome.Visible = True
            Me.Text = "Welcome"
        End If
    End Sub

#End Region

#Region "Change Password"
    '''''''''''''''''''''''''''''''
    ' changer user password
    '''''''''''''''''''''''''''''''
    Private Sub btnchangePW_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnchangePW.Click


        '''''''''''''''''''''''''''''''
        ' validate new password 
        '''''''''''''''''''''''''''''''
        IDS.Data.MsgErr = ""
        If TextcurrentPW.Text = "" Then
            IDS.Data.MsgErr = "Current Password cant be empty"
            MessageBox.Show(IDS.Data.MsgErr)
            Return

        ElseIf TextNewPW.Text <> TextConfirmPW.Text Then
            IDS.Data.MsgErr = "New and confirm PassWord not the same"
            MessageBox.Show(IDS.Data.MsgErr)
            TextConfirmPW.Text = ""
            TextNewPW.Text = ""
            Return

        ElseIf TextNewPW.Text = "" Then
            IDS.Data.MsgErr = "New Password cant be empty"
            MessageBox.Show(IDS.Data.MsgErr)
            Return

        End If

        '''''''''''''''''''''''''''''''
        'update user password into the database system
        '''''''''''''''''''''''''''''''
        DBView = New DataView(DS.Tables("UserTable"))
        DBView.RowFilter = "UserID = '" + CBUserID.Text + "'"
        If DBView.Count = 0 Then
            IDS.Data.MsgErr = "UserID not found"
            MessageBox.Show(IDS.Data.MsgErr)
        Else
            Dim Row As DataRow = DBView(0).Row

            If CStr(Row("Userpassword")).ToUpper = TextcurrentPW.Text.ToUpper Then
                Row("Userpassword") = TextNewPW.Text
                UpdateData("SELECT * FROM UserTable", "UserTable")

                If IDS.Data.MsgErr = "" Then
                    MessageBox.Show("Operation Successful")
                    BtnChangePWBack_Click(sender, e)
                End If
            Else
                IDS.Data.MsgErr = "Wrong Password"
                MessageBox.Show(IDS.Data.MsgErr)
                TextNewPW.Text = ""
                TextConfirmPW.Text = ""
            End If
        End If

    End Sub
    '''''''''''''''''''''''''''''''
    'from change password panel backward to welcome panel
    '''''''''''''''''''''''''''''''
    Private Sub BtnChangePWBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnChangePWBack.Click
        TextConfirmPW.Text = ""
        TextNewPW.Text = ""
        PaneLReset()
        Panelwelcome.Visible = True
        Me.Text = "Welcome"
    End Sub
#End Region

#Region "Change UserID"
    '''''''''''''''''''''''''''''''
    'from changer user id panel backwards to welcome panel
    '''''''''''''''''''''''''''''''
    Private Sub BtnChangeUserIDBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnChangeUserIDBack.Click
        TextNewUserID.Text = ""
        TextConfirmUserID.Text = ""
        PaneLReset()
        Panelwelcome.Visible = True
        Me.Text = "Welcome"
    End Sub
    '''''''''''''''''''''''''''''''
    'changer user id 
    '''''''''''''''''''''''''''''''
    Private Sub BtnchangeUserID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnchangeUserID.Click
        Dim I As Integer

        '''''''''''''''''''''''''''''''
        'validate new userid 
        '''''''''''''''''''''''''''''''
        IDS.Data.MsgErr = ""
        If TextCurrentUSerID.Text = "" Then
            IDS.Data.MsgErr = "Current UserID cant be empty"
            MessageBox.Show(IDS.Data.MsgErr)
            Return

        ElseIf TextNewUserID.Text <> TextConfirmUserID.Text Then
            IDS.Data.MsgErr = "New and confirm UserID are not the same"
            MessageBox.Show(IDS.Data.MsgErr)
            TextConfirmUserID.Text = ""
            TextNewUserID.Text = ""
            Return

        ElseIf TextNewUserID.Text = "" Then
            IDS.Data.MsgErr = "New UserId cant be empty"
            MessageBox.Show(IDS.Data.MsgErr)
            TextConfirmUserID.Text = ""
            TextNewUserID.Text = ""
            Return
        End If



        '''''''''''''''''''''''''''''''
        ' update new user id with info into the database system
        '''''''''''''''''''''''''''''''
        Try

            DBView = New DataView(DS.Tables("UserTable"))
            DBView.RowFilter = "UserID = '" + TextConfirmUserID.Text + "'"
            If DBView.Count > 0 Then
                'new user id already exist
                IDS.Data.MsgErr = "There is already a UserID = " + TextConfirmUserID.Text + ". Operation not allowed"
                MessageBox.Show(IDS.Data.MsgErr)
                TextConfirmUserID.Text = ""
                TextNewUserID.Text = ""
                Return
            Else
            End If


            'Get the list of Groups of the current user involved
            Dim UserInGroupID As New ArrayList
            DBView = New DataView(DS.Tables("GroupUserTable"))
            DBView.RowFilter = "UserID = '" + TextCurrentUSerID.Text + "'"
            If DBView.Count > 0 Then
                I = 0
                For I = 0 To DBView.Count - 1
                    Dim Row As DataRow = DBView(I).Row
                    UserInGroupID.Add(row("GroupID"))
                Next
            End If

            'Work on the User Table - add new user id and delele the current user id
            DBView = New DataView(DS.Tables("UserTable"))
            DBView.RowFilter = "UserID = '" + TextCurrentUSerID.Text + "'"


            'delete current user id info from user table
            IDS.Data.MsgErr = ""
            If DBView.Count = 0 Then
                IDS.Data.MsgErr = "UserID not found"
                MessageBox.Show(IDS.Data.MsgErr)
                Return
            Else
                Dim Row As DataRow = DBView(0).Row
                Dim NewRow As DataRow = DBView.Table.NewRow()
                NewRow("UserID") = TextConfirmUserID.Text
                NewRow("UserName") = Row("UserName")
                NewRow("Department") = Row("Department")
                NewRow("ContactNo") = Row("ContactNo")
                NewRow("UserPassword") = Row("UserPassword")
                DBView.Table.Rows.Add(NewRow)
                row.Delete()
                UpdateData("SELECT * FROM UserTable", "UserTable")
            End If

            'add new user info and group ifo into the groupuser table
            If UserInGroupID.Count > 0 Then
                I = 0
                DBView = New DataView(DS.Tables("GroupUserTable"))
                For I = 0 To UserInGroupID.Count - 1
                    Dim NewRow As DataRow = DBView.Table.NewRow()

                    newrow("GroupID") = CStr(UserInGroupID.Item(I))
                    Newrow("UserID") = TextConfirmUserID.Text
                    DBView.Table.Rows.Add(NewRow)
                Next
                UpdateData("SELECT * FROM GroupUserTable", "GroupUserTable")
            End If

            IDS.Data.MsgErr = ""
        Catch ex As Exception
            IDS.Data.MsgErr = ex.ToString
            MessageBox.Show(ex.ToString, "Error!", MessageBoxButtons.OK)
        End Try



        ' end of rename user id - operatio successful
        If IDS.Data.MsgErr = "" Then
            MessageBox.Show("Operation Successful")
            Call RefreshCBGroupID()
            CBGroupID_SelectedIndexChanged(sender, e)
            CBUserID.Text = TextConfirmUserID.Text
        End If
        TextNewUserID.Text = ""
        TextConfirmUserID.Text = ""
        BtnLogin_Click(sender, e)
    End Sub
#End Region

End Class
