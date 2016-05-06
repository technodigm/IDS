Imports DLL_DataManager
'Imports DLL_Export_Devices
'Imports DLL_Export_Device_Xu
Imports DLL_Export_Service
'Imports DLL_Export_Device_Motor
'Imports DLL_Export_Device_Vision
'Imports System.Data.OleDb
Imports System.IO
Imports System.Threading.Thread

Public Class SystemIO
    Inherits System.Windows.Forms.Form
    Dim EArg As System.EventArgs

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Call formatcol()

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
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label As System.Windows.Forms.Label
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TB_CableLabel As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TB_IONum As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents RD_OFF As System.Windows.Forms.RadioButton
    Friend WithEvents RD_On As System.Windows.Forms.RadioButton
    Friend WithEvents CB_Type As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TB_ModuleName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TB_IOName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Btn_Cancel As System.Windows.Forms.Button
    Friend WithEvents Btn_Apply As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Btn_exit As System.Windows.Forms.Button
    Friend WithEvents TBr_SysIO As System.Windows.Forms.ToolBar
    Friend WithEvents ToolBarButtonNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonEdit As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonDelete As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonImport As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonExport As System.Windows.Forms.ToolBarButton
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CBViewType As System.Windows.Forms.ComboBox
    Friend WithEvents CBViewModule As System.Windows.Forms.ComboBox
    Friend WithEvents TimerIO As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SystemIO))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TBr_SysIO = New System.Windows.Forms.ToolBar
        Me.ToolBarButtonNew = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonEdit = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonDelete = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonImport = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonExport = New System.Windows.Forms.ToolBarButton
        Me.TB_CableLabel = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TB_IONum = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.RD_OFF = New System.Windows.Forms.RadioButton
        Me.RD_On = New System.Windows.Forms.RadioButton
        Me.CB_Type = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.TB_ModuleName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TB_IOName = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.CBViewType = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.CBViewModule = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Btn_Cancel = New System.Windows.Forms.Button
        Me.Btn_Apply = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label = New System.Windows.Forms.Label
        Me.Btn_exit = New System.Windows.Forms.Button
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.TimerIO = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Btn_Cancel)
        Me.Panel1.Controls.Add(Me.Btn_Apply)
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.Label)
        Me.Panel1.Controls.Add(Me.Btn_exit)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1280, 944)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.DataGrid1)
        Me.Panel2.Controls.Add(Me.GroupBox1)
        Me.Panel2.Location = New System.Drawing.Point(8, 48)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1144, 816)
        Me.Panel2.TabIndex = 69
        '
        'DataGrid1
        '
        Me.DataGrid1.CaptionFont = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 248)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(1112, 552)
        Me.DataGrid1.TabIndex = 56
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TBr_SysIO)
        Me.GroupBox1.Controls.Add(Me.TB_CableLabel)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.TB_IONum)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.RD_OFF)
        Me.GroupBox1.Controls.Add(Me.RD_On)
        Me.GroupBox1.Controls.Add(Me.CB_Type)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.TB_ModuleName)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.TB_IOName)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1080, 208)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "System Information Entry Tools"
        '
        'TBr_SysIO
        '
        Me.TBr_SysIO.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TBr_SysIO.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButtonNew, Me.ToolBarButtonEdit, Me.ToolBarButtonDelete, Me.ToolBarButtonImport, Me.ToolBarButtonExport})
        Me.TBr_SysIO.ButtonSize = New System.Drawing.Size(96, 50)
        Me.TBr_SysIO.Dock = System.Windows.Forms.DockStyle.None
        Me.TBr_SysIO.DropDownArrows = True
        Me.TBr_SysIO.ImageList = Me.ImageList1
        Me.TBr_SysIO.Location = New System.Drawing.Point(760, 48)
        Me.TBr_SysIO.Name = "TBr_SysIO"
        Me.TBr_SysIO.ShowToolTips = True
        Me.TBr_SysIO.Size = New System.Drawing.Size(288, 56)
        Me.TBr_SysIO.TabIndex = 68
        '
        'ToolBarButtonNew
        '
        Me.ToolBarButtonNew.ImageIndex = 0
        Me.ToolBarButtonNew.Text = "AddRow"
        Me.ToolBarButtonNew.ToolTipText = "New"
        '
        'ToolBarButtonEdit
        '
        Me.ToolBarButtonEdit.ImageIndex = 1
        Me.ToolBarButtonEdit.Text = "UpdateRow"
        Me.ToolBarButtonEdit.ToolTipText = "Update"
        '
        'ToolBarButtonDelete
        '
        Me.ToolBarButtonDelete.ImageIndex = 2
        Me.ToolBarButtonDelete.Text = "DeleteRow"
        Me.ToolBarButtonDelete.ToolTipText = "Delete"
        '
        'ToolBarButtonImport
        '
        Me.ToolBarButtonImport.Text = "Import"
        Me.ToolBarButtonImport.ToolTipText = "Import"
        Me.ToolBarButtonImport.Visible = False
        '
        'ToolBarButtonExport
        '
        Me.ToolBarButtonExport.Text = "Export"
        Me.ToolBarButtonExport.ToolTipText = "Export"
        Me.ToolBarButtonExport.Visible = False
        '
        'TB_CableLabel
        '
        Me.TB_CableLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_CableLabel.Location = New System.Drawing.Point(848, 160)
        Me.TB_CableLabel.Name = "TB_CableLabel"
        Me.TB_CableLabel.Size = New System.Drawing.Size(184, 27)
        Me.TB_CableLabel.TabIndex = 67
        Me.TB_CableLabel.Text = ""
        '
        'Label5
        '
        Me.Label5.Enabled = False
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(744, 160)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(104, 23)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "Cable Label"
        '
        'TB_IONum
        '
        Me.TB_IONum.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_IONum.Location = New System.Drawing.Point(848, 120)
        Me.TB_IONum.Name = "TB_IONum"
        Me.TB_IONum.Size = New System.Drawing.Size(184, 27)
        Me.TB_IONum.TabIndex = 65
        Me.TB_IONum.Text = ""
        '
        'Label6
        '
        Me.Label6.Enabled = False
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(744, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 23)
        Me.Label6.TabIndex = 64
        Me.Label6.Text = "IO Number"
        '
        'RD_OFF
        '
        Me.RD_OFF.Enabled = False
        Me.RD_OFF.ForeColor = System.Drawing.Color.MidnightBlue
        Me.RD_OFF.Location = New System.Drawing.Point(632, 168)
        Me.RD_OFF.Name = "RD_OFF"
        Me.RD_OFF.Size = New System.Drawing.Size(64, 24)
        Me.RD_OFF.TabIndex = 63
        Me.RD_OFF.Text = "OFF"
        '
        'RD_On
        '
        Me.RD_On.Enabled = False
        Me.RD_On.ForeColor = System.Drawing.SystemColors.Highlight
        Me.RD_On.Location = New System.Drawing.Point(560, 168)
        Me.RD_On.Name = "RD_On"
        Me.RD_On.Size = New System.Drawing.Size(56, 24)
        Me.RD_On.TabIndex = 62
        Me.RD_On.Text = "ON"
        '
        'CB_Type
        '
        Me.CB_Type.Items.AddRange(New Object() {"Input", "Output", "Voltage", "N.A"})
        Me.CB_Type.Location = New System.Drawing.Point(560, 128)
        Me.CB_Type.Name = "CB_Type"
        Me.CB_Type.Size = New System.Drawing.Size(136, 31)
        Me.CB_Type.TabIndex = 61
        Me.CB_Type.Text = "Input"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(480, 168)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 23)
        Me.Label4.TabIndex = 60
        Me.Label4.Text = "Status"
        '
        'Label3
        '
        Me.Label3.Enabled = False
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(480, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 23)
        Me.Label3.TabIndex = 59
        Me.Label3.Text = "Type"
        '
        'TB_ModuleName
        '
        Me.TB_ModuleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_ModuleName.Location = New System.Drawing.Point(144, 168)
        Me.TB_ModuleName.Name = "TB_ModuleName"
        Me.TB_ModuleName.Size = New System.Drawing.Size(320, 27)
        Me.TB_ModuleName.TabIndex = 58
        Me.TB_ModuleName.Text = ""
        '
        'Label2
        '
        Me.Label2.Enabled = False
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(16, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 23)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Module Name"
        '
        'TB_IOName
        '
        Me.TB_IOName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_IOName.Location = New System.Drawing.Point(144, 128)
        Me.TB_IOName.Name = "TB_IOName"
        Me.TB_IOName.Size = New System.Drawing.Size(320, 27)
        Me.TB_IOName.TabIndex = 54
        Me.TB_IOName.Text = ""
        '
        'Label1
        '
        Me.Label1.Enabled = False
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(16, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 23)
        Me.Label1.TabIndex = 56
        Me.Label1.Text = "IO Name"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CBViewType)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.CBViewModule)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(696, 72)
        Me.GroupBox2.TabIndex = 58
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "View By"
        '
        'CBViewType
        '
        Me.CBViewType.Items.AddRange(New Object() {"ALL", "Input", "Output", "Voltage", "N.A"})
        Me.CBViewType.Location = New System.Drawing.Point(552, 32)
        Me.CBViewType.Name = "CBViewType"
        Me.CBViewType.Size = New System.Drawing.Size(128, 31)
        Me.CBViewType.TabIndex = 63
        Me.CBViewType.Text = "ALL"
        '
        'Label10
        '
        Me.Label10.Enabled = False
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(472, 32)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(48, 23)
        Me.Label10.TabIndex = 62
        Me.Label10.Text = "Type"
        '
        'Label7
        '
        Me.Label7.Enabled = False
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(16, 32)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(120, 23)
        Me.Label7.TabIndex = 54
        Me.Label7.Text = "Module Name"
        '
        'CBViewModule
        '
        Me.CBViewModule.Items.AddRange(New Object() {"ALL"})
        Me.CBViewModule.Location = New System.Drawing.Point(136, 32)
        Me.CBViewModule.Name = "CBViewModule"
        Me.CBViewModule.Size = New System.Drawing.Size(328, 31)
        Me.CBViewModule.TabIndex = 36
        Me.CBViewModule.Text = "ALL"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(152, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(832, 40)
        Me.Label9.TabIndex = 61
        Me.Label9.Text = "Note to developer: please prevent track ball from moving the robot when this page" & _
        " is showing, because the vision window is un-visible here."
        Me.Label9.Visible = False
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label8.Location = New System.Drawing.Point(0, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(192, 32)
        Me.Label8.TabIndex = 58
        Me.Label8.Text = "System IO"
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(1040, 864)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(96, 50)
        Me.Btn_Cancel.TabIndex = 57
        Me.Btn_Cancel.Text = "Cancel"
        '
        'Btn_Apply
        '
        Me.Btn_Apply.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Btn_Apply.Location = New System.Drawing.Point(912, 864)
        Me.Btn_Apply.Name = "Btn_Apply"
        Me.Btn_Apply.Size = New System.Drawing.Size(100, 50)
        Me.Btn_Apply.TabIndex = 56
        Me.Btn_Apply.Text = "Apply"
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.Location = New System.Drawing.Point(48, -41)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 28)
        Me.ComboBox1.TabIndex = 34
        Me.ComboBox1.Text = "All System IO"
        '
        'Label
        '
        Me.Label.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label.Location = New System.Drawing.Point(0, -41)
        Me.Label.Name = "Label"
        Me.Label.Size = New System.Drawing.Size(48, 41)
        Me.Label.TabIndex = 28
        Me.Label.Text = "IO"
        '
        'Btn_exit
        '
        Me.Btn_exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Btn_exit.Image = CType(resources.GetObject("Btn_exit.Image"), System.Drawing.Image)
        Me.Btn_exit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Btn_exit.Location = New System.Drawing.Point(1184, 16)
        Me.Btn_exit.Name = "Btn_exit"
        Me.Btn_exit.Size = New System.Drawing.Size(75, 50)
        Me.Btn_exit.TabIndex = 68
        Me.Btn_exit.TabStop = False
        Me.Btn_exit.Text = "Exit"
        Me.Btn_exit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
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
        'TimerIO
        '
        Me.TimerIO.Interval = 10
        '
        'SystemIO
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1280, 944)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "SystemIO"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "SystemIO"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Version 6/07/2007 - Lsgoh "

    Dim DBview As New DataView
    Dim TableSystemIO As New DataTable
    Dim Para As New DLL_DataManager.CIDSData.CIDSParameterID
    Dim dgtablestyle As New DataGridTableStyle
    Dim IONameCol As New DataGridTextBoxColumn
    Dim ModuleNameCol As New DataGridTextBoxColumn
    Dim TypeCol As New DataGridTextBoxColumn
    Dim StatusCol As New DataGridTextBoxColumn
    Dim IONumCol As New DataGridTextBoxColumn
    Dim CableNameCol As New DataGridTextBoxColumn
    Dim Itemcol As New DataGridTextBoxColumn

    Private Sub formatcol()

        dgtablestyle.MappingName = "TableSystemIO"

        IONameCol.MappingName = "IOName"
        IONameCol.HeaderText = "IO Name"
        IONameCol.Width = 300
        IONameCol.Alignment = HorizontalAlignment.Left
        IONameCol.ReadOnly = True

        TypeCol.MappingName = "Type"
        TypeCol.HeaderText = "Type"
        TypeCol.Width = 100
        TypeCol.Alignment = HorizontalAlignment.Center
        TypeCol.ReadOnly = True

        StatusCol.MappingName = "Status"
        StatusCol.HeaderText = "Status"
        StatusCol.Width = 50
        StatusCol.Alignment = HorizontalAlignment.Center
        ' StatusCol.ReadOnly = True

        IONumCol.MappingName = "IO"
        IONumCol.HeaderText = "IO Number"
        IONumCol.Width = 100
        IONumCol.Alignment = HorizontalAlignment.Center
        IONumCol.ReadOnly = True

        CableNameCol.MappingName = "CableName"
        CableNameCol.HeaderText = "CableName"
        CableNameCol.Width = 100
        CableNameCol.Alignment = HorizontalAlignment.Center
        CableNameCol.ReadOnly = True

        ModuleNameCol.MappingName = "ModuleName"
        ModuleNameCol.HeaderText = "Module" + vbNewLine + "Name"
        ModuleNameCol.Width = 250
        ModuleNameCol.Alignment = HorizontalAlignment.Center
        ModuleNameCol.ReadOnly = True

        Itemcol.MappingName = "Item"
        Itemcol.HeaderText = "Item"
        Itemcol.Width = 40
        Itemcol.Alignment = HorizontalAlignment.Left
        Itemcol.ReadOnly = True

        dgtablestyle.GridColumnStyles.Add(Itemcol)
        dgtablestyle.GridColumnStyles.Add(IONameCol)
        dgtablestyle.GridColumnStyles.Add(TypeCol)
        dgtablestyle.GridColumnStyles.Add(StatusCol)
        dgtablestyle.GridColumnStyles.Add(IONumCol)
        dgtablestyle.GridColumnStyles.Add(CableNameCol)
        dgtablestyle.GridColumnStyles.Add(ModuleNameCol)


        DataGrid1.TableStyles.Add(dgtablestyle)
    End Sub
    Private Sub CBViewModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBViewModule.SelectedIndexChanged
        RefreshDisplay()
    End Sub
    Private Sub CBViewType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBViewType.SelectedIndexChanged
        RefreshDisplay()
    End Sub
    Private Sub RefreshDisplay()


        IOReadStop = True
        Sleep(100)

        'for refreshing purpose
        IONameCol.ReadOnly = False
        TypeCol.ReadOnly = False
        'StatusCol.ReadOnly = False
        IONumCol.ReadOnly = False
        CableNameCol.ReadOnly = False
        ModuleNameCol.ReadOnly = False
        Itemcol.ReadOnly = False



        If CBViewModule.Text = "ALL" And CBViewType.Text = "ALL" Then
            FilterString = ""

        ElseIf CBViewModule.Text <> "ALL" And CBViewType.Text = "ALL" Then
            FilterString = "ModuleName = '" + CBViewModule.Text + "'"

        ElseIf CBViewModule.Text = "ALL" And CBViewType.Text <> "ALL" Then
            FilterString = "Type = '" + CBViewType.Text + "'"

        ElseIf CBViewModule.Text <> "ALL" And CBViewType.Text <> "ALL" Then
            FilterString = "ModuleName = '" + CBViewModule.Text + "' AND Type = '" + CBViewType.Text + "'"
        End If

        '''for refreshing purpose
        Dim _datatable As New DataTable
        DBview = New DataView(_datatable)
        DataGrid1.DataSource = DBview
        ''''''''''''''''''''''''''''''''

        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = FilterString
        DataGrid1.DataSource = DBview


        If DBview.Count = 0 Then
            'Ignore if there is no data rows in the table
        Else
            DBview.Sort = "CableName Asc"


            'Check if each module name is listed in the combobox for selection
            Dim moduleItem As Object
            For I As Integer = 0 To DBview.Count - 1
                Dim Row As DataRow = DBview(I).Row
                Row("Item") = I + 1

                ''''''''''''''''''
                moduleItem = Row(6)
                'moduleItem = DataGrid1.Item(I, 5)
                If CBViewModule.Items.Contains(moduleItem) Then
                Else
                    CBViewModule.Items.Add(CString(moduleItem))
                End If
            Next
        End If


        'for refreshing purpose
        IONameCol.ReadOnly = True
        TypeCol.ReadOnly = True
        'StatusCol.ReadOnly = True
        IONumCol.ReadOnly = True
        CableNameCol.ReadOnly = True
        ModuleNameCol.ReadOnly = True
        Itemcol.ReadOnly = True


        Call DataGrid1_Click(Me, EArg)

        IOReadStop = False

    End Sub
    Friend Sub SetIOTable()

        If IDS.Data.Admin.User.RunApplication = "System" Then

            Btn_Apply.Visible = True
            Btn_Cancel.Visible = True
            TBr_SysIO.Visible = True
        Else

            Btn_Apply.Visible = False
            Btn_Cancel.Visible = False
            TBr_SysIO.Visible = False
        End If

        CBViewModule.Items.Clear()
        CBViewModule.Items.Add("ALL")
        CBViewModule.Text = "ALL"

        Call FiletoTable()
        DS.Tables("TableSystemIO").Columns("ID").ColumnMapping = MappingType.Hidden
        DBview = New DataView(DS.Tables("TableSystemIO"))


        DataGrid1.DataSource() = DBview.DataViewManager
        DataGrid1.CaptionText = DataGrid1.DataMember


        Call RefreshDisplay()
    End Sub
    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
        DataGrid1_Click(sender, e)
    End Sub
    Private Sub DataGrid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGrid1.KeyUp
        If e.KeyCode = Keys.Delete Then
            Call RefreshDisplay()
        End If
    End Sub
    Private Sub DataGrid1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.Click
        'Return

        Dim Status As String
        Dim CurrentItem As Object
        Dim CurrentRowIndex As Integer = DataGrid1.CurrentCell.RowNumber()


        Dim i As Integer = 0
        While (True)

            Try
                If IsDBNull(DataGrid1.Item(CurrentRowIndex, i)) Then
                    Return
                End If
            Catch ex As Exception
                Return
            End Try

            CurrentItem = DataGrid1.Item(CurrentRowIndex, i)

            If i = 1 Then
                TB_IOName.Text = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.IOName = TB_IOName.Text
            End If

            If i = 2 Then
                CB_Type.Text = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.Type = CB_Type.Text
            End If

            If i = 3 Then
                Status = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.Status = Status
            End If

            If i = 4 Then
                TB_IONum.Text = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.IO = TB_IONum.Text
            End If

            If i = 5 Then
                TB_CableLabel.Text = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.CableName = TB_CableLabel.Text
            End If

            If i = 6 Then
                TB_ModuleName.Text = CStr(CurrentItem)
                IDSData.Hardware.SystemIO.Template.ModuleName = TB_ModuleName.Text
                Exit While
            End If

            i = i + 1
        End While



        If CB_Type.Text.ToUpper = "OUTPUT" Then

            RD_OFF.Enabled = True
            RD_On.Enabled = True

            If Status = "ON" Then
                RD_On.Checked = True
            ElseIf Status = "OFF" Then
                RD_OFF.Checked = True
            End If

        Else
            RD_OFF.Enabled = False
            RD_On.Enabled = False
        End If


    End Sub
    Private Sub TBr_SysIO_ButtonClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBr_SysIO.ButtonClick

        IOReadStop = True
        Sleep(100)

        Dim Card As String
        Dim IOPort As String
        Dim IOBit As String
        Dim Message As String
        Dim BitFlag As String



        DataGrid1.Refresh()
        Select Case e.Button.Text
            Case "AddRow"                'Add NewRow

                'If ValidateIONumber(TB_IONum.Text, Card, IOPort, IOBit, Message) = False Then
                'MessageBox.Show(Message)
                'Exit Select
                'End If

                Dim Row As DataRow = DBview.Table.Rows.Find(TB_CableLabel.Text)
                If row Is Nothing Then

                    Row = DBview.Table.NewRow()
                    Row("ID") = "FactoryDefault"
                    Row("IOName") = TB_IOName.Text
                    Row("ModuleName") = TB_ModuleName.Text
                    Row("IO") = TB_IONum.Text
                    Row("CableName") = TB_CableLabel.Text
                    If RD_On.Checked Then
                        ' Row("Status") = "ON"
                    ElseIf RD_OFF.Checked Then
                        ' Row("Status") = "OFF"
                    Else
                        'If CB_Type.Text = "Input" Then
                        'Row("Status") = "Reading..."
                        'Else
                        '    Row("Status") = ""
                        'End If
                    End If
                    Row("Status") = " "
                    Row("Type") = CB_Type.Text
                    DBview.Table.Rows.Add(Row)

                    'Add in new module added to the combobox for fitering later
                    If CBViewModule.Items.Contains(TB_ModuleName.Text) Then
                    Else
                        CBViewModule.Items.Add(TB_ModuleName.Text)
                        CBViewModule.SelectedItem = TB_ModuleName.Text
                        RefreshDisplay()

                    End If
                Else
                    MessageBox.Show("Cable Name: " + TB_CableLabel.Text + " already exists.")
                    Exit Select
                End If


            Case "UpdateRow"            'EditRow


                'If ValidateIONumber(TB_IONum.Text, Card, IOPort, IOBit, Message) = False Then
                'MessageBox.Show(Message)
                'Exit Select
                'End If


                Dim Row As DataRow = DBview.Table.Rows.Find(TB_CableLabel.Text)

                If row Is Nothing Then
                    MessageBox.Show("Please use Add button.")
                    Exit Select
                Else
                    Row("IOName") = TB_IOName.Text
                    Row("ModuleName") = TB_ModuleName.Text
                    Row("IO") = TB_IONum.Text
                    Row("CableName") = TB_CableLabel.Text
                    Row("Type") = CB_Type.Text

                    If RD_On.Checked Then
                        'row("Status") = "ON"
                    ElseIf RD_OFF.Checked Then
                        'Row("Status") = "OFF"
                    Else
                        'If CB_Type.Text = "Input" Then
                        'Row("Status") = "Reading..."
                        'Else
                        '    Row("Status") = ""
                        'End If
                    End If
                    Row("Status") = " "
                    'Add in new module added to the combobox for fitering later
                    If CBViewModule.Items.Contains(TB_ModuleName.Text) Then
                    Else
                        CBViewModule.Items.Add(TB_ModuleName.Text)
                    End If
                End If


            Case "DeleteRow"            'DeleteRow
                Dim moduleItem As Object

                If DBview.Count = 0 Then
                Else
                    If DBview.Count = 1 Then
                        'When delete the only row left, remove the module name from combobox
                        moduleItem = DataGrid1.Item(DBview.Count - 1, 5)
                        If CBViewModule.Items.Contains(moduleItem) Then
                            CBViewModule.Items.Remove(CString(moduleItem))
                        End If
                    End If
                    Dim j As Integer
                    j = DataGrid1.CurrentRowIndex

                    moduleItem = DataGrid1.Item(j, 5)
                    DBview.Delete(j)

                    For I As Integer = 0 To DBview.Count - 1
                        'when there are more rows, check if the deleted module name still exist in the datatable
                        If moduleItem = DataGrid1.Item(I, 5) Then
                            Exit Select
                        End If
                    Next
                    CBViewModule.Items.Remove(CString(moduleItem))


                End If
                'Case "Exit"            'Exit
                'Panel1.Hide()
                'MySSMain.PanelRight.Width = 528
                'MySystemConfig.PanelToBeAdded.Width = 528
                'MySSMain.PanelRight.Location = New Point(752, 0)
                'MySystemConfig.PanelToBeAdded.Controls.Remove(MySysIO.Panel1)

            Case "Import"
                ReadText()
            Case "Export"
                WriteText()
        End Select

        RefreshDisplay()
        IOReadStop = False


    End Sub
    Private Sub CB_Type_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Type.SelectedIndexChanged
        Return

        If CB_Type.Text = "Output" Then
            RD_OFF.Checked = True
            RD_On.Enabled = True
            RD_OFF.Enabled = True
        Else
            RD_On.Checked = False
            RD_On.Enabled = False
            RD_OFF.Checked = False
            RD_OFF.Enabled = False
        End If
    End Sub
    Private Sub ReadText()
        Dim fr As StreamReader
        Dim FileString As String
        Dim SplitString() As String
        Dim dataRow As DataRow

        Dim Delimiter As Char = ControlChars.Tab
        Dim row As Integer


        Try
            fr = New StreamReader("C:\SystemIO.txt")
        Catch ex As Exception
            WriteText()
            fr = New StreamReader("C:\SystemIO.txt")
            fr.Close()
            Exit Sub
        End Try

        For row = 0 To DBview.Table.Rows.Count - 1
            dataRow = DBview.Table.Rows(0)
            DBview.Table.Rows.Remove(dataRow)
            DBview.Table.AcceptChanges()
        Next

        FileString = fr.ReadLine
        While Not FileString Is Nothing
            dataRow = DBview.Table.NewRow
            SplitString = FileString.Split(Delimiter)
            dataRow(0) = "FactoryDefault"
            dataRow(1) = SplitString(0)
            dataRow(2) = SplitString(1)
            dataRow(3) = SplitString(2)
            dataRow(4) = SplitString(3)
            dataRow(5) = SplitString(4)
            dataRow(6) = SplitString(5)
            DBview.Table.Rows.Add(dataRow)
            DBview.Table.AcceptChanges()
            FileString = fr.ReadLine
        End While

        fr.Close()

    End Sub
    Private Sub WriteText()
        Dim fw As StreamWriter
        Dim ReadString As String
        Dim row As Integer
        Dim dataRow As DataRow

        fw = New StreamWriter("C:\SystemIO.txt", False)

        For row = 0 To DBview.Table.Rows.Count - 1
            dataRow = DBview.Table.Rows(row)
            ReadString = dataRow(1) & ControlChars.Tab & dataRow(2) & ControlChars.Tab & dataRow(3) & ControlChars.Tab & dataRow(4) & ControlChars.Tab & dataRow(5) & ControlChars.Tab & dataRow(6)
            fw.WriteLine(ReadString)
        Next
        fw.Close()
    End Sub
    Private Sub FiletoTable()

        'IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDS.Data.OpenData()

        DS.Tables("TableSystemIO").Columns("ID").ColumnMapping = MappingType.Hidden
        DBview = New DataView(DS.Tables("TableSystemIO"))
        DS.Tables("TableSystemIO").Clear()
        DS.AcceptChanges()

        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = ""
        Dim I As Integer = 0
        If IDSData.Hardware.SystemIO.ItemArray.Count > 0 Then
            For I = 0 To IDSData.Hardware.SystemIO.ItemArray.Count - 1
                Dim Row As DataRow = DBview.Table.NewRow()
                IDSData.Hardware.SystemIO.Template = CType(IDSData.Hardware.SystemIO.ItemArray(I), _
                                                            CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO)

                Row("ID") = IDS.Data.ParameterID.RecordID
                Row("IOName") = IDSData.Hardware.SystemIO.Template.IOName
                Row("Type") = IDSData.Hardware.SystemIO.Template.Type
                Row("Status") = IDSData.Hardware.SystemIO.Template.Status
                'Row("Status") = ""
                Row("IO") = IDSData.Hardware.SystemIO.Template.IO
                Row("CableName") = IDSData.Hardware.SystemIO.Template.CableName
                Row("ModuleName") = IDSData.Hardware.SystemIO.Template.ModuleName
                DBview.Table.Rows.Add(Row)
            Next
        End If
        DS.AcceptChanges()
    End Sub
    Private Sub TabletoFile()
        'IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDS.Data.OpenData()

        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = ""
        Dim I As Integer = 0
        If DBview.Count > 0 Then
            IDSData.Hardware.SystemIO.ItemArray.Clear()
            For I = 0 To DBview.Count - 1
                Dim Row As DataRow = DBview(I).Row

                IDSData.Hardware.SystemIO.Template = New CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO

                IDS.Data.ParameterID.RecordID = CString(Row("ID"))
                IDSData.Hardware.SystemIO.Template.IOName = CString(Row("IOName"))
                IDSData.Hardware.SystemIO.Template.Type = CString(Row("Type"))
                IDSData.Hardware.SystemIO.Template.Status = CString(Row("Status"))
                IDSData.Hardware.SystemIO.Template.IO = CString(Row("IO"))
                IDSData.Hardware.SystemIO.Template.CableName = CString(Row("CableName"))
                IDSData.Hardware.SystemIO.Template.ModuleName = CString(Row("ModuleName"))

                IDSData.Hardware.SystemIO.ItemArray.Add(IDSData.Hardware.SystemIO.Template)

            Next
        End If

        IDS.Data.SaveData()

    End Sub
    Private Sub Btn_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Cancel.Click
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Call SetIOTable()
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub
    Private Sub Btn_Apply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Apply.Click
        Call TabletoFile()
        RefreshDisplay()
    End Sub
    Private Sub Btn_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_exit.Click

        IOReadClose = True
        Sleep(100)


        If IDS.Data.Admin.User.RunApplication.ToUpper = "SYSTEM" Then

            MySSMain.PanelRight.Width = 528
            MySystemConfig.PanelToBeAdded.Width = 528
            MySSMain.PanelRight.Location = New Point(752, 0)
            MySystemConfig.PanelToBeAdded.Controls.Remove(MySysIO.Panel1)

        Else
            'system setup
            'MyBasicSetup.RadioIO.Checked = False
            MyBasicSetup.PanelRight.Height = 911
            MyBasicSetup.PanelRight.Width = 512
            MyBasicSetup.PanelRight.Location = New Point(768, 33)
            MyBasicSetup.PanelRight.Controls.Remove(MySysIO.Panel1)
        End If


        Dim Respond As DialogResult = MessageBox.Show("Do You want to initialize the IO bits?", "", MessageBoxButtons.YesNo)

        If Respond = DialogResult.Yes Then
            'Initialize PC IO 
            Dim I As Integer = 0


            Dim LocalMC As New CIDSMC
            MC = LocalMC
            IDS.Devices.Motor.SetAllDIOsOff()


            Dim LocalPC As New _CIDSIOcard
            _PC = LocalPC
            For I = 0 To 7
                IDS.Devices.DIO.SetOutputBit(0, I, False)
                IDS.Devices.DIO.SetOutputBit(1, I, False)
            Next

            Dim LocalCAN As New CIDSCAN
            CAN = LocalCAN
            For I = 32 To 47
                Sleep(10)
                IDS.Devices.Motor.SetDIOs(1, I, False)
            Next
        End If
        

    End Sub
    Private Sub BtnViewby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call RefreshDisplay()

    End Sub

#End Region

#Region " Version  25 Jul 2007 Toggle ON/OFF SYSTEM IO - LSGOH "

    Private Function ValidateIONumber(ByRef Mystr As String, ByRef Card As String, ByRef IOPort As String, ByRef IOBit As Integer, ByRef Message As String) As Boolean

        Try
            Dim Str As String() = Mystr.Split("-")
            Card = Str(0).ToUpper
            IOPort = Str(1).ToUpper
            IOBit = Str(2).ToUpper

        Catch ex As Exception
            Message = "Can't Toggle invalid IO Number - " + Mystr.ToUpper
            Return False
        End Try


        If Card = "MC" Then

            If IOPort <> "IN" And IOPort <> "OUT" Then
                Message = "PORT number is not valid - " + IOPort.ToUpper
                Return False
            Else
                If IOBit <= 19 And IOBit >= 0 Then
                Else
                    Message = "BIT number is not valid - " + IOBit.ToString
                    Return False
                End If
            End If

        ElseIf Card = "CAN" Then

            If IOPort <> "I/O" Then
                Message = "PORT number is not valid - " + IOPort.ToUpper
                Return False
            Else
                If IOBit <= 47 And IOBit >= 0 Then
                Else
                    Message = "BIT number is not valid - " + IOBit.ToString

                    Return False
                End If

            End If

        ElseIf Card = "PC" Then

            'If IOPort <> "IDI" And IOPort <> "IDO" Then
            If IOPort <> "IN" And IOPort <> "OUT" Then
                Message = "PORT number is not valid - " + IOPort.ToString
                Return False
            Else
                If IOBit <= 15 And IOBit >= 0 Then
                Else
                    Message = "BIT number is not valid - " + IOBit.ToString
                    Return False
                End If
            End If
        Else
            Message = "CARD number is not valid - " + Card
            Return False
        End If

        Message = ""
        Return True

    End Function
    Private Function ExtractIONumberInfo(ByRef MyStr As String, ByRef Card As String, ByRef IOPort As String, ByRef IOBit As String) As Boolean
        Try
            Dim Str As String() = MyStr.Split("-")
            Card = Str(0).ToUpper
            IOPort = Str(1).ToUpper
            IOBit = Str(2).ToUpper

        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function
    Private Sub ToggleBitOnOff(ByRef BitStatus As Boolean)

        Dim Card As String
        Dim IOPort As String
        Dim IOBit As String
        Dim Message As String
        Dim BitFlag As String

        ' For card Name = MC 
        If IDSData.Hardware.SystemIO.Template.Type.ToUpper = "OUTPUT" Then

            If ValidateIONumber(IDSData.Hardware.SystemIO.Template.IO, Card, IOPort, IOBit, Message) = True Then

                If BitStatus = True Then
                    BitFlag = "ON"
                Else
                    BitFlag = "OFF"
                End If

                If Card = "PC" Then

                    If IOPort = "OUT" Then

                        Dim Portnumber As Integer

                        If IOBit < 8 Then
                            Portnumber = 0
                        Else
                            Portnumber = 1
                            IOBit = IOBit - 8
                        End If

                        If IDS.Devices.DIO.SetOutputBit(Portnumber, IOBit, BitStatus) Then  ' Assumed output as port = 0
                            '    If True Then
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag
                            _PC.OUT.Port(Portnumber).Bit(IOBit) = BitStatus
                        Else
                            MessageBox.Show("PC-IO toggle failed")
                        End If

                    End If

                    ElseIf Card = "MC" Then

                    If IOPort = "OUT" Then
                        ' do something
                        'If True Then
                        If IDS.Devices.Motor.SetDIOs(0, IOBit, BitStatus) Then
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag
                            MC.ExtBit(IOBit) = BitStatus 'add on for display only
                        Else
                            MessageBox.Show("MC-IO toggle failed")
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = "Failed"
                            MC.ExtBit(IOBit) = False  'add on for display only
                        End If

                    End If

                ElseIf Card = "CAN" Then

                    If IOPort = "I/O" Then

                        ' do something
                        'If True Then
                        If IDS.Devices.Motor.SetDIOs(1, IOBit, BitStatus) Then
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag
                            CAN.Bit(IOBit) = BitStatus 'add on for display only
                        Else
                            MessageBox.Show("CAN-IO toggle failed")
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = "Failed"
                            CAN.Bit(IOBit) = False 'add on for display only
                        End If
                    End If

                End If

            Else
            MessageBox.Show(Message)
        End If

        End If
    End Sub

    Private Sub RD_On_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD_On.CheckedChanged

        IOReadStop = True
        Sleep(100)
        Call ToggleBitOnOff(RD_On.Checked)
        IOReadStop = False

    End Sub
    Private Sub TimerIO_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerIO.Tick

        Dim IOBitMap As Integer = 255

        If IOReadStop = True Then

            Application.DoEvents()

        ElseIf IOReadClose = True Then

            TimerIO.Enabled = False
            TimerIO.Stop()
            Return

        Else
            TimerIO.Stop()
            ReadIOFlag = True

            'Update to Datagrid
            Dim Card As String
            Dim IOPort As String
            Dim IOBit As String
            Dim Message As String
            Dim PortNumber As Integer
            Dim I As Integer = 0

            
            Dim a As Integer = 0
            While Not IDS.Devices.Motor.GetDIOs(1, 32, 47, IOBitMap)
                Sleep(1000)
                Application.DoEvents()
                a = a + 1
                If a > 10 Then
                    MessageBox.Show("Can-IO Read Err")
                    Exit While
                End If
            End While
            If a <= 10 Then
                Call ConvertBitMaptoInteger(IOBitMap, "CAN", "", 0)
            End If



            Static Dim failcount As Integer
            Dim ReadByte As Byte
            PortNumber = 0
            If IDS.Devices.DIO.CheckInputByte(PortNumber, ReadByte) Then
                '    If True Then
                IOBitMap = CInt(ReadByte)
                Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 0) 'read input port 0 

            Else
                If failcount < 4 Then
                    MessageBox.Show("PC-IO Read error")
                    failcount += 1
                End If

            End If


            PortNumber = 1
            If IDS.Devices.DIO.CheckInputByte(PortNumber, ReadByte) Then
                '    If True Then
                IOBitMap = CInt(ReadByte)
                Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 1) 'read input port 1
            Else

                If failcount < 4 Then
                    MessageBox.Show("PC-IO error")
                    failcount += 1
                End If

            End If


            a = 0
            While Not IDS.Devices.Motor.GetDIOs(0, 0, 19, IOBitMap)
                Sleep(1000)
                Application.DoEvents()
                a = a + 1
                If a > 10 Then
                    MessageBox.Show("MC-IO Read Err")
                    Exit While
                End If
            End While
            If a <= 10 Then
                Call ConvertBitMaptoInteger(IOBitMap, "MC", "", 0)
            End If


            'Dim Dbview As DataView
            'DBview = New DataView(DS.Tables("TableSystemIO"))
            'DBview.RowFilter = FilterString

            If DBview.Count = 0 Then
                'Ignore if there is no data rows in the table
            Else
                I = 0

                While True

                    Dim Row As DataRow = DBview(I).Row

                    If IsDBNull(Row("Type")) Then
                        GoTo IOList_End
                    End If

                    Dim IOType As String = Row("Type")
                    If IOType.ToUpper <> "INPUT" And IOType.ToUpper <> "OUTPUT" Then
                    Else
                        Dim IONumber As String = Row("IO")
                        If ValidateIONumber(IONumber, Card, IOPort, IOBit, Message) = True Then

                            If Card = "PC" Then

                                If IOBit < 8 Then
                                    PortNumber = 0
                                Else
                                    PortNumber = 1
                                    IOBit = IOBit - 8
                                End If

                                If IOPort = "IN" Then

                                    If _PC.INP.Port(PortNumber).Bit(IOBit) = True Then
                                        Row("Status") = "ON"
                                    ElseIf _PC.INP.Port(PortNumber).Bit(IOBit) = False Then
                                        Row("Status") = "OFF"
                                    End If

                                ElseIf IOPort = "OUT" Then

                                    If _PC.OUT.Port(PortNumber).Bit(IOBit) = True Then
                                        Row("Status") = "ON"
                                    ElseIf _PC.OUT.Port(PortNumber).Bit(IOBit) = False Then
                                        Row("Status") = "OFF"
                                    End If

                                End If



                            ElseIf Card = "MC" Then

                                If MC.Bit(IOBit) = True Then
                                    Row("Status") = "ON"
                                ElseIf MC.Bit(IOBit) = False Then
                                    Row("Status") = "OFF"
                                End If



                            ElseIf Card = "CAN" Then

                                If CAN.Bit(IOBit) = True Then
                                    Row("Status") = "ON"
                                ElseIf CAN.Bit(IOBit) = False Then
                                    Row("Status") = "OFF"
                                End If

                            Else

                                Row("Status") = ""

                            End If

                        Else
                            Row("Status") = ""
                        End If
                    End If

                    Application.DoEvents()
                    DataGrid1.CaptionText = "Refreshing Record = " + CStr(I + 1) + " of " + DBview.Count.ToString

                    If IOReadStop = True Then
                        Exit While
                    End If

                    If I < DBview.Count - 1 Then
                        I += 1
                    Else
                        Exit While
                    End If
                End While
            End If

IOList_End:


            ReadIOFlag = False
            TimerIO.Start()
        End If
    End Sub

    Private Sub ConvertBitMaptoInteger(ByVal BitMap As Integer, ByVal Card As String, ByVal IOtype As String, ByVal PortNumber As Integer)

        If Card = "PC" Then

            If IOtype = "IN" Then
                If (BitMap And &H1) Then _PC.INP.Port(PortNumber).Bit(0) = True Else _PC.INP.Port(PortNumber).Bit(0) = False
                If (BitMap And &H2) Then _PC.INP.Port(PortNumber).Bit(1) = True Else _PC.INP.Port(PortNumber).Bit(1) = False
                If (BitMap And &H4) Then _PC.INP.Port(PortNumber).Bit(2) = True Else _PC.INP.Port(PortNumber).Bit(2) = False
                If (BitMap And &H8) Then _PC.INP.Port(PortNumber).Bit(3) = True Else _PC.INP.Port(PortNumber).Bit(3) = False

                If (BitMap And &H10) Then _PC.INP.Port(PortNumber).Bit(4) = True Else _PC.INP.Port(PortNumber).Bit(4) = False
                If (BitMap And &H20) Then _PC.INP.Port(PortNumber).Bit(5) = True Else _PC.INP.Port(PortNumber).Bit(5) = False
                If (BitMap And &H40) Then _PC.INP.Port(PortNumber).Bit(6) = True Else _PC.INP.Port(PortNumber).Bit(6) = False
                If (BitMap And &H80) Then _PC.INP.Port(PortNumber).Bit(7) = True Else _PC.INP.Port(PortNumber).Bit(7) = False

            End If

        ElseIf Card = "MC" Then

            'basic bits
            If (BitMap And &H1) Then MC.Bit(0) = True Else MC.Bit(0) = False
            If (BitMap And &H2) Then MC.Bit(1) = True Else MC.Bit(1) = False
            If (BitMap And &H4) Then MC.Bit(2) = True Else MC.Bit(2) = False
            If (BitMap And &H8) Then MC.Bit(3) = True Else MC.Bit(3) = False

            If (BitMap And &H10) Then MC.Bit(4) = True Else MC.Bit(4) = False
            If (BitMap And &H20) Then MC.Bit(5) = True Else MC.Bit(5) = False
            If (BitMap And &H40) Then MC.Bit(6) = True Else MC.Bit(6) = False
            If (BitMap And &H80) Then MC.Bit(7) = True Else MC.Bit(7) = False

            If (BitMap And &H100) Then MC.Bit(8) = True Else MC.Bit(8) = False
            If (BitMap And &H200) Then MC.Bit(9) = True Else MC.Bit(9) = False
            If (BitMap And &H400) Then MC.Bit(10) = True Else MC.Bit(10) = False
            If (BitMap And &H800) Then MC.Bit(11) = True Else MC.Bit(11) = False

            If (BitMap And &H1000) Then MC.Bit(12) = True Else MC.Bit(12) = False
            If (BitMap And &H2000) Then MC.Bit(13) = True Else MC.Bit(13) = False
            If (BitMap And &H4000) Then MC.Bit(14) = True Else MC.Bit(14) = False
            If (BitMap And &H8000) Then MC.Bit(15) = True Else MC.Bit(15) = False

            If (BitMap And &H10000) Then MC.Bit(16) = True Else MC.Bit(16) = False
            If (BitMap And &H20000) Then MC.Bit(17) = True Else MC.Bit(17) = False
            If (BitMap And &H40000) Then MC.Bit(18) = True Else MC.Bit(18) = False
            If (BitMap And &H80000) Then MC.Bit(19) = True Else MC.Bit(19) = False

            'extension bits
            If (BitMap And &H1) Then MC.ExtBit(8) = True Else MC.ExtBit(8) = False
            If (BitMap And &H2) Then MC.ExtBit(9) = True Else MC.ExtBit(9) = False
            If (BitMap And &H4) Then MC.ExtBit(10) = True Else MC.ExtBit(10) = False
            If (BitMap And &H8) Then MC.ExtBit(11) = True Else MC.ExtBit(11) = False

            If (BitMap And &H10) Then MC.ExtBit(12) = True Else MC.ExtBit(12) = False
            If (BitMap And &H20) Then MC.ExtBit(13) = True Else MC.ExtBit(13) = False
            If (BitMap And &H40) Then MC.ExtBit(14) = True Else MC.ExtBit(14) = False
            If (BitMap And &H80) Then MC.ExtBit(15) = True Else MC.ExtBit(15) = False

            If (BitMap And &H100) Then MC.ExtBit(16) = True Else MC.ExtBit(16) = False
            If (BitMap And &H200) Then MC.ExtBit(17) = True Else MC.ExtBit(17) = False


        ElseIf Card = "CAN" Then

            If (BitMap And &H1) Then CAN.Bit(32) = True Else CAN.Bit(32) = False
            If (BitMap And &H2) Then CAN.Bit(33) = True Else CAN.Bit(33) = False
            If (BitMap And &H4) Then CAN.Bit(34) = True Else CAN.Bit(34) = False
            If (BitMap And &H8) Then CAN.Bit(35) = True Else CAN.Bit(35) = False

            If (BitMap And &H10) Then CAN.Bit(36) = True Else CAN.Bit(36) = False
            If (BitMap And &H20) Then CAN.Bit(37) = True Else CAN.Bit(37) = False
            If (BitMap And &H40) Then CAN.Bit(38) = True Else CAN.Bit(38) = False
            If (BitMap And &H80) Then CAN.Bit(39) = True Else CAN.Bit(39) = False

            If (BitMap And &H100) Then CAN.Bit(40) = True Else CAN.Bit(40) = False
            If (BitMap And &H200) Then CAN.Bit(41) = True Else CAN.Bit(41) = False
            If (BitMap And &H400) Then CAN.Bit(42) = True Else CAN.Bit(42) = False
            If (BitMap And &H800) Then CAN.Bit(43) = True Else CAN.Bit(43) = False

            If (BitMap And &H1000) Then CAN.Bit(44) = True Else CAN.Bit(44) = False
            If (BitMap And &H2000) Then CAN.Bit(45) = True Else CAN.Bit(45) = False
            If (BitMap And &H4000) Then CAN.Bit(46) = True Else CAN.Bit(46) = False
            If (BitMap And &H8000) Then CAN.Bit(47) = True Else CAN.Bit(47) = False

        End If

    End Sub


    Dim FilterString As String
    Friend ReadIOFlag As Boolean
    Friend IOReadStop As Boolean
    Friend IOReadClose As Boolean


    Public _PC As New _CIDSIOcard
    Public Class _CIDSIOcard

        Public OUT As New CIDSIOType(1)
        Public INP As New CIDSIOType(1)
        Public Class CIDSIOType
            Public Sub New(ByVal _NumPort As Integer)
                Call NumPort(_NumPort)
            End Sub
            Public Port() As CIDSIOPort
            Public Sub NumPort(ByVal Number As Integer)
                ReDim Port(Number)
                Dim I As Integer
                For I = 0 To Number
                    Port(I) = New CIDSIOPort
                Next
            End Sub
            Public Class CIDSIOPort
                Public Bit() As Boolean = {0, 0, 0, 0, 0, 0, 0, 0}
            End Class
        End Class
    End Class


    Public PC As New CIDSIOCard(1)

    'Public CAN As New CIDSIOCard(0)
    'Public MC As New CIDSIOCard(0)
    Public Class CIDSIOCard
        Public Sub New(ByVal _NumPort As Integer)
            Call NumPort(_NumPort)
        End Sub
        Public Port() As CIDSIOPort
        Public Sub NumPort(ByVal Number As Integer)
            ReDim Port(Number)
            Dim I As Integer
            For I = 0 To Number
                Port(I) = New CIDSIOPort
            Next
        End Sub
        Public Class CIDSIOPort
            Public Bit(15) As Boolean
        End Class
    End Class

    Public MC As New CIDSMC
    Public Class CIDSMC
        Public Bit() As Boolean = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Public ExtBit() As Boolean = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    End Class

    Public CAN As New CIDSCAN
    Public Class CIDSCAN
        Public Bit() As Boolean = _
        {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0}

    End Class



#End Region

   
    Private Sub RD_OFF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD_OFF.CheckedChanged

    End Sub
End Class
