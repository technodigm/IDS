Imports DLL_DataManager
Imports DLL_Export_Service
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
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label As System.Windows.Forms.Label
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
    Friend WithEvents ButtonRevert As System.Windows.Forms.Button
    Friend WithEvents ButtonSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ButtonExit As System.Windows.Forms.Button
    Friend WithEvents TBr_SysIO As System.Windows.Forms.ToolBar
    Friend WithEvents ToolBarButtonNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonEdit As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonDelete As System.Windows.Forms.ToolBarButton
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents CBViewType As System.Windows.Forms.ComboBox
    Friend WithEvents CBViewModule As System.Windows.Forms.ComboBox
    Friend WithEvents TimerIO As System.Windows.Forms.Timer
    Friend WithEvents SaveFileDlg As System.Windows.Forms.SaveFileDialog
    Friend WithEvents OpenFileDlg As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ToolBarButtonImportTextfile As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonExportTextFile As System.Windows.Forms.ToolBarButton
    Friend WithEvents ToolBarButtonImportPatFile As System.Windows.Forms.ToolBarButton
    Friend WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(SystemIO))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TBr_SysIO = New System.Windows.Forms.ToolBar
        Me.ToolBarButtonNew = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonEdit = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonDelete = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonImportPatFile = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonImportTextfile = New System.Windows.Forms.ToolBarButton
        Me.ToolBarButtonExportTextFile = New System.Windows.Forms.ToolBarButton
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
        Me.ButtonRevert = New System.Windows.Forms.Button
        Me.ButtonSave = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Label = New System.Windows.Forms.Label
        Me.ButtonExit = New System.Windows.Forms.Button
        Me.TimerIO = New System.Windows.Forms.Timer(Me.components)
        Me.SaveFileDlg = New System.Windows.Forms.SaveFileDialog
        Me.OpenFileDlg = New System.Windows.Forms.OpenFileDialog
        Me.PanelToBeAdded.SuspendLayout()
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
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.SystemColors.Control
        Me.PanelToBeAdded.Controls.Add(Me.Panel2)
        Me.PanelToBeAdded.Controls.Add(Me.Label9)
        Me.PanelToBeAdded.Controls.Add(Me.Label8)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonRevert)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonSave)
        Me.PanelToBeAdded.Controls.Add(Me.ComboBox1)
        Me.PanelToBeAdded.Controls.Add(Me.Label)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonExit)
        Me.PanelToBeAdded.Location = New System.Drawing.Point(0, 0)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(1280, 944)
        Me.PanelToBeAdded.TabIndex = 0
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
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(8, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1080, 208)
        Me.GroupBox1.TabIndex = 55
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "System Information Entry Tools"
        '
        'TBr_SysIO
        '
        Me.TBr_SysIO.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButtonNew, Me.ToolBarButtonEdit, Me.ToolBarButtonDelete, Me.ToolBarButtonImportPatFile, Me.ToolBarButtonImportTextfile, Me.ToolBarButtonExportTextFile})
        Me.TBr_SysIO.ButtonSize = New System.Drawing.Size(100, 45)
        Me.TBr_SysIO.Dock = System.Windows.Forms.DockStyle.None
        Me.TBr_SysIO.DropDownArrows = True
        Me.TBr_SysIO.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TBr_SysIO.ImageList = Me.ImageList1
        Me.TBr_SysIO.Location = New System.Drawing.Point(744, 48)
        Me.TBr_SysIO.Name = "TBr_SysIO"
        Me.TBr_SysIO.ShowToolTips = True
        Me.TBr_SysIO.Size = New System.Drawing.Size(312, 51)
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
        'ToolBarButtonImportPatFile
        '
        Me.ToolBarButtonImportPatFile.Text = "Import PatFile"
        Me.ToolBarButtonImportPatFile.ToolTipText = "Import Pat Encrypted File"
        Me.ToolBarButtonImportPatFile.Visible = False
        '
        'ToolBarButtonImportTextfile
        '
        Me.ToolBarButtonImportTextfile.Text = "Import TextFile"
        Me.ToolBarButtonImportTextfile.ToolTipText = "Import Pat Text File"
        Me.ToolBarButtonImportTextfile.Visible = False
        '
        'ToolBarButtonExportTextFile
        '
        Me.ToolBarButtonExportTextFile.Text = "Export TextFile"
        Me.ToolBarButtonExportTextFile.ToolTipText = "Export Pat Text file"
        Me.ToolBarButtonExportTextFile.Visible = False
        '
        'TB_CableLabel
        '
        Me.TB_CableLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_CableLabel.Location = New System.Drawing.Point(848, 160)
        Me.TB_CableLabel.Name = "TB_CableLabel"
        Me.TB_CableLabel.ReadOnly = True
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
        Me.TB_IONum.ReadOnly = True
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
        Me.CB_Type.Size = New System.Drawing.Size(136, 24)
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
        Me.TB_ModuleName.ReadOnly = True
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
        Me.Label2.Size = New System.Drawing.Size(115, 23)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Module Name"
        '
        'TB_IOName
        '
        Me.TB_IOName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_IOName.Location = New System.Drawing.Point(144, 128)
        Me.TB_IOName.Name = "TB_IOName"
        Me.TB_IOName.ReadOnly = True
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
        Me.CBViewType.Size = New System.Drawing.Size(128, 24)
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
        Me.CBViewModule.Size = New System.Drawing.Size(328, 24)
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
        'ButtonRevert
        '
        Me.ButtonRevert.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonRevert.Location = New System.Drawing.Point(1040, 864)
        Me.ButtonRevert.Name = "ButtonRevert"
        Me.ButtonRevert.Size = New System.Drawing.Size(96, 50)
        Me.ButtonRevert.TabIndex = 57
        Me.ButtonRevert.Text = "Cancel"
        '
        'ButtonSave
        '
        Me.ButtonSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonSave.Location = New System.Drawing.Point(936, 864)
        Me.ButtonSave.Name = "ButtonSave"
        Me.ButtonSave.Size = New System.Drawing.Size(100, 50)
        Me.ButtonSave.TabIndex = 56
        Me.ButtonSave.Text = "Apply"
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
        'ButtonExit
        '
        Me.ButtonExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonExit.Image = CType(resources.GetObject("ButtonExit.Image"), System.Drawing.Image)
        Me.ButtonExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonExit.Location = New System.Drawing.Point(1184, 16)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(75, 50)
        Me.ButtonExit.TabIndex = 68
        Me.ButtonExit.TabStop = False
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'TimerIO
        '
        '
        'SystemIO
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1280, 912)
        Me.Controls.Add(Me.PanelToBeAdded)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "SystemIO"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "SystemIO"
        Me.PanelToBeAdded.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    'declare temporary working variable for data tabe and column 

    Dim DBview As New DataView
    Dim TableSystemIO As New DataTable

    Dim dgtablestyle As New DataGridTableStyle
    Dim IONameCol As New DataGridTextBoxColumn
    Dim ModuleNameCol As New DataGridTextBoxColumn
    Dim TypeCol As New DataGridTextBoxColumn
    Dim StatusCol As New DataGridTextBoxColumn
    Dim IONumCol As New DataGridTextBoxColumn
    Dim CableNameCol As New DataGridTextBoxColumn
    Dim Itemcol As New DataGridTextBoxColumn

    ' format the table styple for datatable
    Private Sub formatcol()

        dgtablestyle.MappingName = "TableSystemIO"

        ' define column properties

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

        ''''''''''''''
        'add column styple to data table
        ''''''''''''''

        dgtablestyle.GridColumnStyles.Add(Itemcol)
        dgtablestyle.GridColumnStyles.Add(IONameCol)
        dgtablestyle.GridColumnStyles.Add(TypeCol)
        dgtablestyle.GridColumnStyles.Add(StatusCol)
        dgtablestyle.GridColumnStyles.Add(IONumCol)
        dgtablestyle.GridColumnStyles.Add(CableNameCol)
        dgtablestyle.GridColumnStyles.Add(ModuleNameCol)

        DataGrid1.TableStyles.Add(dgtablestyle)
    End Sub
    ''''''''''''''
    ' comobo box for view module is selected 
    ''''''''''''''
    Private Sub CBViewModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBViewModule.SelectedIndexChanged
        RefreshDisplay()
    End Sub
    ''''''''''''''
    ' comobo box for view type is selected 
    ''''''''''''''
    Private Sub CBViewType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBViewType.SelectedIndexChanged
        RefreshDisplay()
    End Sub
    ''''''''''''''
    ' refresh the screen with the user selection views
    ''''''''''''''
    Private Sub RefreshDisplay()

        IOReadStop = True                   'io reading stop flag is now true
        mysleep(100)                          'delay 100 minisec

        '''''
        ' set to false for refreshing purpose
        '''''

        IONameCol.ReadOnly = False          '
        TypeCol.ReadOnly = False
        'StatusCol.ReadOnly = False
        IONumCol.ReadOnly = False
        CableNameCol.ReadOnly = False
        ModuleNameCol.ReadOnly = False
        Itemcol.ReadOnly = False

        ''''
        ' to set the filter for item to be view for data table
        '''''
        If CBViewModule.Text = "ALL" And CBViewType.Text = "ALL" Then
            FilterString = ""

        ElseIf CBViewModule.Text <> "ALL" And CBViewType.Text = "ALL" Then
            FilterString = "ModuleName = '" + CBViewModule.Text + "'"

        ElseIf CBViewModule.Text = "ALL" And CBViewType.Text <> "ALL" Then
            FilterString = "Type = '" + CBViewType.Text + "'"

        ElseIf CBViewModule.Text <> "ALL" And CBViewType.Text <> "ALL" Then
            FilterString = "ModuleName = '" + CBViewModule.Text + "' AND Type = '" + CBViewType.Text + "'"
        End If

        ''''
        ' intialize the datagrid for clearing the its screen 
        '''
        Dim _datatable As New DataTable
        DBview = New DataView(_datatable)
        DataGrid1.DataSource = DBview

        '''''
        ' open table system IO
        ' with filter on 
        '''''

        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = FilterString
        DataGrid1.DataSource = DBview


        If DBview.Count = 0 Then
            'Ignore if there is no data rows in the table
        Else

            DBview.Sort = "CableName Asc"   'sort the dataview in ascending order

            ''''''
            '   Check if each module name is listed 
            '   in the view module combobox 
            '   if yes, dun add, else add item to the 
            '   comobo box
            ''''''

            Dim moduleItem As Object
            For I As Integer = 0 To DBview.Count - 1
                Dim Row As DataRow = DBview(I).Row
                Row("Item") = I + 1
                moduleItem = Row(6)
                If CBViewModule.Items.Contains(moduleItem) Then
                Else
                    CBViewModule.Items.Add(CString(moduleItem))
                End If
            Next
        End If

        '''''
        ' reset to true of its original state
        '''''
        IONameCol.ReadOnly = True
        TypeCol.ReadOnly = True
        'StatusCol.ReadOnly = True
        IONumCol.ReadOnly = True
        CableNameCol.ReadOnly = True
        ModuleNameCol.ReadOnly = True
        Itemcol.ReadOnly = True

        '''''
        ' call the click function of datagrid
        '''''
        Call DataGrid1_Click(Me, EArg)
        IOReadStop = False

    End Sub

    '''''
    '   
    '''''
    Friend Sub SetIOTable()

        '''
        ' set visible to true if is running in system mode
        ' else set to false
        '''
        If IDS.Data.Admin.User.RunApplication = "System" Then
            ButtonSave.Visible = True
            ButtonRevert.Visible = True
            TBr_SysIO.Visible = True
        Else

            ButtonSave.Visible = False
            ButtonRevert.Visible = False
            TBr_SysIO.Visible = False
        End If

        '''
        ' initialize the comobox view module with item ALL only
        '''
        CBViewModule.Items.Clear()
        CBViewModule.Items.Add("ALL")
        CBViewModule.Text = "ALL"



        '''
        ' fill the data from pat file to datatable / dataview
        ' and refresh the display 
        '''
        Call FiletoTable()
        DS.Tables("TableSystemIO").Columns("ID").ColumnMapping = MappingType.Hidden
        DBview = New DataView(DS.Tables("TableSystemIO"))
        DataGrid1.DataSource() = DBview.DataViewManager
        DataGrid1.CaptionText = DataGrid1.DataMember
        Call RefreshDisplay()
    End Sub

    Friend Sub EditIO()
        If IDS.Data.Admin.User.RunApplication = "System" Then
            TB_IOName.ReadOnly = False
            TB_ModuleName.ReadOnly = False
            TB_IONum.ReadOnly = False
            TB_CableLabel.ReadOnly = False

            Label1.Enabled = True
            Label2.Enabled = True
            Label3.Enabled = True
            Label5.Enabled = True
            Label6.Enabled = True
            Label7.Enabled = True
            Label10.Enabled = True

            RD_On.Enabled = True
            RD_OFF.Enabled = True
        Else
            TB_IOName.ReadOnly = True
            TB_ModuleName.ReadOnly = True
            TB_IONum.ReadOnly = True
            TB_CableLabel.ReadOnly = True

            Label1.Enabled = False
            Label2.Enabled = False
            Label3.Enabled = False
            Label5.Enabled = False
            Label6.Enabled = False
            Label7.Enabled = False
            Label10.Enabled = False

            RD_On.Enabled = False
            RD_OFF.Enabled = False
        End If
    End Sub

    '''
    ' current cell changed event of datagrid
    '''
    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.CurrentCellChanged
        DataGrid1_Click(sender, e) 'call to the click event of datagrid
    End Sub
    '''
    ' key up event of datagrid
    '''
    Private Sub DataGrid1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGrid1.KeyUp
        If e.KeyCode = Keys.Delete Then
            Call RefreshDisplay()   'call refresh display function
        End If
    End Sub
    '''
    ' click event of datagrid
    '''
    Private Sub DataGrid1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGrid1.Click

        '''
        ' declare temp working variables
        '''

        Dim Status As String
        Dim CurrentItem As Object
        Dim CurrentRowIndex As Integer = DataGrid1.CurrentCell.RowNumber()
        Dim i As Integer = 0

        ''''''''''''
        ' to extract each of all the column value of a selected row
        '''''''''''''
        While (True)

            Try
                If IsDBNull(DataGrid1.Item(CurrentRowIndex, i)) Then    'if is null value, means error 
                    Return                                              'exit function 
                End If
            Catch ex As Exception                                       'error occured, exit function
                Return
            End Try

            CurrentItem = DataGrid1.Item(CurrentRowIndex, i)            'get value of the selected row and col
            If i = 1 Then                                               'if is col 1
                TB_IOName.Text = CStr(CurrentItem)                      'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.IOName = TB_IOName.Text
            End If

            If i = 2 Then                                               'if is col 2
                CB_Type.Text = CStr(CurrentItem)                        'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.Type = CB_Type.Text
            End If

            If i = 3 Then                                               'if is col 3
                Status = CStr(CurrentItem)                              'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.Status = Status
            End If

            If i = 4 Then                                               'if is col 4
                TB_IONum.Text = CStr(CurrentItem)                       'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.IO = TB_IONum.Text
            End If

            If i = 5 Then                                               'if is col 5
                TB_CableLabel.Text = CStr(CurrentItem)                  'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.CableName = TB_CableLabel.Text
            End If

            If i = 6 Then                                               'if is col 6
                TB_ModuleName.Text = CStr(CurrentItem)                  'retreive value to the display item n globe variable
                IDSData.Hardware.SystemIO.Template.ModuleName = TB_ModuleName.Text
                Exit While
            End If
            i = i + 1
        End While


        ''''''''''''
        ' enable the radio button if combox type is output
        '''''''''''''

        If CB_Type.Text.ToUpper = "OUTPUT" Then
            RD_OFF.Enabled = True
            RD_On.Enabled = True

            ''''''''''''
            ' checked the radio button if combox status is ON
            '''''''''''''
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
    '''''''''''''''
    ' button of toolbar are clicked
    '''''''''''''''
    Private Sub TBr_SysIO_ButtonClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBr_SysIO.ButtonClick

        IOReadStop = True       'io reading stop is on
        mysleep(100)              'delay 


        ''''''''''''
        ' declare temp working variables
        '''''''''''''

        Dim Card As String
        Dim IOPort As String
        Dim IOBit As String
        Dim Message As String
        Dim BitFlag As String
        DataGrid1.Refresh()

        ''''''''''''
        ' button type selection 
        '''''''''''''
        Select Case e.Button.Text
            Case "AddRow"
                '''''''''''''''
                'look for rows with same cablelabel text
                '''''''''''''''
                Dim Row As DataRow = DBview.Table.Rows.Find(TB_CableLabel.Text)

                If row Is Nothing Then                          'if is not found
                    Row = DBview.Table.NewRow()                 'create new row
                    Row("ID") = "FactoryDefault"                'save value
                    Row("IOName") = TB_IOName.Text
                    Row("ModuleName") = TB_ModuleName.Text
                    Row("IO") = TB_IONum.Text
                    Row("CableName") = TB_CableLabel.Text
                    Row("Status") = " "
                    Row("Type") = CB_Type.Text
                    DBview.Table.Rows.Add(Row)                  'add to datatable

                    '''''''''''''''
                    'New module added to the combobox view module
                    '''''''''''''''
                    If CBViewModule.Items.Contains(TB_ModuleName.Text) Then
                    Else
                        CBViewModule.Items.Add(TB_ModuleName.Text)
                        CBViewModule.SelectedItem = TB_ModuleName.Text
                        RefreshDisplay()                        'call to refresh the display
                    End If
                Else
                    '''''''''''''''
                    'if there is row with the same cable name, prompt error message
                    ''''''''''''''
                    MessageBox.Show("Cable Name: " + TB_CableLabel.Text + " already exists.")
                    Exit Select
                End If


            Case "UpdateRow"

                '''''''''''''''
                'look for rows with same cablelabel text
                '''''''''''''''
                Dim Row As DataRow = DBview.Table.Rows.Find(TB_CableLabel.Text)

                If row Is Nothing Then                          'if not found, 
                    MessageBox.Show("Please use Add button.")   'prompt error
                    Exit Select
                Else
                    Row("IOName") = TB_IOName.Text              'else, save value to row
                    Row("ModuleName") = TB_ModuleName.Text
                    Row("IO") = TB_IONum.Text
                    Row("CableName") = TB_CableLabel.Text
                    Row("Type") = CB_Type.Text
                    Row("Status") = " "

                    '''''''''''''''
                    'New module added to the combobox view module
                    '''''''''''''''
                    If CBViewModule.Items.Contains(TB_ModuleName.Text) Then
                    Else
                        CBViewModule.Items.Add(TB_ModuleName.Text)
                    End If
                End If


            Case "DeleteRow"
                Dim moduleItem As Object
                If DBview.Count = 0 Then
                Else
                    If DBview.Count = 1 Then
                        '''''''''
                        'remove the module name from combobox, if is the last row 
                        '''''''''
                        moduleItem = DataGrid1.Item(DBview.Count - 1, 5)
                        If CBViewModule.Items.Contains(moduleItem) Then
                            CBViewModule.Items.Remove(CString(moduleItem))
                        End If
                    End If
                    Dim j As Integer
                    j = DataGrid1.CurrentRowIndex
                    moduleItem = DataGrid1.Item(j, 5)
                    DBview.Delete(j)    'delete row number j

                    '''''''''
                    ' when there are more rows, check if the deleted module name 
                    ' still exist in the datatable
                    '''''''''
                    For I As Integer = 0 To DBview.Count - 1
                        If moduleItem = DataGrid1.Item(I, 5) Then
                            Exit Select
                        End If
                    Next
                    CBViewModule.Items.Remove(CString(moduleItem))
                End If


                '''''''''
                ' import pat file in text format to the globel variable
                '  - not used
                '''''''''

            Case "Import PatFile"

                Dim Respond As DialogResult = MessageBox.Show("System variables will be overwritten by loaded file data, are you sure?", "", MessageBoxButtons.YesNo)

                If Respond = DialogResult.Yes Then

                    OpenFileDlg.InitialDirectory = IDSData.Admin.Folder.PatternPath
                    OpenFileDlg.Filter = "Text files (*.pat)|*.pat"
                    OpenFileDlg.ShowDialog()

                    Textflag = False

                    IDS.Data.OpenPathFileData(OpenFileDlg.FileName)
                    Textflag = False
                End If



            Case "Import TextFile"

                Dim Respond As DialogResult = MessageBox.Show("System variables will be overwritten by loaded file data, are you sure?", "", MessageBoxButtons.YesNo)

                If Respond = DialogResult.Yes Then

                    OpenFileDlg.InitialDirectory = IDSData.Admin.Folder.PatternPath
                    OpenFileDlg.Filter = "Text files (*.txt)|*.txt"
                    OpenFileDlg.ShowDialog()

                    Textflag = True
                    IDS.Data.OpenPathFileData(OpenFileDlg.FileName)
                    Textflag = False
                End If

                '''''''''
                ' Export pat file in text format to the globel variable
                '  - not used
                '''''''''
            Case "Export TextFile"

                SaveFileDlg.InitialDirectory = IDSData.Admin.Folder.PatternPath
                SaveFileDlg.Filter = "Text files (*.txt)|*.txt|" & "All files|*.*"
                SaveFileDlg.ShowDialog()

                Textflag = True

                'IDS.Data.SavePathFileData(SaveFileDlg.FileName)
                SavePathFileName(SaveFileDlg.FileName)
                Textflag = False



        End Select

        RefreshDisplay()    'call to refresh the display screen
        IOReadStop = False  'Io reading resume


    End Sub

    '''''''''
    ' to read from pat file to data table
    '''''''''    
    Private Sub FiletoTable()

        IDS.Data.OpenData() 'open file n read data

        '''''''''
        ' open table system io and clear it content 
        '''''''''    
        DS.Tables("TableSystemIO").Columns("ID").ColumnMapping = MappingType.Hidden
        DBview = New DataView(DS.Tables("TableSystemIO"))
        DS.Tables("TableSystemIO").Clear()
        DS.AcceptChanges()

        '''''''''
        ' open table system io and fill it up with the goble variables
        '''''''''    


        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = ""
        Dim I As Integer = 0
        If IDSData.Hardware.SystemIO.ItemArray.Count > 0 Then

            ' iterate thru the arrary item of goble variable system io
            For I = 0 To IDSData.Hardware.SystemIO.ItemArray.Count - 1
                Dim Row As DataRow = DBview.Table.NewRow()  'create new row

                'get the value from the array to the template
                IDSData.Hardware.SystemIO.Template = CType(IDSData.Hardware.SystemIO.ItemArray(I), _
                                                            CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO)
                ' Row("ID") = IDS.Data.ParameterID.RecordID
                Row("IOName") = IDSData.Hardware.SystemIO.Template.IOName
                Row("Type") = IDSData.Hardware.SystemIO.Template.Type
                Row("Status") = IDSData.Hardware.SystemIO.Template.Status
                'Row("Status") = ""
                Row("IO") = IDSData.Hardware.SystemIO.Template.IO
                Row("CableName") = IDSData.Hardware.SystemIO.Template.CableName
                Row("ModuleName") = IDSData.Hardware.SystemIO.Template.ModuleName

                Dim ExistRow As DataRow = DBview.Table.Rows.Find(IDSData.Hardware.SystemIO.Template.CableName)

                If ExistRow Is Nothing Then
                    DBview.Table.Rows.Add(Row)                  'add row to table
                Else
                    MessageBox.Show("cablename = " + IDSData.Hardware.SystemIO.Template.CableName + " already appear in another row, it will not be added to the display table")
                End If
                'DBview.Table.Rows.Add(Row)                  'add row to table
            Next
        End If
        DS.AcceptChanges()                                  'accept changes
    End Sub
    '''''''''
    ' to write from data table to pat file
    '''''''''   
    Private Sub TabletoFile()

        IDS.Data.OpenData()     'open file and read data

        '''''''''
        ' open table system io and write it contents to the pat file
        '''''''''   
        DBview = New DataView(DS.Tables("TableSystemIO"))
        DBview.RowFilter = ""
        Dim I As Integer = 0
        If DBview.Count > 0 Then
            IDSData.Hardware.SystemIO.ItemArray.Clear()     'initialize the system io array
            For I = 0 To DBview.Count - 1                   'iterate thru the rows of datatable
                Dim Row As DataRow = DBview(I).Row          'get row info

                'create new template of systemio class
                IDSData.Hardware.SystemIO.Template = New CIDSData.CIDSHardWare.CIDSSystemIO.CIDSIO

                'save row info to the template
                'IDS.Data.ParameterID.RecordID = CString(Row("ID"))
                IDSData.Hardware.SystemIO.Template.IO = CString(Row("IOName"))
                IDSData.Hardware.SystemIO.Template.IOName = CString(Row("IOName"))
                IDSData.Hardware.SystemIO.Template.Type = CString(Row("Type"))
                IDSData.Hardware.SystemIO.Template.Status = CString(Row("Status"))
                IDSData.Hardware.SystemIO.Template.IO = CString(Row("IO"))
                IDSData.Hardware.SystemIO.Template.CableName = CString(Row("CableName"))
                IDSData.Hardware.SystemIO.Template.ModuleName = CString(Row("ModuleName"))

                'add the template data to the system io array
                IDSData.Hardware.SystemIO.ItemArray.Add(IDSData.Hardware.SystemIO.Template)

            Next
        End If

        IDS.Data.SaveLocalData() 'save the globe variable to the pat file

    End Sub

    ' cancel button is clicked
    Private Sub ButtonRevert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRevert.Click
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Call SetIOTable()   ' call to set IO table function
        Me.Cursor = System.Windows.Forms.Cursors.Default
    End Sub

    ' apply button is clicked
    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        Call TabletoFile()  ' call to save datatable to the pat file
        RefreshDisplay()    ' refresh the display panel
    End Sub

    ' exit button is clicked
    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click

        IOReadClose = True      'io read connection is closed
        mysleep(100)              'delay 

        ' remove panel from the form 
        RemovePanel(CurrentControl)
        If IDS.Data.Admin.User.RunApplication.ToUpper = "SYSTEM" Then
            MySetup.PanelRight.Width = 528
            MySetup.PanelRight.Location = New Point(752, 0)
        Else
            MySettings.PanelRight.Height = 911
            MySettings.PanelRight.Width = 512
            MySettings.PanelRight.Location = New Point(768, 33)
        End If

        ' prompt user to initialize ALL IO bits
        Dim Respond As DialogResult = MessageBox.Show("Do You want to initialize the IO bits?", "", MessageBoxButtons.YesNo)

        If Respond = DialogResult.Yes Then

            'Initialize trio controller IO 
            Dim I As Integer = 0
            Dim LocalMC As New CIDSMC
            MC = LocalMC
            IDS.Devices.Motor.SetAllDIOsOff()

            'Initialize PCIO card
            Dim LocalPC As New _CIDSIOcard
            _PC = LocalPC
            For I = 0 To 7   'iterate thru the io bits
                IDS.Devices.DIO.SetOutputBit(0, I, False)
                IDS.Devices.DIO.SetOutputBit(1, I, False)
            Next

            'Initialize CAN IO 
            Dim LocalCAN As New CIDSCAN
            CAN = LocalCAN
            For I = 16 To 31    'iterate thru the io bits
                mysleep(10)
                IDS.Devices.Motor.SetDIOs(1, I, False)
            Next
        End If

        '   Xue Wen                                                     '
        '   Uncheck "RadioIO".                                          '
        '   If not, user only can use it after pressing other button.  
        TimerIO.Enabled = False
        TimerIO.Stop()

    End Sub
    '''''''''
    ' button view by is clicked
    '''''''''   
    Private Sub BtnViewby_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call RefreshDisplay()
    End Sub

    '''''''''
    ' validate the IO number format 
    ''''''''' 
    Private Function ValidateIONumber(ByRef Mystr As String, ByRef Card As String, ByRef IOPort As String, ByRef IOBit As Integer, ByRef Message As String) As Boolean

        Try
            Dim Str As String() = Mystr.Split("-")  'split the string of IO number to
            Card = Str(0).ToUpper                   'card type
            IOPort = Str(1).ToUpper                 'port number
            IOBit = Str(2).ToUpper                  'bit number

        Catch ex As Exception
            'prompt error if the string format is in correct 
            Message = "Can't Toggle invalid IO Number - " + Mystr.ToUpper
            Return False    'exit
        End Try

        '''''''''
        ' Trio controller IO number
        ''''''''' 
        If Card = "MC" Then

            'prompt error if the port is in correct 
            If IOPort <> "IN" And IOPort <> "OUT" Then
                Message = "PORT number is not valid - " + IOPort.ToUpper
                Return False
            Else
                'prompt error if the bitnumber is in correct 
                If IOBit <= 19 And IOBit >= 0 Then
                Else
                    Message = "BIT number is not valid - " + IOBit.ToString
                    Return False
                End If
            End If

            '''''''''
            ' CAN IO number
            ''''''''' 
        ElseIf Card = "CAN" Then
            'prompt error if the port is in correct 
            If IOPort <> "I/O" Then
                Message = "PORT number is not valid - " + IOPort.ToUpper
                Return False
            Else
                'prompt error if the bitnumber is in correct 
                If IOBit <= 31 And IOBit >= 0 Then
                Else
                    Message = "BIT number is not valid - " + IOBit.ToString

                    Return False
                End If

            End If

            '''''''''
            ' PC IO number
            ''''''''' 
        ElseIf Card = "PC" Then

            'prompt error if the port is in correct 
            If IOPort <> "IN" And IOPort <> "OUT" Then
                Message = "PORT number is not valid - " + IOPort.ToString
                Return False
            Else
                'prompt error if the bitnumber is in correct 
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

        Message = ""    'clear message
        Return True     'format is correct

    End Function
    '''''''''
    ' extract io number information
    ''''''''' 
    Private Function ExtractIONumberInfo(ByRef MyStr As String, ByRef Card As String, ByRef IOPort As String, ByRef IOBit As String) As Boolean
        Try
            Dim Str As String() = MyStr.Split("-")
            Card = Str(0).ToUpper       'card
            IOPort = Str(1).ToUpper     'port number
            IOBit = Str(2).ToUpper      'io bit number

        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function

    '''''''''
    ' toggle bit on or off
    ''''''''' 
    Private Sub ToggleBitOnOff(ByRef BitStatus As Boolean)


        '''''''''
        ' declare working variables
        ''''''''' 
        Dim Card As String
        Dim IOPort As String
        Dim IOBit As String
        Dim Message As String
        Dim BitFlag As String

        ''''''''' 
        ' if is io type = output ,
        ' validate the io number first
        ' follow by set/reset bit according to the bit status
        ''''''''' 
        If IDSData.Hardware.SystemIO.Template.Type.ToUpper = "OUTPUT" Then

            If ValidateIONumber(IDSData.Hardware.SystemIO.Template.IO, Card, IOPort, IOBit, Message) = True Then

                ''''''''' 
                ' change the true/false to ON/OFF
                ''''''''' 
                If BitStatus = True Then
                    BitFlag = "ON"
                Else
                    BitFlag = "OFF"
                End If

                ''''''''' 
                ' PC io card
                ''''''''' 
                If Card = "PC" Then

                    If IOPort = "OUT" Then          'is output type

                        Dim Portnumber As Integer
                        If IOBit < 8 Then           'if iobit number is less than 8
                            Portnumber = 0          'is at port 0
                        Else
                            Portnumber = 1          'is at port 1
                            IOBit = IOBit - 8       'reformate the io bit
                        End If

                        ''''''''' 
                        ' set/rest IO
                        ''''''''' 
                        If IDS.Devices.DIO.SetOutputBit(Portnumber, IOBit, BitStatus) Then
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag  'update on the display
                            _PC.OUT.Port(Portnumber).Bit(IOBit) = BitStatus                 'update to the working variable
                        Else
                            MessageBox.Show("PC-IO toggle failed")                          'return false, prompt err
                        End If

                    End If


                    ''''''''' 
                    ' trio controller io
                    ''''''''' 
                ElseIf Card = "MC" Then

                    If IOPort = "OUT" Then          'is output type

                        ''''''''' 
                        ' set/rest IO
                        ''''''''' 
                        'If IDS.Devices.Motor.SetDIOs(0, IOBit, BitStatus) Then
                        'DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag  'update to the display
                        'MC.ExtBit(IOBit) = BitStatus                                    'update to the working variable
                        'Else                                                                'return false,
                        'MessageBox.Show("MC-IO toggle failed")                          'prompt err
                        'DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = "Failed" 'update to the display
                        'MC.ExtBit(IOBit) = False   
                        'update to the working variable
                        'End If
                        MessageBox.Show("MC-IO does not used for output!")

                    End If

                    ''''''''' 
                    ' Can IO
                    ''''''''' 
                ElseIf Card = "CAN" Then

                    If IOPort = "I/O" Then            'can be input or output
                        ''''''''' 
                        ' set/rest IO
                        ''''''''' 
                        If IDS.Devices.Motor.SetDIOs(1, IOBit, BitStatus) Then              '
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = BitFlag  'update to the display
                            CAN.Bit(IOBit) = BitStatus                                      'update to the working variable
                        Else                                                                'return false,
                            MessageBox.Show("CAN-IO toggle failed")                         'prompt err
                            DataGrid1.Item(DataGrid1.CurrentCell.RowNumber(), 3) = "Failed" 'update to the display
                            CAN.Bit(IOBit) = False                                          'update to the working variable
                        End If
                    End If
                End If
            Else
                MessageBox.Show(Message)
            End If

        End If
    End Sub

    ''''''''' 
    ' toggle ON /OFF
    ''''''''' 
    Private Sub RD_On_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RD_On.CheckedChanged

        ''''
        '   Xue Wen                                                                             '
        '   Not to do anything if "Module Name" is not "Control I/O", "System" and "ALL".       '
        ''''
        If (CBViewType.Text = "Output") Then
            If (Trim(CBViewModule.Text) <> "Control I/O") And (Trim(CBViewModule.Text) <> "System") And (Trim(CBViewModule.Text) <> "ALL") Then
                MsgBox("Cannot do ON/OFF in this Module.")
                Exit Sub
            End If
        End If

        IOReadStop = True                   'stop the io reading flag
        mysleep(100)                          'delay
        Call ToggleBitOnOff(RD_On.Checked)  'call to the toggle function
        IOReadStop = False                  'start the io reading flag

    End Sub

    ''''''''' 
    ' IO reading timer
    ''''''''' 
    Private Sub TimerIO_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerIO.Tick

        Dim IOBitMap As Integer = 255


        ''''''''' ''''''
        ' read the IO bit status from the PC IO card, CAN, trio controller
        ''''''''' ''''''
        If IOReadStop = True Then       'io read stop flag 
            TraceDoEvents()      'is true, just do background events only
        ElseIf IOReadClose = True Then  'io read stop flag
            TimerIO.Enabled = False     'is true, stopped the timer and exit from the function
            TimerIO.Stop()
            Return
        Else                            'else
            TimerIO.Stop()              'read IO status from the hardwares
            ReadIOFlag = True           'timer functin is reading IO now flag


            ''''''''' 
            ' Declare working variables
            ''''''''' 
            Dim Card As String
            Dim IOPort As String
            Dim IOBit As String
            Dim Message As String
            Dim PortNumber As Integer
            Dim I As Integer = 0
            Dim a As Integer = 0


            ''''''''' 
            ' get can io bits
            ''''''''' 
            While Not IDS.Devices.Motor.GetDIOs(1, 16, 31, IOBitMap)
                mysleep(1000)
                TraceDoEvents()
                a = a + 1
                If a > 10 Then
                    MessageBox.Show("Can-IO Read Err")
                    Exit While
                End If
            End While
            If a <= 10 Then

                ''''''''' 
                ' convert the bit map to integer
                ''''''''' 
                Call ConvertBitMaptoInteger(IOBitMap, "CAN", "", 0)
            End If


            ''''''''' 
            ' get pc IO
            ''''''''' 
            Static Dim failcount As Integer
            Dim ReadByte As Byte
            PortNumber = 0
            If IDS.Devices.DIO.CheckInputByte(PortNumber, ReadByte) Then
                IOBitMap = CInt(ReadByte)
                ''''''''' 
                ' convert the bit map to integer
                ''''''''' 
                Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 0) 'read input port 0 
            Else
                If failcount < 4 Then
                    MessageBox.Show("PC-IO Read error")
                    failcount += 1
                End If

            End If


            PortNumber = 1
            If IDS.Devices.DIO.CheckInputByte(PortNumber, ReadByte) Then
                IOBitMap = CInt(ReadByte)
                ''''''''' 
                ' convert the bit map to integer
                ''''''''' 
                Call ConvertBitMaptoInteger(IOBitMap, "PC", "IN", 1) 'read input port 1
            Else

                If failcount < 4 Then
                    MessageBox.Show("PC-IO error")
                    failcount += 1
                End If

            End If

            ''''''''' 
            ' Get trio controller IO bits
            ''''''''' 
            a = 0
            While Not IDS.Devices.Motor.GetDIOs(0, 0, 15, IOBitMap)
                mysleep(1000)
                TraceDoEvents()
                a = a + 1
                If a > 10 Then
                    MessageBox.Show("MC-IO Read Err")
                    Exit While
                End If
            End While
            If a <= 10 Then
                ''''''''' 
                ' convert the bit map to integer
                ''''''''' 
                Call ConvertBitMaptoInteger(IOBitMap, "MC", "", 0)
            End If


            ''''''''' 
            ''''''''' 
            '
            '   update the Io status to the datagrid display
            '
            ''''''''' 
            ''''''''' 


            If DBview.Count = 0 Then
            Else
                I = 0
                While True

                    Dim Row As DataRow = DBview(I).Row

                    ''''''''' 
                    ' terminate the display if invalid type
                    ''''''''' 
                    If IsDBNull(Row("Type")) Then
                        GoTo IOList_End
                    End If

                    Dim IOType As String = Row("Type")
                    If IOType.ToUpper <> "INPUT" And IOType.ToUpper <> "OUTPUT" Then
                    Else
                        ''''''''' 
                        ' if type is input or output
                        ' update the value 
                        ' to the datagrid display
                        '''''''' 
                        Dim IONumber As String = Row("IO")
                        If ValidateIONumber(IONumber, Card, IOPort, IOBit, Message) = True Then


                            ''''''''' 
                            ' PC IO card
                            '''''''' 
                            If Card = "PC" Then
                                If IOBit < 8 Then       'IO bit is more than 8
                                    PortNumber = 0      'port 0
                                Else
                                    PortNumber = 1      'else, is port 1
                                    IOBit = IOBit - 8   'reformat the io bit
                                End If

                                ''''''''' 
                                ' Input type, update on/off to the datarow
                                '''''''' 
                                If IOPort = "IN" Then

                                    If _PC.INP.Port(PortNumber).Bit(IOBit) = True Then
                                        Row("Status") = "ON"
                                    ElseIf _PC.INP.Port(PortNumber).Bit(IOBit) = False Then
                                        Row("Status") = "OFF"
                                    End If

                                    ''''''''' 
                                    ' output type, update to the data row
                                    '''''''' 
                                ElseIf IOPort = "OUT" Then
                                    If _PC.OUT.Port(PortNumber).Bit(IOBit) = True Then
                                        Row("Status") = "ON"
                                    ElseIf _PC.OUT.Port(PortNumber).Bit(IOBit) = False Then
                                        Row("Status") = "OFF"
                                    End If

                                End If



                                ''''''''' 
                                ' Trio controller IO 
                                '''''''' 
                            ElseIf Card = "MC" Then
                                'update to the data row
                                If MC.Bit(IOBit) = True Then
                                    Row("Status") = "ON"
                                ElseIf MC.Bit(IOBit) = False Then
                                    Row("Status") = "OFF"
                                End If



                                ''''''''' 
                                ' Can IO 
                                '''''''' 
                            ElseIf Card = "CAN" Then
                                ' update to the data row
                                If CAN.Bit(IOBit) = True Then
                                    Row("Status") = "ON"
                                ElseIf CAN.Bit(IOBit) = False Then
                                    Row("Status") = "OFF"
                                End If


                                ''''''''' 
                                ' else, others display nothing
                                '''''''' 
                            Else
                                Row("Status") = ""
                            End If
                        Else
                            ''''''''' 
                            ' else, others display nothing
                            '''''''' 
                            Row("Status") = ""
                        End If
                    End If

                    TraceDoEvents()  'do backgroud event and update the status to the caption of datagrid
                    DataGrid1.CaptionText = "Refreshing Record = " + CStr(I + 1) + " of " + DBview.Count.ToString

                    If IOReadStop = True Then       'exit from loop if read stop flag is on
                        Exit While
                    End If

                    If I < DBview.Count - 1 Then    'iterate thru rows of the datagrid table
                        I += 1
                    Else
                        Exit While
                    End If
                End While
            End If

IOList_End:  'end of update io status to datagrid

            ReadIOFlag = False                      'timer function has stopped reading IO hardware
            TimerIO.Start()                         'start timer again
        End If
    End Sub
    ''''''''''
    ' convert the bit map to the respective bit number of the IO bit array
    ''''''''  '
    Private Sub ConvertBitMaptoInteger(ByVal BitMap As Integer, ByVal Card As String, ByVal IOtype As String, ByVal PortNumber As Integer)

        If Card = "PC" Then

            ''''''''''
            ' convert the bit map from bit 0 to 7
            ''''''''  '
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

            ''''''''''
            ' convert the bit map from 0 to 16 (basic type)
            ''''''''  '
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
            'If (BitMap And &H20000) Then MC.Bit(17) = True Else MC.Bit(17) = False
            'If (BitMap And &H40000) Then MC.Bit(18) = True Else MC.Bit(18) = False
            'If (BitMap And &H80000) Then MC.Bit(19) = True Else MC.Bit(19) = False

            ''''''''''
            ' convert the bit map from 8 to 17 (extension type)
            ''''''''  '
            'If (BitMap And &H1) Then MC.ExtBit(8) = True Else MC.ExtBit(8) = False
            'If (BitMap And &H2) Then MC.ExtBit(9) = True Else MC.ExtBit(9) = False
            'If (BitMap And &H4) Then MC.ExtBit(10) = True Else MC.ExtBit(10) = False
            'If (BitMap And &H8) Then MC.ExtBit(11) = True Else MC.ExtBit(11) = False

            'If (BitMap And &H10) Then MC.ExtBit(12) = True Else MC.ExtBit(12) = False
            'If (BitMap And &H20) Then MC.ExtBit(13) = True Else MC.ExtBit(13) = False
            'If (BitMap And &H40) Then MC.ExtBit(14) = True Else MC.ExtBit(14) = False
            'If (BitMap And &H80) Then MC.ExtBit(15) = True Else MC.ExtBit(15) = False

            'If (BitMap And &H100) Then MC.ExtBit(16) = True Else MC.ExtBit(16) = False
            'If (BitMap And &H200) Then MC.ExtBit(17) = True Else MC.ExtBit(17) = False


        ElseIf Card = "CAN" Then
            ''''''''''
            ' convert the bit map from 16 to 31 
            ''''''''  '
            If (BitMap And &H1) Then CAN.Bit(16) = True Else CAN.Bit(16) = False
            If (BitMap And &H2) Then CAN.Bit(17) = True Else CAN.Bit(17) = False
            If (BitMap And &H4) Then CAN.Bit(18) = True Else CAN.Bit(18) = False
            If (BitMap And &H8) Then CAN.Bit(19) = True Else CAN.Bit(19) = False

            If (BitMap And &H10) Then CAN.Bit(20) = True Else CAN.Bit(20) = False
            If (BitMap And &H20) Then CAN.Bit(21) = True Else CAN.Bit(21) = False
            If (BitMap And &H40) Then CAN.Bit(22) = True Else CAN.Bit(22) = False
            If (BitMap And &H80) Then CAN.Bit(23) = True Else CAN.Bit(23) = False

            If (BitMap And &H100) Then CAN.Bit(23) = True Else CAN.Bit(24) = False
            If (BitMap And &H200) Then CAN.Bit(25) = True Else CAN.Bit(25) = False
            If (BitMap And &H400) Then CAN.Bit(26) = True Else CAN.Bit(26) = False
            If (BitMap And &H800) Then CAN.Bit(27) = True Else CAN.Bit(27) = False

            If (BitMap And &H1000) Then CAN.Bit(28) = True Else CAN.Bit(28) = False
            If (BitMap And &H2000) Then CAN.Bit(29) = True Else CAN.Bit(29) = False
            If (BitMap And &H4000) Then CAN.Bit(30) = True Else CAN.Bit(30) = False
            If (BitMap And &H8000) Then CAN.Bit(31) = True Else CAN.Bit(31) = False

        End If

    End Sub
    ''''''''''
    ' declare the flags for IO reading status
    ''''''''  '
    Dim FilterString As String
    Friend ReadIOFlag As Boolean
    Friend IOReadStop As Boolean
    Friend IOReadClose As Boolean

    ''''''''''
    ' class for the PC card IO
    ''''''''  '
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

    ''''''''''
    ' class for the Trio controller IO
    ''''''''  '
    Public MC As New CIDSMC
    Public Class CIDSMC
        Public Bit() As Boolean = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Public ExtBit() As Boolean = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
    End Class
    ''''''''''
    ' class for the Can IO
    ''''''''  '
    Public CAN As New CIDSCAN
    Public Class CIDSCAN
        Public Bit() As Boolean = _
        {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
         0, 0, 0, 0, 0, 0, 0, 0}

    End Class

#Region " Not Used "


    '''''''''
    'not used
    '''''''''
    Private Sub CB_Type_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Type.SelectedIndexChanged
        Return
    End Sub
    '''''''''
    'to read text from file - not used
    '''''''''    
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

    '''''''''
    'to write text to file - not used
    '''''''''    
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


    Public PC As New CIDSIOCard(1)      'not used
    Public Class CIDSIOCard             'not used
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

#End Region


End Class
