Imports DLL_DataManager
Imports DLL_Export_Service
Imports Microsoft.DirectX.DirectInput
'Imports DLL_SetupAndCalibrate
Imports System.Threading
'Imports Microsoft.win32
Imports Microsoft.Office.Interop
Imports System.Messaging

Public Class FormProduction

    Inherits System.Windows.Forms.Form

    'this is to set the potlife timer on or off
    Public m_PotLifeOn As Boolean = False

    'potlife time tracking flags
    Public ElapsedTimeSincePotLifeStart As Integer = 0
    Public RemainingTimeOfPotLifeStart As Date

    'autopurging time tracking flags
    Public CurrentTime As Date
    Public StartTimeOfMachineIdle As Date = Now
    Public ElapsedTimeSinceMachineIdle As Integer = 0 'mins

    'check whether the machine has been running in continuous mode 
    Public HasBeenRunning As Boolean = False
    Private Production_info(50) As String

    'constants
    Public Const tickToMinute As Double = 0.0000000016666666  '0.0000001/60.0

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'TimerMonitor.Enabled = True
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        TimerMonitor.Enabled = False
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
    Friend WithEvents TBBStart As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBPause As System.Windows.Forms.ToolBarButton
    Friend WithEvents TBBStop As System.Windows.Forms.ToolBarButton
    Friend WithEvents ImageListOper As System.Windows.Forms.ImageList
    Friend WithEvents ImageListPotEtc As System.Windows.Forms.ImageList
    Public WithEvents MainMenuProduction As System.Windows.Forms.MainMenu
    Friend WithEvents ImageListGeneralTools As System.Windows.Forms.ImageList
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBoxRobotPos As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PanelProDownTimeInfor As System.Windows.Forms.Panel
    Friend WithEvents RichTextBoxNote As System.Windows.Forms.RichTextBox
    Friend WithEvents CheckBoxPotOn As System.Windows.Forms.CheckBox
    Friend WithEvents DoorLock As System.Windows.Forms.CheckBox
    Friend WithEvents ButtonPotReset As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ButtonOpenFile As System.Windows.Forms.Button
    Friend WithEvents TextBoxFilename As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ButtonClean As System.Windows.Forms.Button
    Friend WithEvents ButtonPurge As System.Windows.Forms.Button
    Friend WithEvents ButtonChgSyringe As System.Windows.Forms.Button
    Friend WithEvents ButtonHome As System.Windows.Forms.Button
    Friend WithEvents LabelMessege As System.Windows.Forms.Label
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents TBOperation As System.Windows.Forms.ToolBar
    Friend WithEvents PBYellow As System.Windows.Forms.PictureBox
    Friend WithEvents PBRed As System.Windows.Forms.PictureBox
    Friend WithEvents PBGreen As System.Windows.Forms.PictureBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ImageListOperation As System.Windows.Forms.ImageList
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TimerMonitor As System.Windows.Forms.Timer
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents ButtonCalibrate As System.Windows.Forms.Button
    Friend WithEvents ContinuousMode As System.Windows.Forms.CheckBox
    Public WithEvents PanelToBeAdded As System.Windows.Forms.Panel
    Friend WithEvents ComboBoxFineStep As System.Windows.Forms.NumericUpDown
    Friend WithEvents ButtonStepZdown As System.Windows.Forms.Button
    Friend WithEvents ButtonStepZup As System.Windows.Forms.Button
    Friend WithEvents ButtonStepXminus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepXplus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepYminus As System.Windows.Forms.Button
    Friend WithEvents ButtonStepYplus As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents LabelStep As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormProduction))
        Me.TBBStart = New System.Windows.Forms.ToolBarButton
        Me.TBBPause = New System.Windows.Forms.ToolBarButton
        Me.TBBStop = New System.Windows.Forms.ToolBarButton
        Me.ImageListOper = New System.Windows.Forms.ImageList(Me.components)
        Me.ImageListPotEtc = New System.Windows.Forms.ImageList(Me.components)
        Me.MainMenuProduction = New System.Windows.Forms.MainMenu
        Me.ImageListGeneralTools = New System.Windows.Forms.ImageList(Me.components)
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TextBoxRobotPos = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.PanelToBeAdded = New System.Windows.Forms.Panel
        Me.ComboBoxFineStep = New System.Windows.Forms.NumericUpDown
        Me.ButtonStepZdown = New System.Windows.Forms.Button
        Me.ButtonStepZup = New System.Windows.Forms.Button
        Me.ButtonStepXminus = New System.Windows.Forms.Button
        Me.ButtonStepXplus = New System.Windows.Forms.Button
        Me.ButtonStepYminus = New System.Windows.Forms.Button
        Me.ButtonStepYplus = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.LabelStep = New System.Windows.Forms.Label
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.TBOperation = New System.Windows.Forms.ToolBar
        Me.PBRed = New System.Windows.Forms.PictureBox
        Me.PBGreen = New System.Windows.Forms.PictureBox
        Me.PBYellow = New System.Windows.Forms.PictureBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.Button4 = New System.Windows.Forms.Button
        Me.ContinuousMode = New System.Windows.Forms.CheckBox
        Me.PanelProDownTimeInfor = New System.Windows.Forms.Panel
        Me.RichTextBoxNote = New System.Windows.Forms.RichTextBox
        Me.CheckBoxPotOn = New System.Windows.Forms.CheckBox
        Me.DoorLock = New System.Windows.Forms.CheckBox
        Me.ButtonPotReset = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.ButtonOpenFile = New System.Windows.Forms.Button
        Me.TextBoxFilename = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.ButtonClean = New System.Windows.Forms.Button
        Me.ButtonPurge = New System.Windows.Forms.Button
        Me.ButtonChgSyringe = New System.Windows.Forms.Button
        Me.ButtonHome = New System.Windows.Forms.Button
        Me.ButtonCalibrate = New System.Windows.Forms.Button
        Me.LabelMessege = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ImageListOperation = New System.Windows.Forms.ImageList(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TimerMonitor = New System.Windows.Forms.Timer(Me.components)
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.Panel2.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.PanelToBeAdded.SuspendLayout()
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel6.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.PanelProDownTimeInfor.SuspendLayout()
        Me.SuspendLayout()
        '
        'TBBStart
        '
        Me.TBBStart.Enabled = False
        Me.TBBStart.ImageIndex = 0
        '
        'TBBPause
        '
        Me.TBBPause.Enabled = False
        Me.TBBPause.ImageIndex = 1
        '
        'TBBStop
        '
        Me.TBBStop.Enabled = False
        Me.TBBStop.ImageIndex = 2
        '
        'ImageListOper
        '
        Me.ImageListOper.ImageSize = New System.Drawing.Size(48, 48)
        Me.ImageListOper.ImageStream = CType(resources.GetObject("ImageListOper.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOper.TransparentColor = System.Drawing.Color.White
        '
        'ImageListPotEtc
        '
        Me.ImageListPotEtc.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListPotEtc.ImageStream = CType(resources.GetObject("ImageListPotEtc.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListPotEtc.TransparentColor = System.Drawing.Color.Transparent
        '
        'ImageListGeneralTools
        '
        Me.ImageListGeneralTools.ImageSize = New System.Drawing.Size(36, 28)
        Me.ImageListGeneralTools.ImageStream = CType(resources.GetObject("ImageListGeneralTools.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListGeneralTools.TransparentColor = System.Drawing.Color.White
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.TextBoxRobotPos)
        Me.Panel2.Controls.Add(Me.Label4)
        Me.Panel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Panel2.Location = New System.Drawing.Point(0, 960)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(768, 32)
        Me.Panel2.TabIndex = 4
        '
        'TextBoxRobotPos
        '
        Me.TextBoxRobotPos.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBoxRobotPos.Location = New System.Drawing.Point(504, 3)
        Me.TextBoxRobotPos.Name = "TextBoxRobotPos"
        Me.TextBoxRobotPos.ReadOnly = True
        Me.TextBoxRobotPos.Size = New System.Drawing.Size(256, 21)
        Me.TextBoxRobotPos.TabIndex = 8
        Me.TextBoxRobotPos.Text = "X: -100.000,   Y: -100.000,  Z: -100.000"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(450, 5)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(45, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Robot"
        '
        'Panel5
        '
        Me.Panel5.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.Panel5.Controls.Add(Me.Label2)
        Me.Panel5.Controls.Add(Me.PanelToBeAdded)
        Me.Panel5.Controls.Add(Me.Panel6)
        Me.Panel5.Controls.Add(Me.TextBox1)
        Me.Panel5.Controls.Add(Me.TextBox2)
        Me.Panel5.Controls.Add(Me.GroupBox1)
        Me.Panel5.Controls.Add(Me.ContinuousMode)
        Me.Panel5.ForeColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        Me.Panel5.Location = New System.Drawing.Point(768, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(504, 992)
        Me.Panel5.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(88, 600)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 23)
        Me.Label2.TabIndex = 135
        Me.Label2.Text = "Machine State :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PanelToBeAdded
        '
        Me.PanelToBeAdded.BackColor = System.Drawing.Color.LightSteelBlue
        Me.PanelToBeAdded.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PanelToBeAdded.Controls.Add(Me.ComboBoxFineStep)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepZdown)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepZup)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepXminus)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepXplus)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepYminus)
        Me.PanelToBeAdded.Controls.Add(Me.ButtonStepYplus)
        Me.PanelToBeAdded.Controls.Add(Me.Label1)
        Me.PanelToBeAdded.Controls.Add(Me.LabelStep)
        Me.PanelToBeAdded.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PanelToBeAdded.Location = New System.Drawing.Point(84, 256)
        Me.PanelToBeAdded.Name = "PanelToBeAdded"
        Me.PanelToBeAdded.Size = New System.Drawing.Size(336, 280)
        Me.PanelToBeAdded.TabIndex = 134
        Me.PanelToBeAdded.Visible = False
        '
        'ComboBoxFineStep
        '
        Me.ComboBoxFineStep.DecimalPlaces = 3
        Me.ComboBoxFineStep.Location = New System.Drawing.Point(152, 240)
        Me.ComboBoxFineStep.Maximum = New Decimal(New Integer() {200, 0, 0, 0})
        Me.ComboBoxFineStep.Name = "ComboBoxFineStep"
        Me.ComboBoxFineStep.Size = New System.Drawing.Size(88, 27)
        Me.ComboBoxFineStep.TabIndex = 9
        '
        'ButtonStepZdown
        '
        Me.ButtonStepZdown.BackColor = System.Drawing.Color.Linen
        Me.ButtonStepZdown.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepZdown.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonStepZdown.Location = New System.Drawing.Point(264, 128)
        Me.ButtonStepZdown.Name = "ButtonStepZdown"
        Me.ButtonStepZdown.Size = New System.Drawing.Size(48, 72)
        Me.ButtonStepZdown.TabIndex = 8
        Me.ButtonStepZdown.Text = "Dn"
        Me.ButtonStepZdown.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ButtonStepZup
        '
        Me.ButtonStepZup.BackColor = System.Drawing.Color.Linen
        Me.ButtonStepZup.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepZup.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonStepZup.Location = New System.Drawing.Point(264, 40)
        Me.ButtonStepZup.Name = "ButtonStepZup"
        Me.ButtonStepZup.Size = New System.Drawing.Size(48, 72)
        Me.ButtonStepZup.TabIndex = 7
        Me.ButtonStepZup.Text = "Up"
        Me.ButtonStepZup.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonStepXminus
        '
        Me.ButtonStepXminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.ButtonStepXminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepXminus.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ButtonStepXminus.Location = New System.Drawing.Point(24, 96)
        Me.ButtonStepXminus.Name = "ButtonStepXminus"
        Me.ButtonStepXminus.Size = New System.Drawing.Size(80, 48)
        Me.ButtonStepXminus.TabIndex = 6
        Me.ButtonStepXminus.Text = "X-"
        Me.ButtonStepXminus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ButtonStepXplus
        '
        Me.ButtonStepXplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.ButtonStepXplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepXplus.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ButtonStepXplus.Location = New System.Drawing.Point(152, 96)
        Me.ButtonStepXplus.Name = "ButtonStepXplus"
        Me.ButtonStepXplus.Size = New System.Drawing.Size(80, 48)
        Me.ButtonStepXplus.TabIndex = 5
        Me.ButtonStepXplus.Text = "X+"
        Me.ButtonStepXplus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonStepYminus
        '
        Me.ButtonStepYminus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.ButtonStepYminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepYminus.ImageAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ButtonStepYminus.Location = New System.Drawing.Point(104, 144)
        Me.ButtonStepYminus.Name = "ButtonStepYminus"
        Me.ButtonStepYminus.Size = New System.Drawing.Size(48, 80)
        Me.ButtonStepYminus.TabIndex = 4
        Me.ButtonStepYminus.Text = "Y-"
        Me.ButtonStepYminus.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ButtonStepYplus
        '
        Me.ButtonStepYplus.BackColor = System.Drawing.Color.LightGoldenrodYellow
        Me.ButtonStepYplus.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonStepYplus.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonStepYplus.Location = New System.Drawing.Point(104, 16)
        Me.ButtonStepYplus.Name = "ButtonStepYplus"
        Me.ButtonStepYplus.Size = New System.Drawing.Size(48, 80)
        Me.ButtonStepYplus.TabIndex = 3
        Me.ButtonStepYplus.Text = "Y+"
        Me.ButtonStepYplus.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.LightSteelBlue
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(248, 240)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "mm"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelStep
        '
        Me.LabelStep.BackColor = System.Drawing.Color.LightSteelBlue
        Me.LabelStep.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.LabelStep.Location = New System.Drawing.Point(56, 240)
        Me.LabelStep.Name = "LabelStep"
        Me.LabelStep.Size = New System.Drawing.Size(88, 24)
        Me.LabelStep.TabIndex = 0
        Me.LabelStep.Text = "Fine Step:"
        Me.LabelStep.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel6
        '
        Me.Panel6.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.Panel6.Controls.Add(Me.TBOperation)
        Me.Panel6.Controls.Add(Me.PBRed)
        Me.Panel6.Controls.Add(Me.PBGreen)
        Me.Panel6.Controls.Add(Me.PBYellow)
        Me.Panel6.Location = New System.Drawing.Point(100, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(308, 216)
        Me.Panel6.TabIndex = 0
        '
        'TBOperation
        '
        Me.TBOperation.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.TBOperation.BackColor = System.Drawing.SystemColors.Menu
        Me.TBOperation.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.TBOperation.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.TBBStart, Me.TBBPause, Me.TBBStop})
        Me.TBOperation.ButtonSize = New System.Drawing.Size(90, 60)
        Me.TBOperation.Dock = System.Windows.Forms.DockStyle.None
        Me.TBOperation.DropDownArrows = True
        Me.TBOperation.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TBOperation.ForeColor = System.Drawing.Color.Black
        Me.TBOperation.ImageList = Me.ImageListOper
        Me.TBOperation.Location = New System.Drawing.Point(18, 112)
        Me.TBOperation.Name = "TBOperation"
        Me.TBOperation.ShowToolTips = True
        Me.TBOperation.Size = New System.Drawing.Size(274, 68)
        Me.TBOperation.TabIndex = 128
        '
        'PBRed
        '
        Me.PBRed.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBRed.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBRed.Image = CType(resources.GetObject("PBRed.Image"), System.Drawing.Image)
        Me.PBRed.Location = New System.Drawing.Point(13, 11)
        Me.PBRed.Name = "PBRed"
        Me.PBRed.Size = New System.Drawing.Size(279, 75)
        Me.PBRed.TabIndex = 129
        Me.PBRed.TabStop = False
        '
        'PBGreen
        '
        Me.PBGreen.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBGreen.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBGreen.Image = CType(resources.GetObject("PBGreen.Image"), System.Drawing.Image)
        Me.PBGreen.Location = New System.Drawing.Point(13, 11)
        Me.PBGreen.Name = "PBGreen"
        Me.PBGreen.Size = New System.Drawing.Size(279, 75)
        Me.PBGreen.TabIndex = 130
        Me.PBGreen.TabStop = False
        '
        'PBYellow
        '
        Me.PBYellow.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PBYellow.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.PBYellow.Image = CType(resources.GetObject("PBYellow.Image"), System.Drawing.Image)
        Me.PBYellow.Location = New System.Drawing.Point(13, 11)
        Me.PBYellow.Name = "PBYellow"
        Me.PBYellow.Size = New System.Drawing.Size(279, 75)
        Me.PBYellow.TabIndex = 131
        Me.PBYellow.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(24, 952)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(184, 27)
        Me.TextBox1.TabIndex = 132
        Me.TextBox1.Text = "prev state"
        Me.TextBox1.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(232, 600)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(184, 27)
        Me.TextBox2.TabIndex = 132
        Me.TextBox2.Text = "current state"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.Button3)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.ComboBox3)
        Me.GroupBox1.Controls.Add(Me.Button4)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.GroupBox1.Location = New System.Drawing.Point(16, 680)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(480, 256)
        Me.GroupBox1.TabIndex = 130
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Test"
        Me.GroupBox1.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"Production Start", "Board Comes", "Board Goes", "Pause", "Resume", "Board Partial Failure", "Board Total Failure", "Production End", "Purge And Clean", "Dot Size Check Passed", "Volume Calibration Passed", "Volume Calibration Failed", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed"})
        Me.ComboBox1.Location = New System.Drawing.Point(16, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox1.TabIndex = 140
        Me.ComboBox1.Text = "ComboBox1"
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button1.Location = New System.Drawing.Point(16, 80)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(176, 40)
        Me.Button1.TabIndex = 139
        Me.Button1.Text = "enter 1 SPC event"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button2.Location = New System.Drawing.Point(16, 192)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(208, 56)
        Me.Button2.TabIndex = 139
        Me.Button2.Text = "enter all SPC events"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button3.Location = New System.Drawing.Point(16, 136)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(192, 40)
        Me.Button3.TabIndex = 139
        Me.Button3.Text = "generate SPC report"
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ComboBox2
        '
        Me.ComboBox2.Items.AddRange(New Object() {"Production Start", "Board Comes", "Board Goes", "Pause", "Resume", "Board Partial Failure", "Board Total Failure", "Production End", "Purge And Clean", "Dot Size Check Passed", "Volume Calibration Passed", "Volume Calibration Failed", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed"})
        Me.ComboBox2.Location = New System.Drawing.Point(16, 32)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox2.TabIndex = 140
        Me.ComboBox2.Text = "ComboBox1"
        '
        'ComboBox3
        '
        Me.ComboBox3.Items.AddRange(New Object() {"Width Adjustment Failed", "Lifter 1 Blocked", "Lifter 2 Blocked", "Lifter 3 Blocked", "Retrieve Timeout", "Stage 1 Travel Timeout", "Release Travel Timeout", "Stage 3 Travel Timeout", "Board Not Aligned", "PLC Communication Error", "PLC No Signal", "Low Incoming Air Pressure", "Fiducial Failed", "Height Point Failed", "Chip Finding Failed", "Dot Size Check Failed", "E-Stop", "Camera Signal Failed", "Volume Calibration Failed"})
        Me.ComboBox3.Location = New System.Drawing.Point(288, 32)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(176, 28)
        Me.ComboBox3.TabIndex = 140
        Me.ComboBox3.Text = "ComboBox1"
        '
        'Button4
        '
        Me.Button4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button4.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.Button4.Location = New System.Drawing.Point(288, 80)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(176, 40)
        Me.Button4.TabIndex = 139
        Me.Button4.Text = "generate popup"
        Me.Button4.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ContinuousMode
        '
        Me.ContinuousMode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ContinuousMode.ForeColor = System.Drawing.Color.Black
        Me.ContinuousMode.Location = New System.Drawing.Point(96, 552)
        Me.ContinuousMode.Name = "ContinuousMode"
        Me.ContinuousMode.Size = New System.Drawing.Size(176, 40)
        Me.ContinuousMode.TabIndex = 133
        Me.ContinuousMode.Text = "Continuous Mode"
        '
        'PanelProDownTimeInfor
        '
        Me.PanelProDownTimeInfor.BackColor = System.Drawing.SystemColors.Info
        Me.PanelProDownTimeInfor.Controls.Add(Me.RichTextBoxNote)
        Me.PanelProDownTimeInfor.Controls.Add(Me.CheckBoxPotOn)
        Me.PanelProDownTimeInfor.Controls.Add(Me.DoorLock)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPotReset)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label6)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonOpenFile)
        Me.PanelProDownTimeInfor.Controls.Add(Me.TextBoxFilename)
        Me.PanelProDownTimeInfor.Controls.Add(Me.Label5)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonClean)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonPurge)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonChgSyringe)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonHome)
        Me.PanelProDownTimeInfor.Controls.Add(Me.ButtonCalibrate)
        Me.PanelProDownTimeInfor.Controls.Add(Me.LabelMessege)
        Me.PanelProDownTimeInfor.Location = New System.Drawing.Point(0, 0)
        Me.PanelProDownTimeInfor.Name = "PanelProDownTimeInfor"
        Me.PanelProDownTimeInfor.Size = New System.Drawing.Size(768, 960)
        Me.PanelProDownTimeInfor.TabIndex = 6
        '
        'RichTextBoxNote
        '
        Me.RichTextBoxNote.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.RichTextBoxNote.Location = New System.Drawing.Point(16, 184)
        Me.RichTextBoxNote.Name = "RichTextBoxNote"
        Me.RichTextBoxNote.ReadOnly = True
        Me.RichTextBoxNote.Size = New System.Drawing.Size(728, 720)
        Me.RichTextBoxNote.TabIndex = 98
        Me.RichTextBoxNote.Text = ""
        '
        'CheckBoxPotOn
        '
        Me.CheckBoxPotOn.Appearance = System.Windows.Forms.Appearance.Button
        Me.CheckBoxPotOn.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.CheckBoxPotOn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBoxPotOn.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.CheckBoxPotOn.ImageIndex = 6
        Me.CheckBoxPotOn.ImageList = Me.ImageListPotEtc
        Me.CheckBoxPotOn.Location = New System.Drawing.Point(458, 16)
        Me.CheckBoxPotOn.Name = "CheckBoxPotOn"
        Me.CheckBoxPotOn.Size = New System.Drawing.Size(75, 56)
        Me.CheckBoxPotOn.TabIndex = 119
        Me.CheckBoxPotOn.Text = "Pot Life On"
        Me.CheckBoxPotOn.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'DoorLock
        '
        Me.DoorLock.Appearance = System.Windows.Forms.Appearance.Button
        Me.DoorLock.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.DoorLock.Checked = True
        Me.DoorLock.CheckState = System.Windows.Forms.CheckState.Checked
        Me.DoorLock.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.DoorLock.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.DoorLock.ImageIndex = 3
        Me.DoorLock.ImageList = Me.ImageListPotEtc
        Me.DoorLock.Location = New System.Drawing.Point(656, 16)
        Me.DoorLock.Name = "DoorLock"
        Me.DoorLock.Size = New System.Drawing.Size(75, 56)
        Me.DoorLock.TabIndex = 116
        Me.DoorLock.Text = "Lock Door"
        Me.DoorLock.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonPotReset
        '
        Me.ButtonPotReset.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonPotReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPotReset.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPotReset.ImageIndex = 1
        Me.ButtonPotReset.ImageList = Me.ImageListPotEtc
        Me.ButtonPotReset.Location = New System.Drawing.Point(383, 16)
        Me.ButtonPotReset.Name = "ButtonPotReset"
        Me.ButtonPotReset.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPotReset.TabIndex = 105
        Me.ButtonPotReset.Text = "Reset Pot"
        Me.ButtonPotReset.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 160)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 97
        Me.Label6.Text = "Note :"
        '
        'ButtonOpenFile
        '
        Me.ButtonOpenFile.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.ButtonOpenFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonOpenFile.Location = New System.Drawing.Point(672, 112)
        Me.ButtonOpenFile.Name = "ButtonOpenFile"
        Me.ButtonOpenFile.Size = New System.Drawing.Size(72, 26)
        Me.ButtonOpenFile.TabIndex = 96
        Me.ButtonOpenFile.Text = "Open"
        '
        'TextBoxFilename
        '
        Me.TextBoxFilename.BackColor = System.Drawing.Color.White
        Me.TextBoxFilename.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.TextBoxFilename.Location = New System.Drawing.Point(56, 112)
        Me.TextBoxFilename.Name = "TextBoxFilename"
        Me.TextBoxFilename.ReadOnly = True
        Me.TextBoxFilename.Size = New System.Drawing.Size(600, 27)
        Me.TextBoxFilename.TabIndex = 95
        Me.TextBoxFilename.Text = ""
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 112)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 94
        Me.Label5.Text = "File:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonClean
        '
        Me.ButtonClean.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonClean.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonClean.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonClean.ImageIndex = 5
        Me.ButtonClean.ImageList = Me.ImageListGeneralTools
        Me.ButtonClean.Location = New System.Drawing.Point(158, 16)
        Me.ButtonClean.Name = "ButtonClean"
        Me.ButtonClean.Size = New System.Drawing.Size(75, 56)
        Me.ButtonClean.TabIndex = 89
        Me.ButtonClean.Text = "Clean On"
        Me.ButtonClean.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonPurge
        '
        Me.ButtonPurge.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonPurge.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonPurge.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonPurge.ImageIndex = 4
        Me.ButtonPurge.ImageList = Me.ImageListGeneralTools
        Me.ButtonPurge.Location = New System.Drawing.Point(83, 16)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(75, 56)
        Me.ButtonPurge.TabIndex = 88
        Me.ButtonPurge.Text = "Purge On"
        Me.ButtonPurge.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonChgSyringe
        '
        Me.ButtonChgSyringe.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonChgSyringe.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonChgSyringe.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonChgSyringe.ImageIndex = 1
        Me.ButtonChgSyringe.ImageList = Me.ImageListGeneralTools
        Me.ButtonChgSyringe.Location = New System.Drawing.Point(233, 16)
        Me.ButtonChgSyringe.Name = "ButtonChgSyringe"
        Me.ButtonChgSyringe.Size = New System.Drawing.Size(75, 56)
        Me.ButtonChgSyringe.TabIndex = 87
        Me.ButtonChgSyringe.Text = "Chg. Syr."
        Me.ButtonChgSyringe.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonHome
        '
        Me.ButtonHome.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonHome.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonHome.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonHome.ImageIndex = 0
        Me.ButtonHome.ImageList = Me.ImageListGeneralTools
        Me.ButtonHome.Location = New System.Drawing.Point(8, 16)
        Me.ButtonHome.Name = "ButtonHome"
        Me.ButtonHome.Size = New System.Drawing.Size(75, 56)
        Me.ButtonHome.TabIndex = 86
        Me.ButtonHome.Text = "Home"
        Me.ButtonHome.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'ButtonCalibrate
        '
        Me.ButtonCalibrate.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.ButtonCalibrate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.ButtonCalibrate.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ButtonCalibrate.ImageIndex = 5
        Me.ButtonCalibrate.ImageList = Me.ImageListGeneralTools
        Me.ButtonCalibrate.Location = New System.Drawing.Point(308, 16)
        Me.ButtonCalibrate.Name = "ButtonCalibrate"
        Me.ButtonCalibrate.Size = New System.Drawing.Size(75, 56)
        Me.ButtonCalibrate.TabIndex = 88
        Me.ButtonCalibrate.Text = "Move Calibrate"
        Me.ButtonCalibrate.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'LabelMessege
        '
        Me.LabelMessege.BackColor = System.Drawing.SystemColors.Menu
        Me.LabelMessege.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LabelMessege.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelMessege.ForeColor = System.Drawing.Color.Blue
        Me.LabelMessege.Location = New System.Drawing.Point(8, 920)
        Me.LabelMessege.Name = "LabelMessege"
        Me.LabelMessege.Size = New System.Drawing.Size(750, 32)
        Me.LabelMessege.TabIndex = 85
        Me.LabelMessege.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ImageListOperation
        '
        Me.ImageListOperation.ImageSize = New System.Drawing.Size(32, 32)
        Me.ImageListOperation.ImageStream = CType(resources.GetObject("ImageListOperation.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListOperation.TransparentColor = System.Drawing.Color.Transparent
        '
        'TimerMonitor
        '
        Me.TimerMonitor.Interval = 50
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.SlateGray
        Me.PanelVision.Location = New System.Drawing.Point(0, 384)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(760, 32)
        Me.PanelVision.TabIndex = 7
        Me.PanelVision.Visible = False
        '
        'FormProduction
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1276, 990)
        Me.Controls.Add(Me.PanelVision)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.PanelProDownTimeInfor)
        Me.Controls.Add(Me.Panel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Menu = Me.MainMenuProduction
        Me.Name = "FormProduction"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Production"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel2.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.PanelToBeAdded.ResumeLayout(False)
        CType(Me.ComboBoxFineStep, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel6.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.PanelProDownTimeInfor.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'reset!
        ResetToIdle()

        'gui visibility
        Panel5.Controls.Add(m_Tri.SteppingButtons)
        m_Tri.SteppingButtons.Location = New Point(84, 192)
        m_Tri.SteppingButtons.BringToFront()
        m_Tri.SteppingButtons.Show()

        'initialize private flags
        HasBeenRunning = False
        m_PotLifeOn = False
        TextBoxFilename.Text = " "
        ProductionInfoDispClear()

        m_Execution.m_Pattern.SubCallSheetInitialization(200)     '200 subsheets maximum without duplicated name
        Programming.ErrorSubSheetStructIni(200, 500)    '500 subsheets maximum with duplicated name

        ' get default value from the default pat file
        IDS.Data.ParameterID.RecordID = "FactoryDefault"
        IDSData.Admin.Folder.FileExtension = "Pat"
        IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        SystemSetupDataRetrieve(SettingsMode.GlobalSettings)
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()

        'if jetting valve or auger valve turn on material air
        With IDS.Data.Hardware.Dispenser.Left
            If .HeadType = "Jetting Valve" Or .HeadType = "Auger Valve" Then
                m_Tri.TurnOn("Material Air")
            End If
        End With

        'error handling message
        Form_Service.ResetEventCode()

        'hardware

        'vision

        'motion controller
        m_Tri.Connect_Controller()
        SetState("Homing")

        'timers start
        IDS.StartErrorCheck()
        TimerMonitor.Start()
        Programming.IOCheck.Start()

        ' background threads help to update UI without slowing them down.
        ThreadMonitor = New Threading.Thread(AddressOf StateMonitor)
        ThreadMonitor.Priority = Threading.ThreadPriority.Normal
        'ThreadMonitor.IsBackground = True
        ThreadMonitor.Start()

        ThreadExecutor = New Threading.Thread(AddressOf StateChanger)
        ThreadExecutor.Priority = Threading.ThreadPriority.Normal
        ThreadExecutor.Start()

        MouseTimer = New System.Threading.Timer(AddressOf MouseJogging, Nothing, 0, 200)

    End Sub

    Private isJogON As Boolean = False
    Shared mouse_pos As New Point
    Shared cursor_hide As Boolean = False
    Shared isPress As Boolean
    Dim deadzone As Integer = 3
    Dim jogspeed As Integer = 0
    Const maxSpeed As Integer = 100
    Const maxMouseRangeX = 600.0 '6000.0 '4500.0
    Const maxMouseRangeY = 300.0 '2000.0
    Const ratioLB = 0.4
    Const ratioUB = 1.05
    Dim ratio As Double
    Dim countMouseTimer As Integer = 0
    Private m_TrackBall As New DLL_Export_Device_Motor.Mouse(Me)
    Private m_keyBoard As New DLL_Export_Device_Motor.Keyboard(Me)

    Private Sub MouseJogging(ByVal state As Object)

        If Production.ButtonCalibrate.Text = "Set Calibrate" Then
        Else
            If IsBusy() And Not IsJogging() Then Exit Sub
            If m_Tri.MachineHoming Or m_Tri.MachineRunning Or m_Tri.Stepping Then Exit Sub
        End If

        m_keyBoard.Poll()
        m_TrackBall.Poll()
        isPress = m_keyBoard.State.Item(Key.LeftControl)

        Dim x As Integer
        Dim y As Integer
        Dim VrData(3) As Single

        If isPress Then

            SetState("Jogging")

            VrData(0) = 0
            VrData(1) = 0.0
            VrData(2) = 0.0

            Dim isPressAlt As Boolean = m_keyBoard.State.Item(Key.LeftAlt)
            If isPressAlt Then
                Exit Sub
            End If
            x = m_TrackBall.MouseX()
            y = m_TrackBall.MouseY()

            Dim ratio As Double

            If Math.Abs(x) >= Math.Abs(y) Then
                If x > deadzone Then
                    jogspeed = CInt(Math.Abs(x) / maxMouseRangeX * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(y) / x
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X+ Y-
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = -jogspeed
                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = 0

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                ElseIf x < -deadzone Then
                    jogspeed = CInt(Math.Abs(x) / maxMouseRangeX * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(y) / x
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X- Y+
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'X-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = 0

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_Tri.SetTrioMotionValues("Jogging", VrData)
                    isJogON = True 'False
                End If
            Else
                If y < -deadzone Then
                    jogspeed = CInt(Math.Abs(y) / maxMouseRangeY * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If

                    ratio = CDbl(x) / y
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X- Y+
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = jogspeed

                        SetState("Jogging")
                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X+ Y+
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y+
                        VrData(0) = 1
                        VrData(1) = 0
                        VrData(2) = jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If

                ElseIf y > deadzone Then
                    jogspeed = CInt(Math.Abs(y) / maxMouseRangeY * maxSpeed)
                    If (jogspeed > maxSpeed) Then
                        jogspeed = maxSpeed
                    End If
                    ratio = CDbl(x) / y
                    If (ratio > ratioLB) And (ratio < ratioUB) Then   'X+ Y-
                        VrData(0) = 1
                        VrData(1) = jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    ElseIf (ratio < -ratioLB) And (ratio > -ratioUB) Then 'X- Y-
                        VrData(0) = 1
                        VrData(1) = -jogspeed
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    Else   'Y-
                        VrData(0) = 1
                        VrData(1) = 0
                        VrData(2) = -jogspeed

                        m_Tri.SetTrioMotionValues("Jogging", VrData)
                        isJogON = True
                    End If
                Else
                    VrData(0) = 2
                    VrData(1) = 0.0
                    VrData(2) = 0.0

                    m_Tri.SetTrioMotionValues("Jogging", VrData)
                    isJogON = True 'False
                End If
            End If
        Else
            If isJogON Then
                VrData(0) = 2
                VrData(1) = 0.0
                VrData(2) = 0.0

                m_Tri.SetTrioMotionValues("Jogging", VrData)
                isJogON = False
                If m_EditStateFlag Then
                    'reset to idle without the camera thing
                    SetState("Idle")
                    m_Tri.SetMachineStop()
                    SetLampsToReadyMode()
                    UnlockMovementButtons()
                    ChangeButtonState("Idle")
                Else
                    If Production.ButtonCalibrate.Text = "Set Calibrate" Then
                        Production.ButtonCalibrate.Enabled = True
                        m_Tri.SteppingButtons.Enabled = True
                    Else
                        ResetToIdle()
                    End If
                End If
            End If

            countMouseTimer += 1
            If (countMouseTimer >= 5) Then
                TraceGCCollect()
                countMouseTimer = 0
            End If
        End If

    End Sub

    Private Sub FormProduction_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        If ContinuousMode.Checked = True Then
            ContinuousMode.Checked = False
        End If


        'error handling
        Form_Service.ResetEventCode()

        'vision

        'timers stop
        IDS.StopErrorCheck()
        TimerMonitor.Stop()
        Programming.IOCheck.Stop()
        ThreadMonitor.Abort()
        ThreadExecutor.Abort()
        MouseTimer.Dispose()

        'motion controller
        m_Tri.TurnOff("Material Air")
        m_Tri.Disconnect_Controller()

        'hardware
        OffLaser()
        OffTowerLamp()
        UnlockDoor()

        Close()
        IDS.FrmExecution.Hide()

    End Sub

    Public Sub ProductionInfoDispClear()
        Dim i As Integer
        For i = 0 To 49
            Production_info(i) = ""
        Next

        RichTextBoxNote.Lines = Production_info
        RichTextBoxNote.Show()
    End Sub


    Public Sub ProductionInfoDisp()
        '  IDS.Data.OpenData()
        Production_info(0) = "Author: " + IDS.Data.Hardware.SPC.ProgAuthorName
        Production_info(1) = "Contact: " + IDS.Data.Hardware.SPC.ProgAuthorContact
        Production_info(2) = "Notes: " + IDS.Data.Hardware.SPC.ProductionNote
        Production_info(3) = "Needle: Left"
        Production_info(4) = "        Material: " & IDS.Data.Hardware.Dispenser.Left.MaterialInfo

        Production_info(7) = "        Left Needle Color "

        RichTextBoxNote.Lines = Production_info

        RichTextBoxNote.Show()
        If Not GetInputState = 0 Then Application.DoEvents()
    End Sub

    ' open file sub
    Private Sub OpenFile()

        ButtonOpenFile.Enabled = False

        If Programming.ProductionFileOpen() < 0 Then
            ButtonOpenFile.Enabled = True
            Exit Sub
        End If
        'IDS.Data.ParameterID.RecordID = ""
        'IDSData.Admin.Folder.FileExtension = "Pat"
        'IDSData.Admin.Folder.PatternPath = "C:\IDS\Pattern_Dir"
        IDS.Data.OpenData()
        If Nothing = TextBoxFilename.Text Then
            ButtonOpenFile.Enabled = True
            Return
        End If

        LabelMessage("Please wait, system is uploading..")
        Dim filename As String = m_Execution.m_File.FolderWithNameFromFileName(TextBoxFilename.Text)
        m_Execution.m_Pattern.m_ErrorChk.GetErrorCheckParameter()
        Programming.Disp_Dispenser_Unit_info()

        TextBoxFilename.Refresh()
        RichTextBoxNote.Refresh()

        ProductionInfoDisp()

        LabelMessage("Finish..")
        ButtonOpenFile.Enabled = True
    End Sub

    Private Sub ButtonOpenFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOpenFile.Click
        OpenFile()
    End Sub

    Public Sub ProdPurgeClean() 'for now, do not discriminate between left and right

        Dim cmdStr As String
        Dim needleIO, mode, purgeduration, cleanduration As Integer

        LockMovementButtons()
        LabelMessage("Purge and clean.")
        If ShouldLog() Then Form_Service.LogEventInSPCReport("Purge And Clean")
        needleIO = gIOLeftNeedle
        purgeduration = IDS.Data.Hardware.Dispenser.Left.AutoPurgingDuration
        cleanduration = IDS.Data.Hardware.Dispenser.Left.AutoCleaningDuration

        offset(0) = gLeftNeedleOffs(0)
        offset(1) = gLeftNeedleOffs(1)
        offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.PurgePosition.Z
        position(0) = IDS.Data.Hardware.Gantry.PurgePosition.X - offset(0)
        position(1) = IDS.Data.Hardware.Gantry.PurgePosition.Y - offset(1)

        m_Tri.SetMachineRun()
        m_Tri.Set_XY_Speed(IDS.Data.Hardware.Gantry.ElementXYSpeed)
        m_Tri.Set_Z_Speed(IDS.Data.Hardware.Gantry.ElementZSpeed)
        If Not m_Tri.Move_Z(0) Then GoTo StopCalibration 'z axis move to purge pos
        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration 'x and y axis move to purge pos
        If Not m_Tri.Move_Z(offset(2)) Then GoTo StopCalibration 'z axis move to purge pos
        cmdStr = "WA(" + purgeduration.ToString + ")"
        m_Tri.TurnOn("Left Needle IO")
        m_Tri.m_TriCtrl.Execute(cmdStr) 'wait after turning on valve
        m_Tri.TurnOff("Left Needle IO")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo StopCalibration 'z axis move to 0 before move to clean position

        offset(2) = gLeftNeedleOffs(2) + IDS.Data.Hardware.Gantry.CleanPosition.Z
        position(0) = IDS.Data.Hardware.Gantry.CleanPosition.X - offset(0)
        position(1) = IDS.Data.Hardware.Gantry.CleanPosition.Y - offset(1)

        If Not m_Tri.Move_XY(position) Then GoTo StopCalibration 'x and y axis move to clean pos
        If Not m_Tri.Move_Z(offset(2)) Then GoTo StopCalibration 'z axis move to clean pos
        cmdStr = "WA(" + cleanduration.ToString + ")"
        m_Tri.TurnOn("Clean")
        m_Tri.m_TriCtrl.Execute(cmdStr) 'wait after turning on valve
        m_Tri.TurnOff("Clean")
        If Not m_Tri.Move_Z(SafePosition) Then GoTo StopCalibration 'z axis move to 0

        'commented out since kr changed to FSM to handle state
        'If IsRunning() Then 'if in the middle of a continuous run
        '    m_Tri.SetMachineStop()
        '    Exit Sub
        'End If

StopCalibration:
        ResetToIdle()
    End Sub

    Public Function ProdChangeSyringeMaterial(ByVal needle As String) As Integer

    End Function

    ' Start is a continuation of tmrrunpro_tick. it is called when continuous    '
    ' run is not checked (conveyor box), and skips the PLC signal check.             '
    ' it will check the current mode, enable/disable the appropriate timers,         '
    ' check the door/key status, set the status message.                             '
    ' it will then continue the dispensing process through runproduction()           '

    Public Sub Start() 'kr dec24 made some changes - better check how they affect runtime.

        Programming.SetRunMode(4) 'Wet mode

        If DoorCheck() = False Then
            MyMsgBox("Door is open")
            ResetToIdle()
            Exit Sub
        End If

        If (ContinuousMode.Checked = True) Then
            ResetTimer("Reset Volume Calibration Timers")
            ResetTimer("Reset Autopurging Timers")
        End If

        GeneratePatternAndRunProgram()

    End Sub

    Public Sub TBOperation_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles TBOperation.ButtonClick


        If e.Button Is TBOperation.Buttons(0) Then 'run

            SetState("Start")

        ElseIf e.Button Is TBOperation.Buttons(1) Then 'pause

            PauseDispensing()

        ElseIf e.Button Is TBOperation.Buttons(2) Then 'stop

            StopDispensing()

        End If

    End Sub

    Private Sub ButtonHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonHome.Click
        SetState("Homing")
    End Sub

    Private Sub ButtonChgSyringe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonChgSyringe.Click
        SetState("Change Syringe")
    End Sub

    Private Sub ButtonPurge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPurge.Click
        If SetState("Purge") Then DoPurge()
    End Sub

    Private Sub ButtonClean_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonClean.Click
        If SetState("Clean") Then DoClean()
    End Sub

    Private Sub DoorLock_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DoorLock.CheckedChanged

        If IsBusy() Then
            LabelMessage("Can't unlock the door when machine running!")
            DoorLock.Text = "Unlock Door"
            DoorLock.ImageIndex = 2
            Exit Sub
        End If

        If DoorLock.Checked = False Then
            DoorLock.Text = "Unlock Door"
            DoorLock.ImageIndex = 2
            LockDoor()
        Else
            DoorLock.Text = "Lock Door"
            DoorLock.ImageIndex = 3
            UnlockDoor()
        End If
    End Sub

    Private Sub CheckBoxPotOn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBoxPotOn.CheckedChanged

        If CheckBoxPotOn.Checked = True Then
            CheckBoxPotOn.Text = "Turn Pot Life Off"
            CheckBoxPotOn.ImageIndex = 0
            m_PotLifeOn = True
            ResetTimer("Reset Potlife Start Timer")
        Else
            CheckBoxPotOn.Text = "Turn Pot Life On"
            CheckBoxPotOn.ImageIndex = 6
            m_PotLifeOn = False
        End If

    End Sub

    Private Sub ButtonPotReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonPotReset.Click

        ResetTimer("Reset Potlife Start Timer")

    End Sub

    '''''''''''''''
    '                                       kr                                       '
    ' timermonitor is to monitor for potlife and machine state changes               '
    ' if the run state changes to idle, change back to teaching camera vision        '
    ' changes in run state also updates the tower lights                             '
    ' timermonitor updates continuously, does not disable itself                     '
    ' calls prodpurgeclean, volumcalib and prodchangesyringe                         '

    Public Delegate Sub DispensingFinish()
    Public DispensingFinishCallback As DispensingFinish = New DispensingFinish(AddressOf DispensingCallback)

    Public Sub DispensingSwitchCallback(ByVal ar As IAsyncResult)
        ar.AsyncState.EndInvoke(ar)
    End Sub

    Public Sub DispensingCallback()
        If ContinuousMode.Checked = True Then
            Finish_Dispensing()
        Else
            ResetToIdle()
        End If
    End Sub


    Public Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '                                       kr                                       '
        ' this option allows user to choose between continuous operation or to operate   '
        ' on a single board using the buttons start, retrieve and release to control the '
        ' conveyor manually. otherwise the conveyor will be controlled by signals        '

        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag
        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        If door_interlock And door_close = False Then
            LabelMessage("Door is open. Please close to turn on continuous mode.")
            Exit Sub
        ElseIf ContinuousMode.Checked Then
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned off.")
            End If
        ElseIf ContinuousMode.Checked = False Then
            If IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption Then
                LabelMessage("Auto purging and auto cleaning is turned on.")
            End If
        End If

        ContinuousMode.Checked = Not ContinuousMode.Checked

    End Sub

    Public Sub WriteSPCReport()

        'Dim checkID As Integer
        'checkID = CInt(IDS.__ID) Mod 100
        'If (checkID <> 0) Then
        '    '   Cannot write the event while device error is occured.   '
        '    Exit Sub
        'End If


        If ShouldLog() And HasBeenRunning Then Form_Service.LogEventInSPCReport("Production End")

        'production end
        If IDS.Data.Hardware.SPC.EnableSPCReport And HasBeenRunning Then
            Form_Service.GenerateSPCReport()
            HasBeenRunning = False
        End If


    End Sub


    '                                       kr                                       '
    ' tmrRunPro_tick ensures continuous operation in production (if conveyor ticked) '
    ' enabled by tmrfinish, prodpurgeclean, prodvolumcalib and prodchangesyringe     '
    ' also enabled when starting production run                                      '
    ' this function will check to see if starting is valid (m_ButtonState, mode)      '
    ' or else it will exit. if all checks are passed, the last check is for a signal '
    ' from the PLC to tell when the board is ready for dispensing.                   '
    ' then it will call Start                                                    '

    Public Sub Conveyor_Check()

        'Dim onoff As Integer = 1
        'Dim conveyorErrorID, ErrorID As Integer

        'ChangeButtonState("Only Stop Displayed")

        'While onoff = 1 And Not m_Tri.MotionError

        '    m_Tri.m_TriCtrl.In(gIOloadReady, gIOloadReady, onoff) 'when ready set onoff = 0
        '    LabelMessage("Waiting for board..")
        '    TraceDoEvents()

        '    If IsIdle() Or m_Tri.MotionError Then
        '        Form_Service.LogEventInSPCReport("Motion controller error")
        '        ResetToIdle()
        '        Exit Sub
        '    End If

        '    '3202,3203,3204 are errors from lifter
        '    '3205,3206,3208 are conveyor retrieve or stage 1/3 jamming errors.
        '    conveyorErrorID = CInt(IDS.__ID) Mod 10000
        '    If conveyorErrorID = 3202 Or conveyorErrorID = 3203 Or conveyorErrorID = 3204 Or conveyorErrorID = 3205 Or conveyorErrorID = 3206 Or conveyorErrorID = 3208 Then
        '        LabelMessage("No board is coming or conveyor jammed.")
        '        ResetToIdle()
        '        Exit Sub
        '    End If

        'End While

        'ErrorID = CInt(IDS.__ID) Mod 100
        'While ErrorID <> 0
        '    ErrorID = CInt(IDS.__ID) Mod 100
        '    LabelMessage("Clear pending error messages first.")
        '    TraceDoEvents()
        'End While

        If ShouldLog() Then Form_Service.LogEventInSPCReport("Board Comes")
        Start()

    End Sub

    Public Sub ResetTimer(ByVal command As String)

        Select Case command
            Case "Reset Autopurging Timers"
                StartTimeOfMachineIdle = Now
            Case "Reset Potlife Start Timer"
                RemainingTimeOfPotLifeStart = Now
        End Select

    End Sub

    '''''''''''''''
    '                                       kr                                       '
    ' show and set tower light according to machine state                            '
    '''''''''''''''

    Public Sub Finish_Dispensing()

        ' for both single and continuous run
        m_Tri.m_TriCtrl.Op(gIODispensingDone, 1) 'release board
        TravelToParkPosition()
        m_Tri.m_TriCtrl.Op(gIODispensingDone, 0) 'release board

        Form_Service.LogEventInSPCReport("Board Goes")

        If IsRunning() Then
            LabelMessage("Dispensing cycle re-starting.")
            HasBeenRunning = True
            Conveyor_Check() 'after timer finish automatically call back Start()
            Exit Sub
        End If

    End Sub


    Public Function DoorCheck()

        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag

        If door_interlock = False Then
            Return True
        ElseIf door_interlock = True Then
            If door_close Then
                LockDoor()
            End If
            Return door_close
        End If

    End Function

    Private Sub OnHeaters()
        'turn on heaters
        'With IDS.Devices.Thermal.FrmHeater
        '    If ContinuousMode.Checked = True Then
        '        If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '            .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.OperationTemp)
        '            .On_Off_thermal(1, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '            .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.OperationTemp)
        '            .On_Off_thermal(2, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '            .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.OperationTemp)
        '            .On_Off_thermal(3, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '            .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.OperationTemp)
        '            .On_Off_thermal(4, True)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '            .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.OperationTemp)
        '            .On_Off_thermal(5, True)
        '        End If
        '    ElseIf ContinuousMode.Checked = False Then
        '        If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '            .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.StandbyTemp)
        '            .On_Off_thermal(1, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '            .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.StandbyTemp)
        '            .On_Off_thermal(2, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '            .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.StandbyTemp)
        '            .On_Off_thermal(3, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '            .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.StandbyTemp)
        '            .On_Off_thermal(4, False)
        '        End If
        '        If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '            .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.StandbyTemp)
        '            .On_Off_thermal(5, False)
        '        End If
        '    End If
        'End With
    End Sub

    Private Sub HeaterSettings()
        'With IDS.Devices.Thermal.FrmHeater
        '    If IDS.Data.Hardware.Thermal.Needle.OnOff = True Then
        '        Needle.Visible = True
        '        NeedlePicture.Visible = True
        '        NeedleLabel.Visible = True
        '        MyHeaterSettings.Heater_Enabled(0) = True
        '        .Set_Temperature(1, IDS.Data.Hardware.Thermal.Needle.StandbyTemp)
        '        .Set_Alarm(1, IDS.Data.Hardware.Thermal.Needle.AlarmThreshold)
        '    Else
        '        Needle.Visible = False
        '        NeedlePicture.Visible = False
        '        NeedleLabel.Visible = False
        '        MyHeaterSettings.Heater_Enabled(0) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Syringe.OnOff = True Then
        '        Syringe.Visible = True
        '        SyringePicture.Visible = True
        '        SyringeLabel.Visible = True
        '        MyHeaterSettings.Heater_Enabled(1) = True
        '        .Set_Temperature(2, IDS.Data.Hardware.Thermal.Syringe.StandbyTemp)
        '        .Set_Alarm(2, IDS.Data.Hardware.Thermal.Syringe.AlarmThreshold)
        '    Else
        '        Syringe.Visible = False
        '        SyringePicture.Visible = False
        '        SyringeLabel.Visible = False
        '        MyHeaterSettings.Heater_Enabled(1) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station1.OnOff = True Then
        '        Station1.Visible = True
        '        Station1Picture.Visible = True
        '        Station1Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(2) = True
        '        .Set_Temperature(3, IDS.Data.Hardware.Thermal.Station1.StandbyTemp)
        '        .Set_Alarm(3, IDS.Data.Hardware.Thermal.Station1.AlarmThreshold)
        '    Else
        '        Station1.Visible = False
        '        Station1Picture.Visible = False
        '        Station1Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(2) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station2.OnOff = True Then
        '        Station2.Visible = True
        '        Station2Picture.Visible = True
        '        Station2Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(3) = True
        '        .Set_Temperature(4, IDS.Data.Hardware.Thermal.Station2.StandbyTemp)
        '        .Set_Alarm(4, IDS.Data.Hardware.Thermal.Station2.AlarmThreshold)
        '    Else
        '        Station2.Visible = False
        '        Station2Picture.Visible = False
        '        Station2Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(3) = False
        '    End If
        '    If IDS.Data.Hardware.Thermal.Station3.OnOff = True Then
        '        Station3.Visible = True
        '        Station3Picture.Visible = True
        '        Station3Label.Visible = True
        '        MyHeaterSettings.Heater_Enabled(4) = True
        '        .Set_Temperature(5, IDS.Data.Hardware.Thermal.Station3.StandbyTemp)
        '        .Set_Alarm(5, IDS.Data.Hardware.Thermal.Station3.AlarmThreshold)
        '    Else
        '        Station3.Visible = False
        '        Station3Picture.Visible = False
        '        Station3Label.Visible = False
        '        MyHeaterSettings.Heater_Enabled(4) = False
        '    End If
        'End With
    End Sub

    Public Sub PressureSettings()

    End Sub

    Public Sub OnPressure()
        IDS.Data.OpenData()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form_Service.LogEventInSPCReport(ComboBox1.SelectedItem.ToString)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        For i As Integer = 0 To ComboBox1.Items.Count - 1
            Form_Service.LogEventInSPCReport(ComboBox1.Items(i))
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form_Service.GenerateSPCReport()
    End Sub

    Private Sub TimerMonitor_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerMonitor.Tick

        Dim door_close As Boolean = IDS.Devices.DIO.DIO.doorclose_flag
        Dim door_interlock As Boolean = IDS.Devices.DIO.DIO.interlockon_flag
        Dim LowMaterialSensorLeft As Boolean = IDS.Devices.DIO.DIO.LevelSensor(True)

        'reset the timers if 
        '1) door is not closed when door interlocked is true
        '2) program is currently paused or running
        '3) conveyor mode is not ticked
        If (door_close = False And door_interlock = True) And IsIdle() Then
            ResetTimer("Reset Autopurging Timers")
            ResetTimer("Reset Potlife Start Timer")
            Exit Sub
        End If

        Dim auto_purging_enabled As Boolean = IDS.Data.Hardware.Dispenser.Left.AutoPurgingOption And IsIdle() And Form_Service.Visible = False
        CurrentTime = Now

        If m_PotLifeOn Then
            ElapsedTimeSincePotLifeStart = (CurrentTime.Ticks - RemainingTimeOfPotLifeStart.Ticks) * tickToMinute
            If ElapsedTimeSincePotLifeStart >= IDS.Data.Hardware.Dispenser.Left.PotLifeDuration Then
                'what do we want to do when potlife is up?
            End If
        End If

        'auto purging and auto cleaning is only carried out if conveyor mode is ticked and auto purging option is enabled
        If auto_purging_enabled Then
            ElapsedTimeSinceMachineIdle = (CurrentTime.Ticks - StartTimeOfMachineIdle.Ticks) * tickToMinute    'elapsed time
            If ElapsedTimeSinceMachineIdle >= IDS.Data.Hardware.Dispenser.Left.AutoPurgingInterval Then  'when elapsed time > time
                ResetTimer("Reset Autopurging Timers")
                SetState("AutoPurge")
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Form_Service.DisplayErrorMessage(ComboBox3.SelectedItem.ToString)
    End Sub

    Private Sub ButtonCalibrate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCalibrate.Click
        If SetState("Needle Calibration") Then DoCalibrate()
    End Sub

    Private Sub RichTextBoxNote_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBoxNote.TextChanged

    End Sub

End Class
