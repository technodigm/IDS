
Imports Microsoft.VisualBasic
Imports DLL_DataManager
Public Class FormGeneral
    Inherits System.Windows.Forms.Form
    ''''''''''''''''''''''''''''''''''''''''
    '   create instance of export device
    Friend FrmImportConveyor As New DLL_Export_Device_Interface.IDS_CConveyor
    Friend FrmImportTempController As New DLL_Export_Device_Interface.IDS_CTempController

    ''''''''''''''''''''''''''''''''''''''''
    '   create instance of export module
    Friend FrmImportService As New DLL_Export_Service.FormExport
    Friend FrmImportVision As New DLL_Export_Vision.FormExport
    Friend FrmImportExe As New DLL_Export_Execution.FormExport
    Friend FormMyExport As New DLL_Export_SetupAndCaliberate.FormExport

    '''''''''''''''''''''''''''''''''''''''''
    '   create instance 
    Friend MyForm1 As New Form1
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Call InitialSetup()

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
    Friend WithEvents TabBasicSetup As System.Windows.Forms.TabPage
    Friend WithEvents TabDispensingParameter As System.Windows.Forms.TabPage
    Friend WithEvents TabDispenserUnit As System.Windows.Forms.TabPage
    Friend WithEvents TabCaliberation As System.Windows.Forms.TabPage
    Friend WithEvents TabMotionControl As System.Windows.Forms.TabPage
    Friend WithEvents TabSPC As System.Windows.Forms.TabPage
    Friend WithEvents TaBLifterVacuum As System.Windows.Forms.TabPage
    Friend WithEvents TabConveyor As System.Windows.Forms.TabPage
    Friend WithEvents TabThermalController As System.Windows.Forms.TabPage
    Friend WithEvents TabDiagnostic As System.Windows.Forms.TabPage
    Friend WithEvents TabResetDefault As System.Windows.Forms.TabPage
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents TabGeneral As System.Windows.Forms.TabControl
    Public WithEvents TabIDS As System.Windows.Forms.TabControl
    Friend WithEvents GBConveyor As System.Windows.Forms.GroupBox
    Friend WithEvents ReleaseBtn As System.Windows.Forms.Button
    Friend WithEvents Retrievebtn As System.Windows.Forms.Button
    Friend WithEvents PBPostLifter As System.Windows.Forms.PictureBox
    Friend WithEvents PBDispLifter As System.Windows.Forms.PictureBox
    Friend WithEvents PBPostPallet As System.Windows.Forms.PictureBox
    Friend WithEvents PBdispPallet As System.Windows.Forms.PictureBox
    Friend WithEvents PBPrePallet As System.Windows.Forms.PictureBox
    Friend WithEvents PBPrelifter As System.Windows.Forms.PictureBox
    Friend WithEvents PBRelease_1 As System.Windows.Forms.PictureBox
    Friend WithEvents PBRetreive_2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents imageListElement As System.Windows.Forms.ImageList
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents ImageListProgram As System.Windows.Forms.ImageList
    Friend WithEvents PanelVision As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents GBGraphic As System.Windows.Forms.GroupBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents GBStepper As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents PBArrowRight As System.Windows.Forms.PictureBox
    Friend WithEvents PBArrowLeft As System.Windows.Forms.PictureBox
    Friend WithEvents PBArrowBackward As System.Windows.Forms.PictureBox
    Friend WithEvents PBArrowForward As System.Windows.Forms.PictureBox
    Friend WithEvents PBArrowDown As System.Windows.Forms.PictureBox
    Friend WithEvents PBArrowUp As System.Windows.Forms.PictureBox
    Friend WithEvents GBThermal As System.Windows.Forms.GroupBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GBDispenser As System.Windows.Forms.GroupBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Label54 As System.Windows.Forms.Label
    Friend WithEvents Label55 As System.Windows.Forms.Label
    Friend WithEvents Label51 As System.Windows.Forms.Label
    Friend WithEvents Label49 As System.Windows.Forms.Label
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Public WithEvents PanelGeneral As System.Windows.Forms.Panel
    Friend WithEvents BtnAddExecutionPanel As System.Windows.Forms.Button
    Friend WithEvents BtnAddVisionPanel As System.Windows.Forms.Button
    Public WithEvents PanelSetup As System.Windows.Forms.Panel
    Friend WithEvents PanelImport As System.Windows.Forms.Panel
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents BtnLoadForm As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormGeneral))
        Me.TabGeneral = New System.Windows.Forms.TabControl
        Me.TabBasicSetup = New System.Windows.Forms.TabPage
        Me.TabThermalController = New System.Windows.Forms.TabPage
        Me.TabDispenserUnit = New System.Windows.Forms.TabPage
        Me.TabResetDefault = New System.Windows.Forms.TabPage
        Me.TabCaliberation = New System.Windows.Forms.TabPage
        Me.TabDispensingParameter = New System.Windows.Forms.TabPage
        Me.TabConveyor = New System.Windows.Forms.TabPage
        Me.TaBLifterVacuum = New System.Windows.Forms.TabPage
        Me.TabSPC = New System.Windows.Forms.TabPage
        Me.BtnLoadForm = New System.Windows.Forms.Button
        Me.TabMotionControl = New System.Windows.Forms.TabPage
        Me.TabDiagnostic = New System.Windows.Forms.TabPage
        Me.TabIDS = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.PanelSetup = New System.Windows.Forms.Panel
        Me.PanelGeneral = New System.Windows.Forms.Panel
        Me.GBGraphic = New System.Windows.Forms.GroupBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.GBStepper = New System.Windows.Forms.GroupBox
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.PBArrowRight = New System.Windows.Forms.PictureBox
        Me.PBArrowLeft = New System.Windows.Forms.PictureBox
        Me.PBArrowBackward = New System.Windows.Forms.PictureBox
        Me.PBArrowForward = New System.Windows.Forms.PictureBox
        Me.PBArrowDown = New System.Windows.Forms.PictureBox
        Me.PBArrowUp = New System.Windows.Forms.PictureBox
        Me.GBThermal = New System.Windows.Forms.GroupBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.GBDispenser = New System.Windows.Forms.GroupBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Label54 = New System.Windows.Forms.Label
        Me.Label55 = New System.Windows.Forms.Label
        Me.Label51 = New System.Windows.Forms.Label
        Me.Label49 = New System.Windows.Forms.Label
        Me.Label47 = New System.Windows.Forms.Label
        Me.Label46 = New System.Windows.Forms.Label
        Me.PanelImport = New System.Windows.Forms.Panel
        Me.BtnAddExecutionPanel = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.GBConveyor = New System.Windows.Forms.GroupBox
        Me.ReleaseBtn = New System.Windows.Forms.Button
        Me.Retrievebtn = New System.Windows.Forms.Button
        Me.PBPostLifter = New System.Windows.Forms.PictureBox
        Me.PBDispLifter = New System.Windows.Forms.PictureBox
        Me.PBPostPallet = New System.Windows.Forms.PictureBox
        Me.PBdispPallet = New System.Windows.Forms.PictureBox
        Me.PBPrePallet = New System.Windows.Forms.PictureBox
        Me.PBPrelifter = New System.Windows.Forms.PictureBox
        Me.PBRelease_1 = New System.Windows.Forms.PictureBox
        Me.PBRetreive_2 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PanelVision = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.BtnAddVisionPanel = New System.Windows.Forms.Button
        Me.ImageListProgram = New System.Windows.Forms.ImageList(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.imageListElement = New System.Windows.Forms.ImageList(Me.components)
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.TabGeneral.SuspendLayout()
        Me.TabSPC.SuspendLayout()
        Me.TabIDS.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.PanelSetup.SuspendLayout()
        Me.PanelGeneral.SuspendLayout()
        Me.GBGraphic.SuspendLayout()
        Me.GBStepper.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GBThermal.SuspendLayout()
        Me.GBDispenser.SuspendLayout()
        Me.PanelImport.SuspendLayout()
        Me.GBConveyor.SuspendLayout()
        Me.PanelVision.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabGeneral
        '
        Me.TabGeneral.Controls.Add(Me.TabBasicSetup)
        Me.TabGeneral.Controls.Add(Me.TabThermalController)
        Me.TabGeneral.Controls.Add(Me.TabDispenserUnit)
        Me.TabGeneral.Controls.Add(Me.TabResetDefault)
        Me.TabGeneral.Controls.Add(Me.TabCaliberation)
        Me.TabGeneral.Controls.Add(Me.TabDispensingParameter)
        Me.TabGeneral.Controls.Add(Me.TabConveyor)
        Me.TabGeneral.Controls.Add(Me.TaBLifterVacuum)
        Me.TabGeneral.Controls.Add(Me.TabSPC)
        Me.TabGeneral.Controls.Add(Me.TabMotionControl)
        Me.TabGeneral.Controls.Add(Me.TabDiagnostic)
        Me.TabGeneral.Location = New System.Drawing.Point(0, 0)
        Me.TabGeneral.Multiline = True
        Me.TabGeneral.Name = "TabGeneral"
        Me.TabGeneral.Padding = New System.Drawing.Point(15, 6)
        Me.TabGeneral.SelectedIndex = 0
        Me.TabGeneral.Size = New System.Drawing.Size(536, 648)
        Me.TabGeneral.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabGeneral.TabIndex = 58
        '
        'TabBasicSetup
        '
        Me.TabBasicSetup.ForeColor = System.Drawing.Color.Blue
        Me.TabBasicSetup.Location = New System.Drawing.Point(4, 76)
        Me.TabBasicSetup.Name = "TabBasicSetup"
        Me.TabBasicSetup.Size = New System.Drawing.Size(528, 568)
        Me.TabBasicSetup.TabIndex = 0
        Me.TabBasicSetup.Text = "Basic Setup"
        '
        'TabThermalController
        '
        Me.TabThermalController.Location = New System.Drawing.Point(4, 76)
        Me.TabThermalController.Name = "TabThermalController"
        Me.TabThermalController.Size = New System.Drawing.Size(528, 568)
        Me.TabThermalController.TabIndex = 8
        Me.TabThermalController.Text = "Thermal Controller"
        Me.TabThermalController.Visible = False
        '
        'TabDispenserUnit
        '
        Me.TabDispenserUnit.Location = New System.Drawing.Point(4, 76)
        Me.TabDispenserUnit.Name = "TabDispenserUnit"
        Me.TabDispenserUnit.Size = New System.Drawing.Size(528, 568)
        Me.TabDispenserUnit.TabIndex = 2
        Me.TabDispenserUnit.Text = "Dispenser Unit"
        Me.TabDispenserUnit.Visible = False
        '
        'TabResetDefault
        '
        Me.TabResetDefault.Location = New System.Drawing.Point(4, 76)
        Me.TabResetDefault.Name = "TabResetDefault"
        Me.TabResetDefault.Size = New System.Drawing.Size(528, 568)
        Me.TabResetDefault.TabIndex = 10
        Me.TabResetDefault.Text = "Reset Default"
        Me.TabResetDefault.Visible = False
        '
        'TabCaliberation
        '
        Me.TabCaliberation.Location = New System.Drawing.Point(4, 76)
        Me.TabCaliberation.Name = "TabCaliberation"
        Me.TabCaliberation.Size = New System.Drawing.Size(528, 568)
        Me.TabCaliberation.TabIndex = 5
        Me.TabCaliberation.Text = "Caliberation"
        Me.TabCaliberation.Visible = False
        '
        'TabDispensingParameter
        '
        Me.TabDispensingParameter.Location = New System.Drawing.Point(4, 76)
        Me.TabDispensingParameter.Name = "TabDispensingParameter"
        Me.TabDispensingParameter.Size = New System.Drawing.Size(528, 568)
        Me.TabDispensingParameter.TabIndex = 3
        Me.TabDispensingParameter.Text = "Dispensing Parameter"
        Me.TabDispensingParameter.Visible = False
        '
        'TabConveyor
        '
        Me.TabConveyor.Location = New System.Drawing.Point(4, 76)
        Me.TabConveyor.Name = "TabConveyor"
        Me.TabConveyor.Size = New System.Drawing.Size(528, 568)
        Me.TabConveyor.TabIndex = 7
        Me.TabConveyor.Text = "Conveyor"
        Me.TabConveyor.Visible = False
        '
        'TaBLifterVacuum
        '
        Me.TaBLifterVacuum.Location = New System.Drawing.Point(4, 76)
        Me.TaBLifterVacuum.Name = "TaBLifterVacuum"
        Me.TaBLifterVacuum.Size = New System.Drawing.Size(528, 568)
        Me.TaBLifterVacuum.TabIndex = 6
        Me.TaBLifterVacuum.Text = "Lifter&Vacuum"
        Me.TaBLifterVacuum.Visible = False
        '
        'TabSPC
        '
        Me.TabSPC.Controls.Add(Me.BtnLoadForm)
        Me.TabSPC.Location = New System.Drawing.Point(4, 76)
        Me.TabSPC.Name = "TabSPC"
        Me.TabSPC.Size = New System.Drawing.Size(528, 568)
        Me.TabSPC.TabIndex = 1
        Me.TabSPC.Text = "SPC"
        Me.TabSPC.Visible = False
        '
        'BtnLoadForm
        '
        Me.BtnLoadForm.Location = New System.Drawing.Point(24, 40)
        Me.BtnLoadForm.Name = "BtnLoadForm"
        Me.BtnLoadForm.Size = New System.Drawing.Size(96, 40)
        Me.BtnLoadForm.TabIndex = 4
        Me.BtnLoadForm.Text = "LoadForm"
        '
        'TabMotionControl
        '
        Me.TabMotionControl.Location = New System.Drawing.Point(4, 76)
        Me.TabMotionControl.Name = "TabMotionControl"
        Me.TabMotionControl.Size = New System.Drawing.Size(528, 568)
        Me.TabMotionControl.TabIndex = 4
        Me.TabMotionControl.Text = "Motion Control"
        Me.TabMotionControl.Visible = False
        '
        'TabDiagnostic
        '
        Me.TabDiagnostic.Location = New System.Drawing.Point(4, 76)
        Me.TabDiagnostic.Name = "TabDiagnostic"
        Me.TabDiagnostic.Size = New System.Drawing.Size(528, 568)
        Me.TabDiagnostic.TabIndex = 9
        Me.TabDiagnostic.Text = "Diagnostic"
        Me.TabDiagnostic.Visible = False
        '
        'TabIDS
        '
        Me.TabIDS.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TabIDS.Controls.Add(Me.TabPage1)
        Me.TabIDS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabIDS.Location = New System.Drawing.Point(0, 0)
        Me.TabIDS.Multiline = True
        Me.TabIDS.Name = "TabIDS"
        Me.TabIDS.Padding = New System.Drawing.Point(20, 6)
        Me.TabIDS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.TabIDS.SelectedIndex = 0
        Me.TabIDS.Size = New System.Drawing.Size(536, 680)
        Me.TabIDS.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabIDS.TabIndex = 59
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.TabGeneral)
        Me.TabPage1.Location = New System.Drawing.Point(4, 28)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(528, 648)
        Me.TabPage1.TabIndex = 14
        Me.TabPage1.Text = "General"
        '
        'PanelSetup
        '
        Me.PanelSetup.Controls.Add(Me.PanelGeneral)
        Me.PanelSetup.Controls.Add(Me.PanelImport)
        Me.PanelSetup.Controls.Add(Me.GBConveyor)
        Me.PanelSetup.Controls.Add(Me.PanelVision)
        Me.PanelSetup.Location = New System.Drawing.Point(536, 0)
        Me.PanelSetup.Name = "PanelSetup"
        Me.PanelSetup.Size = New System.Drawing.Size(656, 664)
        Me.PanelSetup.TabIndex = 70
        '
        'PanelGeneral
        '
        Me.PanelGeneral.BackColor = System.Drawing.Color.LightGray
        Me.PanelGeneral.Controls.Add(Me.GBGraphic)
        Me.PanelGeneral.Controls.Add(Me.GBStepper)
        Me.PanelGeneral.Controls.Add(Me.GBThermal)
        Me.PanelGeneral.Controls.Add(Me.GBDispenser)
        Me.PanelGeneral.Location = New System.Drawing.Point(232, 0)
        Me.PanelGeneral.Name = "PanelGeneral"
        Me.PanelGeneral.Size = New System.Drawing.Size(248, 256)
        Me.PanelGeneral.TabIndex = 69
        '
        'GBGraphic
        '
        Me.GBGraphic.Controls.Add(Me.PictureBox2)
        Me.GBGraphic.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBGraphic.Location = New System.Drawing.Point(264, 8)
        Me.GBGraphic.Name = "GBGraphic"
        Me.GBGraphic.Size = New System.Drawing.Size(248, 160)
        Me.GBGraphic.TabIndex = 65
        Me.GBGraphic.TabStop = False
        Me.GBGraphic.Text = "Graphic "
        Me.GBGraphic.Visible = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(8, 16)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(232, 136)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 69
        Me.PictureBox2.TabStop = False
        '
        'GBStepper
        '
        Me.GBStepper.Controls.Add(Me.NumericUpDown1)
        Me.GBStepper.Controls.Add(Me.Label5)
        Me.GBStepper.Controls.Add(Me.PBArrowRight)
        Me.GBStepper.Controls.Add(Me.PBArrowLeft)
        Me.GBStepper.Controls.Add(Me.PBArrowBackward)
        Me.GBStepper.Controls.Add(Me.PBArrowForward)
        Me.GBStepper.Controls.Add(Me.PBArrowDown)
        Me.GBStepper.Controls.Add(Me.PBArrowUp)
        Me.GBStepper.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBStepper.Location = New System.Drawing.Point(272, 176)
        Me.GBStepper.Name = "GBStepper"
        Me.GBStepper.Size = New System.Drawing.Size(248, 152)
        Me.GBStepper.TabIndex = 63
        Me.GBStepper.TabStop = False
        Me.GBStepper.Text = "Stepper "
        Me.GBStepper.Visible = False
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(16, 72)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(48, 20)
        Me.NumericUpDown1.TabIndex = 33
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(8, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 16)
        Me.Label5.TabIndex = 32
        Me.Label5.Text = "Fine Step"
        '
        'PBArrowRight
        '
        Me.PBArrowRight.Image = CType(resources.GetObject("PBArrowRight.Image"), System.Drawing.Image)
        Me.PBArrowRight.Location = New System.Drawing.Point(200, 56)
        Me.PBArrowRight.Name = "PBArrowRight"
        Me.PBArrowRight.Size = New System.Drawing.Size(32, 32)
        Me.PBArrowRight.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowRight.TabIndex = 20
        Me.PBArrowRight.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PBArrowRight, "Arrow Right")
        '
        'PBArrowLeft
        '
        Me.PBArrowLeft.Image = CType(resources.GetObject("PBArrowLeft.Image"), System.Drawing.Image)
        Me.PBArrowLeft.Location = New System.Drawing.Point(136, 56)
        Me.PBArrowLeft.Name = "PBArrowLeft"
        Me.PBArrowLeft.Size = New System.Drawing.Size(32, 32)
        Me.PBArrowLeft.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowLeft.TabIndex = 21
        Me.PBArrowLeft.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PBArrowLeft, "Arrow Left")
        '
        'PBArrowBackward
        '
        Me.PBArrowBackward.Image = CType(resources.GetObject("PBArrowBackward.Image"), System.Drawing.Image)
        Me.PBArrowBackward.Location = New System.Drawing.Point(168, 24)
        Me.PBArrowBackward.Name = "PBArrowBackward"
        Me.PBArrowBackward.Size = New System.Drawing.Size(32, 32)
        Me.PBArrowBackward.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowBackward.TabIndex = 22
        Me.PBArrowBackward.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PBArrowBackward, "Arrow Backward")
        '
        'PBArrowForward
        '
        Me.PBArrowForward.Image = CType(resources.GetObject("PBArrowForward.Image"), System.Drawing.Image)
        Me.PBArrowForward.Location = New System.Drawing.Point(168, 88)
        Me.PBArrowForward.Name = "PBArrowForward"
        Me.PBArrowForward.Size = New System.Drawing.Size(32, 32)
        Me.PBArrowForward.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowForward.TabIndex = 23
        Me.PBArrowForward.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PBArrowForward, "Arrow Forward")
        '
        'PBArrowDown
        '
        Me.HelpProvider1.SetHelpNavigator(Me.PBArrowDown, System.Windows.Forms.HelpNavigator.TableOfContents)
        Me.PBArrowDown.Image = CType(resources.GetObject("PBArrowDown.Image"), System.Drawing.Image)
        Me.PBArrowDown.Location = New System.Drawing.Point(88, 80)
        Me.PBArrowDown.Name = "PBArrowDown"
        Me.HelpProvider1.SetShowHelp(Me.PBArrowDown, True)
        Me.PBArrowDown.Size = New System.Drawing.Size(40, 40)
        Me.PBArrowDown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowDown.TabIndex = 26
        Me.PBArrowDown.TabStop = False
        Me.PBArrowDown.Tag = "Arrow Down"
        Me.ToolTip1.SetToolTip(Me.PBArrowDown, "Arrow Down")
        '
        'PBArrowUp
        '
        Me.PBArrowUp.Image = CType(resources.GetObject("PBArrowUp.Image"), System.Drawing.Image)
        Me.PBArrowUp.Location = New System.Drawing.Point(88, 24)
        Me.PBArrowUp.Name = "PBArrowUp"
        Me.PBArrowUp.Size = New System.Drawing.Size(40, 40)
        Me.PBArrowUp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBArrowUp.TabIndex = 27
        Me.PBArrowUp.TabStop = False
        Me.PBArrowUp.Tag = ""
        Me.ToolTip1.SetToolTip(Me.PBArrowUp, "Arrow Up")
        '
        'GBThermal
        '
        Me.GBThermal.BackColor = System.Drawing.Color.LightGray
        Me.GBThermal.Controls.Add(Me.Label16)
        Me.GBThermal.Controls.Add(Me.Label6)
        Me.GBThermal.Controls.Add(Me.Label7)
        Me.GBThermal.Controls.Add(Me.Label19)
        Me.GBThermal.Controls.Add(Me.Label17)
        Me.GBThermal.Controls.Add(Me.Label15)
        Me.GBThermal.Controls.Add(Me.Label13)
        Me.GBThermal.Controls.Add(Me.Label14)
        Me.GBThermal.Controls.Add(Me.Label11)
        Me.GBThermal.Controls.Add(Me.Label12)
        Me.GBThermal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBThermal.ForeColor = System.Drawing.Color.Black
        Me.GBThermal.Location = New System.Drawing.Point(272, 336)
        Me.GBThermal.Name = "GBThermal"
        Me.GBThermal.Size = New System.Drawing.Size(248, 152)
        Me.GBThermal.TabIndex = 23
        Me.GBThermal.TabStop = False
        Me.GBThermal.Text = "Thermal"
        Me.GBThermal.Visible = False
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Menu
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Black
        Me.Label16.Location = New System.Drawing.Point(40, 56)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(88, 16)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "Needle Heater"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Menu
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(40, 24)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Springe Heater"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Menu
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(176, 96)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "Post Heat"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.Color.Gainsboro
        Me.Label19.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Blue
        Me.Label19.Location = New System.Drawing.Point(176, 112)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(64, 16)
        Me.Label19.TabIndex = 33
        Me.Label19.Text = "1.000"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.Gainsboro
        Me.Label17.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label17.ForeColor = System.Drawing.Color.Blue
        Me.Label17.Location = New System.Drawing.Point(96, 112)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(64, 16)
        Me.Label17.TabIndex = 31
        Me.Label17.Text = "1.000"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Gainsboro
        Me.Label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Blue
        Me.Label15.Location = New System.Drawing.Point(8, 112)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 29
        Me.Label15.Text = "1.000"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Gainsboro
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Blue
        Me.Label13.Location = New System.Drawing.Point(128, 56)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 16)
        Me.Label13.TabIndex = 27
        Me.Label13.Text = "1.000"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Menu
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.Black
        Me.Label14.Location = New System.Drawing.Point(88, 96)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(72, 16)
        Me.Label14.TabIndex = 26
        Me.Label14.Text = "Disp. Heater"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.Color.Gainsboro
        Me.Label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(128, 24)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(56, 16)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "1.000"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Menu
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(8, 96)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 16)
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Pre  Heat"
        '
        'GBDispenser
        '
        Me.GBDispenser.BackColor = System.Drawing.Color.LightGray
        Me.GBDispenser.Controls.Add(Me.Label26)
        Me.GBDispenser.Controls.Add(Me.RadioButton2)
        Me.GBDispenser.Controls.Add(Me.RadioButton1)
        Me.GBDispenser.Controls.Add(Me.Button4)
        Me.GBDispenser.Controls.Add(Me.Button5)
        Me.GBDispenser.Controls.Add(Me.Label54)
        Me.GBDispenser.Controls.Add(Me.Label55)
        Me.GBDispenser.Controls.Add(Me.Label51)
        Me.GBDispenser.Controls.Add(Me.Label49)
        Me.GBDispenser.Controls.Add(Me.Label47)
        Me.GBDispenser.Controls.Add(Me.Label46)
        Me.GBDispenser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBDispenser.ForeColor = System.Drawing.Color.Black
        Me.GBDispenser.Location = New System.Drawing.Point(272, 496)
        Me.GBDispenser.Name = "GBDispenser"
        Me.GBDispenser.Size = New System.Drawing.Size(248, 152)
        Me.GBDispenser.TabIndex = 4
        Me.GBDispenser.TabStop = False
        Me.GBDispenser.Text = "Dispenser"
        Me.GBDispenser.Visible = False
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.SystemColors.Menu
        Me.Label26.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label26.ForeColor = System.Drawing.Color.Black
        Me.Label26.Location = New System.Drawing.Point(16, 24)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(56, 16)
        Me.Label26.TabIndex = 39
        Me.Label26.Text = "Valve"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'RadioButton2
        '
        Me.RadioButton2.BackColor = System.Drawing.SystemColors.Menu
        Me.RadioButton2.Checked = True
        Me.RadioButton2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton2.ForeColor = System.Drawing.Color.Black
        Me.RadioButton2.Location = New System.Drawing.Point(152, 24)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(64, 16)
        Me.RadioButton2.TabIndex = 38
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "Right"
        '
        'RadioButton1
        '
        Me.RadioButton1.BackColor = System.Drawing.SystemColors.Menu
        Me.RadioButton1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton1.ForeColor = System.Drawing.Color.Black
        Me.RadioButton1.Location = New System.Drawing.Point(88, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(64, 16)
        Me.RadioButton1.TabIndex = 37
        Me.RadioButton1.Text = "Left"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.SystemColors.Menu
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(152, 112)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(64, 24)
        Me.Button4.TabIndex = 36
        Me.Button4.Text = "Update"
        '
        'Button5
        '
        Me.Button5.BackColor = System.Drawing.SystemColors.Menu
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.ForeColor = System.Drawing.Color.Black
        Me.Button5.Location = New System.Drawing.Point(80, 112)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(64, 24)
        Me.Button5.TabIndex = 35
        Me.Button5.Text = "Update"
        '
        'Label54
        '
        Me.Label54.BackColor = System.Drawing.Color.Gainsboro
        Me.Label54.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label54.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label54.ForeColor = System.Drawing.Color.Blue
        Me.Label54.Location = New System.Drawing.Point(160, 48)
        Me.Label54.Name = "Label54"
        Me.Label54.Size = New System.Drawing.Size(56, 16)
        Me.Label54.TabIndex = 34
        Me.Label54.Text = "1.000"
        '
        'Label55
        '
        Me.Label55.BackColor = System.Drawing.Color.Gainsboro
        Me.Label55.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label55.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label55.ForeColor = System.Drawing.Color.Blue
        Me.Label55.Location = New System.Drawing.Point(160, 72)
        Me.Label55.Name = "Label55"
        Me.Label55.Size = New System.Drawing.Size(56, 16)
        Me.Label55.TabIndex = 33
        Me.Label55.Text = "1.000"
        '
        'Label51
        '
        Me.Label51.BackColor = System.Drawing.SystemColors.Menu
        Me.Label51.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label51.ForeColor = System.Drawing.Color.Black
        Me.Label51.Location = New System.Drawing.Point(16, 72)
        Me.Label51.Name = "Label51"
        Me.Label51.Size = New System.Drawing.Size(64, 16)
        Me.Label51.TabIndex = 29
        Me.Label51.Text = "Suck Back"
        Me.Label51.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label49
        '
        Me.Label49.BackColor = System.Drawing.Color.Gainsboro
        Me.Label49.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label49.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label49.ForeColor = System.Drawing.Color.Blue
        Me.Label49.Location = New System.Drawing.Point(88, 48)
        Me.Label49.Name = "Label49"
        Me.Label49.Size = New System.Drawing.Size(56, 16)
        Me.Label49.TabIndex = 28
        Me.Label49.Text = "1.000"
        '
        'Label47
        '
        Me.Label47.BackColor = System.Drawing.Color.Gainsboro
        Me.Label47.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label47.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.ForeColor = System.Drawing.Color.Blue
        Me.Label47.Location = New System.Drawing.Point(88, 72)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(56, 16)
        Me.Label47.TabIndex = 27
        Me.Label47.Text = "1.000"
        '
        'Label46
        '
        Me.Label46.BackColor = System.Drawing.SystemColors.Menu
        Me.Label46.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.Black
        Me.Label46.Location = New System.Drawing.Point(16, 48)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(56, 16)
        Me.Label46.TabIndex = 26
        Me.Label46.Text = "Pressure"
        Me.Label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PanelImport
        '
        Me.PanelImport.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(192, Byte), CType(255, Byte))
        Me.PanelImport.Controls.Add(Me.BtnAddExecutionPanel)
        Me.PanelImport.Controls.Add(Me.Label4)
        Me.PanelImport.Location = New System.Drawing.Point(0, 0)
        Me.PanelImport.Name = "PanelImport"
        Me.PanelImport.Size = New System.Drawing.Size(232, 256)
        Me.PanelImport.TabIndex = 65
        '
        'BtnAddExecutionPanel
        '
        Me.BtnAddExecutionPanel.Location = New System.Drawing.Point(56, 24)
        Me.BtnAddExecutionPanel.Name = "BtnAddExecutionPanel"
        Me.BtnAddExecutionPanel.Size = New System.Drawing.Size(120, 56)
        Me.BtnAddExecutionPanel.TabIndex = 2
        Me.BtnAddExecutionPanel.Text = "Click ME to Add Panel"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(168, 88)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Import components of DLL Execution to here"
        '
        'GBConveyor
        '
        Me.GBConveyor.Controls.Add(Me.ReleaseBtn)
        Me.GBConveyor.Controls.Add(Me.Retrievebtn)
        Me.GBConveyor.Controls.Add(Me.PBPostLifter)
        Me.GBConveyor.Controls.Add(Me.PBDispLifter)
        Me.GBConveyor.Controls.Add(Me.PBPostPallet)
        Me.GBConveyor.Controls.Add(Me.PBdispPallet)
        Me.GBConveyor.Controls.Add(Me.PBPrePallet)
        Me.GBConveyor.Controls.Add(Me.PBPrelifter)
        Me.GBConveyor.Controls.Add(Me.PBRelease_1)
        Me.GBConveyor.Controls.Add(Me.PBRetreive_2)
        Me.GBConveyor.Controls.Add(Me.PictureBox3)
        Me.GBConveyor.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GBConveyor.Location = New System.Drawing.Point(744, 16)
        Me.GBConveyor.Name = "GBConveyor"
        Me.GBConveyor.Size = New System.Drawing.Size(248, 152)
        Me.GBConveyor.TabIndex = 64
        Me.GBConveyor.TabStop = False
        Me.GBConveyor.Text = "Conveyor"
        Me.GBConveyor.Visible = False
        '
        'ReleaseBtn
        '
        Me.ReleaseBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReleaseBtn.Location = New System.Drawing.Point(144, 112)
        Me.ReleaseBtn.Name = "ReleaseBtn"
        Me.ReleaseBtn.Size = New System.Drawing.Size(56, 24)
        Me.ReleaseBtn.TabIndex = 86
        Me.ReleaseBtn.Text = "Release"
        '
        'Retrievebtn
        '
        Me.Retrievebtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Retrievebtn.Location = New System.Drawing.Point(48, 112)
        Me.Retrievebtn.Name = "Retrievebtn"
        Me.Retrievebtn.Size = New System.Drawing.Size(56, 24)
        Me.Retrievebtn.TabIndex = 85
        Me.Retrievebtn.Text = "Retrieve"
        '
        'PBPostLifter
        '
        Me.PBPostLifter.Image = CType(resources.GetObject("PBPostLifter.Image"), System.Drawing.Image)
        Me.PBPostLifter.Location = New System.Drawing.Point(180, 64)
        Me.PBPostLifter.Name = "PBPostLifter"
        Me.PBPostLifter.Size = New System.Drawing.Size(40, 32)
        Me.PBPostLifter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBPostLifter.TabIndex = 84
        Me.PBPostLifter.TabStop = False
        '
        'PBDispLifter
        '
        Me.PBDispLifter.Image = CType(resources.GetObject("PBDispLifter.Image"), System.Drawing.Image)
        Me.PBDispLifter.Location = New System.Drawing.Point(96, 64)
        Me.PBDispLifter.Name = "PBDispLifter"
        Me.PBDispLifter.Size = New System.Drawing.Size(40, 32)
        Me.PBDispLifter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBDispLifter.TabIndex = 83
        Me.PBDispLifter.TabStop = False
        '
        'PBPostPallet
        '
        Me.PBPostPallet.Image = CType(resources.GetObject("PBPostPallet.Image"), System.Drawing.Image)
        Me.PBPostPallet.Location = New System.Drawing.Point(180, 40)
        Me.PBPostPallet.Name = "PBPostPallet"
        Me.PBPostPallet.Size = New System.Drawing.Size(40, 16)
        Me.PBPostPallet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBPostPallet.TabIndex = 82
        Me.PBPostPallet.TabStop = False
        '
        'PBdispPallet
        '
        Me.PBdispPallet.Image = CType(resources.GetObject("PBdispPallet.Image"), System.Drawing.Image)
        Me.PBdispPallet.Location = New System.Drawing.Point(96, 40)
        Me.PBdispPallet.Name = "PBdispPallet"
        Me.PBdispPallet.Size = New System.Drawing.Size(40, 16)
        Me.PBdispPallet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBdispPallet.TabIndex = 80
        Me.PBdispPallet.TabStop = False
        '
        'PBPrePallet
        '
        Me.PBPrePallet.Image = CType(resources.GetObject("PBPrePallet.Image"), System.Drawing.Image)
        Me.PBPrePallet.Location = New System.Drawing.Point(20, 40)
        Me.PBPrePallet.Name = "PBPrePallet"
        Me.PBPrePallet.Size = New System.Drawing.Size(32, 16)
        Me.PBPrePallet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBPrePallet.TabIndex = 79
        Me.PBPrePallet.TabStop = False
        '
        'PBPrelifter
        '
        Me.PBPrelifter.Image = CType(resources.GetObject("PBPrelifter.Image"), System.Drawing.Image)
        Me.PBPrelifter.Location = New System.Drawing.Point(12, 64)
        Me.PBPrelifter.Name = "PBPrelifter"
        Me.PBPrelifter.Size = New System.Drawing.Size(40, 32)
        Me.PBPrelifter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBPrelifter.TabIndex = 81
        Me.PBPrelifter.TabStop = False
        '
        'PBRelease_1
        '
        Me.PBRelease_1.Image = CType(resources.GetObject("PBRelease_1.Image"), System.Drawing.Image)
        Me.PBRelease_1.Location = New System.Drawing.Point(152, 56)
        Me.PBRelease_1.Name = "PBRelease_1"
        Me.PBRelease_1.Size = New System.Drawing.Size(16, 16)
        Me.PBRelease_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBRelease_1.TabIndex = 89
        Me.PBRelease_1.TabStop = False
        '
        'PBRetreive_2
        '
        Me.PBRetreive_2.Image = CType(resources.GetObject("PBRetreive_2.Image"), System.Drawing.Image)
        Me.PBRetreive_2.Location = New System.Drawing.Point(64, 56)
        Me.PBRetreive_2.Name = "PBRetreive_2"
        Me.PBRetreive_2.Size = New System.Drawing.Size(16, 16)
        Me.PBRetreive_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PBRetreive_2.TabIndex = 88
        Me.PBRetreive_2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.SystemColors.Menu
        Me.PictureBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox3.ForeColor = System.Drawing.Color.Black
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(8, 12)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(232, 132)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 90
        Me.PictureBox3.TabStop = False
        '
        'PanelVision
        '
        Me.PanelVision.BackColor = System.Drawing.Color.Thistle
        Me.PanelVision.Controls.Add(Me.Label2)
        Me.PanelVision.Controls.Add(Me.BtnAddVisionPanel)
        Me.PanelVision.Location = New System.Drawing.Point(0, 256)
        Me.PanelVision.Name = "PanelVision"
        Me.PanelVision.Size = New System.Drawing.Size(480, 408)
        Me.PanelVision.TabIndex = 66
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(128, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(232, 56)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Import compoent of DLL vision  to here"
        '
        'BtnAddVisionPanel
        '
        Me.BtnAddVisionPanel.Location = New System.Drawing.Point(184, 88)
        Me.BtnAddVisionPanel.Name = "BtnAddVisionPanel"
        Me.BtnAddVisionPanel.Size = New System.Drawing.Size(120, 56)
        Me.BtnAddVisionPanel.TabIndex = 3
        Me.BtnAddVisionPanel.Text = "Click ME to Add Panel"
        '
        'ImageListProgram
        '
        Me.ImageListProgram.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageListProgram.ImageStream = CType(resources.GetObject("ImageListProgram.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListProgram.TransparentColor = System.Drawing.Color.Transparent
        '
        'HelpProvider1
        '
        Me.HelpProvider1.HelpNamespace = "D:\Epoxy\HelpFiles\Help One\HOne.chm"
        '
        'imageListElement
        '
        Me.imageListElement.ImageSize = New System.Drawing.Size(16, 16)
        Me.imageListElement.ImageStream = CType(resources.GetObject("imageListElement.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imageListElement.TransparentColor = System.Drawing.Color.Transparent
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "Files"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem6, Me.MenuItem7, Me.MenuItem8})
        Me.MenuItem4.Text = "Operations"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "Start"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.Text = "Pause"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "Stop"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 3
        Me.MenuItem8.Text = "Conveyor Mode"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 2
        Me.MenuItem9.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem10, Me.MenuItem11, Me.MenuItem12, Me.MenuItem13, Me.MenuItem14, Me.MenuItem15})
        Me.MenuItem9.Text = "Maintenance"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 0
        Me.MenuItem10.Text = "Goto Soft Home"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 1
        Me.MenuItem11.Text = "Change Syringe"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.Text = "Clean Syringe"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 3
        Me.MenuItem13.Text = "Purge Syringe"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 4
        Me.MenuItem14.Text = "Volume Calibration"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 5
        Me.MenuItem15.Text = "Needle Calibratipm"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem4, Me.MenuItem9, Me.MenuItem1})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 3
        Me.MenuItem1.Text = ""
        '
        'FormGeneral
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 13)
        Me.ClientSize = New System.Drawing.Size(1024, 701)
        Me.Controls.Add(Me.TabIDS)
        Me.Controls.Add(Me.PanelSetup)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Menu = Me.MainMenu1
        Me.Name = "FormGeneral"
        Me.Text = "FormGeneral"
        Me.TabGeneral.ResumeLayout(False)
        Me.TabSPC.ResumeLayout(False)
        Me.TabIDS.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.PanelSetup.ResumeLayout(False)
        Me.PanelGeneral.ResumeLayout(False)
        Me.GBGraphic.ResumeLayout(False)
        Me.GBStepper.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GBThermal.ResumeLayout(False)
        Me.GBDispenser.ResumeLayout(False)
        Me.PanelImport.ResumeLayout(False)
        Me.GBConveyor.ResumeLayout(False)
        Me.PanelVision.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " LSGoh "

#Region " Data Reference "

    Public Sub SetDataReference(ByRef DataManagerIDS As DLL_DataManager.CIDS, ByRef DataManagerDS As DataSet)

        _IDS = DataManagerIDS
        DS = DataManagerDS

        Call FormMyExport.SetDataReference(_IDS, DS)

    End Sub
#End Region

    ' Below demo some of the ways of calling to other DLL

#Region " DLL Service "

    Private Sub CallService()
        Call FrmImportService.SPC.AAAAA()
        Call FrmImportService.SPC.ExampleHardwareNeedle(_IDS.HardWare.Needle)
    End Sub

#End Region

#Region " DLL Execution "
    Private Sub BtnAddExecutionPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddExecutionPanel.Click

        FrmImportExe.PanelMotor.Top = 50
        FrmImportExe.PanelMotor.Left = 0
        PanelImport.Controls.Add(FrmImportExe.PanelMotor)
        FrmImportExe.PanelMotor.BringToFront()
        FrmImportExe.PanelMotor.Show()
    End Sub

#End Region

#Region " DLL Vision "
    Private Sub BtnAddVisionPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddVisionPanel.Click

        FrmImportVision.PanelVision.Top = 0
        FrmImportVision.PanelVision.Left = 0
        PanelVision.Controls.Add(FrmImportVision.PanelVision)
        FrmImportVision.PanelVision.BringToFront()
        FrmImportVision.PanelVision.Show()

    End Sub

#End Region

#Region " DLL Setup and Caliberation "

    Private Sub InitialSetup()
        FormMyExport.PanelGeneral.Top = 0
        FormMyExport.PanelGeneral.Left = 0
        PanelGeneral.Controls.Add(FormMyExport.PanelGeneral)
    End Sub
#End Region

#Region "Devices "
    Private Sub ConveyorChangemode()

        '''''''''''''''''''''''''''''''''''''''''''
        ' change conveyor mode
        ''''''''''''''''''''''''''''''''''''''''''''
        Call FrmImportConveyor.IDS_ConveyorSetMode(1)

    End Sub
#End Region

#End Region

End Class
